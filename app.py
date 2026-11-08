import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime, timedelta, date, timezone
import tempfile
import os
import re
import json
import pdfplumber
from collections import defaultdict
import traceback
from urllib.parse import quote

# ============ 文件上传 + session_state 内存缓存 ============
class _CachedUploadFile:
    """伪装成 UploadedFile，重跑后从 session_state 内存恢复"""
    def __init__(self, name, data):
        self.name = name
        self.size = len(data)
        self.type = "application/octet-stream"
        self._data = data
        self._pos = 0

    def getvalue(self):
        return self._data

    def read(self, size=-1):
        if size == -1:
            result = self._data[self._pos:]
            self._pos = len(self._data)
        else:
            result = self._data[self._pos:self._pos + size]
            self._pos += size
        return result

    def seek(self, pos, whence=0):
        if whence == 0:
            self._pos = pos
        elif whence == 1:
            self._pos += pos
        elif whence == 2:
            self._pos = len(self._data) + pos
        return self._pos

    def tell(self):
        return self._pos


def sess_file_uploader(label, type=None, key=None, **kwargs):
    """文件上传 + session_state 内存缓存"""
    uploader_key = key or label.replace(" ", "_").replace("/", "_")[:80]

    uploaded = st.file_uploader(label, type=type, key=key, **kwargs)

    bytes_key = f"__uf_bytes__{uploader_key}"
    name_key = f"__uf_name__{uploader_key}"
    id_key = f"__uf_id__{uploader_key}"

    if uploaded is not None:
        file_id = f"{uploaded.name}__{uploaded.size}"
        if st.session_state.get(id_key) != file_id:
            st.session_state[id_key] = file_id
            st.session_state[bytes_key] = uploaded.getvalue()
            st.session_state[name_key] = uploaded.name
        return uploaded

    if bytes_key in st.session_state:
        return _CachedUploadFile(
            st.session_state.get(name_key, "cached_file.xlsx"),
            st.session_state[bytes_key],
        )

    return None
# =========================================================

# ==============================
# 通用工具函数（供各功能使用）
# ==============================
def parse_duration(dur_str) -> int:
    if pd.isna(dur_str):
        return 0
    s = str(dur_str).strip().replace('：', ':')
    s_lower = s.lower()
    days_pattern = re.compile(r'(\d+)\s*days?\s*,?\s*(\d+):(\d{2})(?::(\d{2}))?', re.IGNORECASE)
    match = days_pattern.search(s_lower)
    if match:
        days = int(match.group(1))
        hours = int(match.group(2))
        minutes = int(match.group(3))
        return days * 24 * 60 + hours * 60 + minutes
    time_pattern = re.compile(r'(\d+):(\d{2})(?::(\d{2}))?')
    match = time_pattern.search(s)
    if match:
        hours = int(match.group(1))
        minutes = int(match.group(2))
        return hours * 60 + minutes
    try:
        return int(float(s))
    except:
        return 0

def format_duration(total_minutes: int) -> str:
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}:00"

def detect_header_row(df_raw, keywords):
    best_score = -1
    best_row = 0
    for i in range(min(10, len(df_raw))):
        row_values = [str(v).strip() for v in df_raw.iloc[i].values]
        score = sum(1 for kw in keywords if any(kw in val for val in row_values))
        if score > best_score:
            best_score = score
            best_row = i
    if best_score > 0:
        return best_row, True
    else:
        return 0, False

def read_excel_with_auto_header(file, keywords):
    raw = pd.read_excel(file, header=None, dtype=str)
    header_idx, found = detect_header_row(raw, keywords)
    if not found:
        st.warning("未检测到表头，将使用第一行作为列名，可能导致错误。")
        return pd.read_excel(file, header=0)
    df = pd.read_excel(file, header=header_idx)
    df = df.dropna(axis=1, how='all')
    df = df.loc[:, ~df.columns.str.contains('Unnamed', case=False)]
    df = df.reset_index(drop=True)
    return df

def auto_match_column(df, candidates):
    cols_lower = {col.strip().lower(): col for col in df.columns}
    for cand in candidates:
        cand_lower = cand.strip().lower()
        if cand_lower in cols_lower:
            return cols_lower[cand_lower]
    for cand in candidates:
        for col in df.columns:
            if cand in col:
                return col
    return None

def format_time(value):
    if pd.isna(value) or value == "" or value is None:
        return ""
    if isinstance(value, str):
        if ":" in value:
            parts = value.split(":")
            if len(parts) == 2:
                return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:00"
            elif len(parts) == 3:
                return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:{parts[2].zfill(2)}"
        return value
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M:%S")
    return str(value)

# ==============================
# 功能 A：每日飞行数据自动更新（仅修改第2行日期和第5行数据，不改第3行）
# ==============================
from openpyxl.cell.rich_text import CellRichText, TextBlock


def update_cell_date_only(cell, new_date_str):
    """只替换单元格中的日期部分（如 9月7日 → 9月8日），保留其他文本和样式"""
    val = cell.value
    if val is None:
        return

    if isinstance(val, CellRichText):
        new_blocks = []
        for block in val:
            if isinstance(block, TextBlock):
                new_text = re.sub(r'\d+月\d+日', new_date_str, block.text)
                if new_text != block.text:
                    new_blocks.append(TextBlock(block.font, new_text))
                else:
                    new_blocks.append(block)
            else:
                new_text = re.sub(r'\d+月\d+日', new_date_str, str(block))
                new_blocks.append(new_text)
        cell.value = CellRichText(new_blocks)
    else:
        new_val = re.sub(r'\d+月\d+日', new_date_str, str(val))
        cell.value = new_val


def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col, purpose_col, future_df=None):
    wb = load_workbook(excel1_path, rich_text=True)
    ws = wb.active

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    before_yesterday = today - timedelta(days=2)

    # ---- 日期字符串 ----
    today_str = f"{today.month}月{today.day}日"           # 今日（用于 A1、文件名）
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"  # 昨日（用于 H2/I2 数据日期）

    # ---- 0. 修改 A1 标题中的日期为【今日】 ----
    update_cell_date_only(ws.cell(row=1, column=1), today_str)

    # ---- 1. 修改 H2、I2 中的日期为【昨日】（数据日期） ----
    update_cell_date_only(ws.cell(row=2, column=8), yesterday_str)
    update_cell_date_only(ws.cell(row=2, column=9), yesterday_str)

    # ---- 2. 筛选有效航段 ----
    valid_mask = excel2_df[flight_col].notna()
    valid_df = excel2_df[valid_mask].copy()

    total_minutes = 0
    for val in valid_df[flight_col]:
        total_minutes += parse_duration(val)

    total_hours = total_minutes / 60.0
    total_hours_str = f"{total_hours:.2f}"
    total_flights = len(valid_df)

    # ---- 3. 写入 H5 和 I5 ----
    ws.cell(row=5, column=8).value = total_hours_str
    ws.cell(row=5, column=9).value = total_flights

    # ---- 4. J5 累计时间 ----
    old_j5 = ws.cell(row=5, column=10).value
    if old_j5 is None or old_j5 == "":
        old_j5_float = 0.0
    else:
        try:
            old_j5_float = float(old_j5)
        except:
            old_j5_float = 0.0
    new_j5_float = old_j5_float + total_hours
    ws.cell(row=5, column=10).value = f"{new_j5_float:.2f}"

    # ---- 5. K5 累计架次 ----
    old_k5 = ws.cell(row=5, column=11).value
    if old_k5 is None or old_k5 == "":
        old_k5 = 0
    else:
        try:
            old_k5 = int(old_k5)
        except:
            old_k5 = 0
    new_k5 = old_k5 + total_flights
    ws.cell(row=5, column=11).value = new_k5

    # ---- 6. M5：昨日运行地点（去掉 /通航运输 后缀） ----
    m5_locations = set()
    for _, row in valid_df.iterrows():
        dep = str(row[dep_col]).strip() if pd.notna(row[dep_col]) else ''
        arr = str(row[arr_col]).strip() if pd.notna(row[arr_col]) else ''
        if dep:
            m5_locations.add(dep)
        if arr:
            m5_locations.add(arr)
    m5_list = sorted(m5_locations)
    m5_value = '、'.join(m5_list) if m5_list else ''
    ws.cell(row=5, column=13).value = m5_value

    # ---- 7. N5：使用航空器数量 ----
    reg_series_valid = valid_df[reg_col].dropna()
    unique_regs = reg_series_valid.astype(str).unique()
    ws.cell(row=5, column=14).value = len(unique_regs)

    # ---- 8. S5：次日（明日）计划运行地点（保留 /通航运输 后缀） ----
    future_display = ""
    if future_df is not None and not future_df.empty:
        future_locations = set()
        for _, row in future_df.iterrows():
            dep = str(row[dep_col]).strip() if pd.notna(row[dep_col]) else ''
            arr = str(row[arr_col]).strip() if pd.notna(row[arr_col]) else ''
            if dep:
                future_locations.add(dep)
            if arr:
                future_locations.add(arr)
        future_list = sorted(future_locations)
        s5_value = '、'.join(future_list) + '/通航运输' if future_list else '/通航运输'
        ws.cell(row=5, column=19).value = s5_value
        future_display = '、'.join(future_list) if future_list else ""
    else:
        ws.cell(row=5, column=19).value = ""

    # ========== 统计信息 ==========
    stats = {
        '昨日飞行时间': total_hours_str,
        '架次': total_flights,
        '注册号数量': len(unique_regs),
        f'截止{before_yesterday.month}月{before_yesterday.day}日总飞行时间': f"{old_j5_float:.2f}",
        f'截止{yesterday.month}月{yesterday.day}日总飞行时间': f"{new_j5_float:.2f}",
        f'截止{before_yesterday.month}月{before_yesterday.day}日总飞行架次': old_k5,
        f'截止{yesterday.month}月{yesterday.day}日总飞行架次': new_k5,
        '昨日运行地点': '、'.join(m5_list) if m5_list else "",
        '次日运行地点': future_display,
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats


def run_feature_a():
    excel1_file = sess_file_uploader("📂 上传：昨日飞行计划日报（operation-每日飞行数据）", type=["xlsx", "xlsm"], key="a_excel1")
    excel2_file = sess_file_uploader("📂 上传：航段数据导出（昨日B机）", type=["xlsx", "xlsm"], key="a_excel2")
    excel3_file = sess_file_uploader(
        "📂 上传：航段数据导出（今天的次日，比如今日12号，上传13号的航段数据）（用于填入次日飞行计划运行地点）",
        type=["xlsx", "xlsm"],
        key="a_excel3"
    )

    if excel1_file and excel2_file:
        with st.spinner("正在自动处理..."):
            try:
                keywords = ["客户", "航班号", "出发城市", "到达城市", "实际飞行时间", "飞机注册号", "用途"]
                excel2_df = read_excel_with_auto_header(excel2_file, keywords)
                if excel2_df.empty or len(excel2_df.columns) == 0:
                    st.error("航段数据没有有效列，请检查文件格式。")
                    return

                flight_col = auto_match_column(excel2_df, ["实际飞行时间", "飞行时间", "航段时间"])
                dep_col = auto_match_column(excel2_df, ["出发城市", "起飞机场", "出发地"])
                arr_col = auto_match_column(excel2_df, ["到达城市", "目的地机场", "到达地"])
                reg_col = auto_match_column(excel2_df, ["飞机注册号", "注册号", "机号"])
                purpose_col = auto_match_column(excel2_df, ["用途", "任务性质", "飞行任务"])

                missing = []
                if not flight_col: missing.append("飞行时间")
                if not dep_col: missing.append("出发城市")
                if not arr_col: missing.append("到达城市")
                if not reg_col: missing.append("飞机注册号")
                if not purpose_col: missing.append("用途")
                if missing:
                    st.error(f"未能自动匹配以下列：{', '.join(missing)}，请检查文件列名是否包含关键词。")
                    return

                # 读取次日计划数据（可选）
                future_df = None
                if excel3_file is not None:
                    future_df = read_excel_with_auto_header(excel3_file, keywords)
                    if future_df.empty:
                        st.warning("次日计划文件为空或格式不正确，S5列将不填充。")
                        future_df = None

                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                output_path, stats = update_excel1(
                    excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col, purpose_col,
                    future_df=future_df
                )

                st.success("✅ 处理完成")

                # ---- 显示核心指标 ----
                col1, col2, col3 = st.columns(3)
                col1.metric("昨日总飞行时间", stats['昨日飞行时间'])
                col2.metric("架次", stats['架次'])
                col3.metric("使用航空器数量", stats['注册号数量'])

                # 截止时间指标
                time_labels = [k for k in stats.keys() if k.startswith('截止') and '总飞行时间' in k]
                flights_labels = [k for k in stats.keys() if k.startswith('截止') and '总飞行架次' in k]

                col4, col5 = st.columns(2)
                col4.metric(time_labels[0], stats[time_labels[0]])
                col5.metric(time_labels[1], stats[time_labels[1]])

                col6, col7 = st.columns(2)
                col6.metric(flights_labels[0], stats[flights_labels[0]])
                col7.metric(flights_labels[1], stats[flights_labels[1]])

                # ---- 显示运行地点 ----
                st.subheader("📍 运行地点")
                if stats['昨日运行地点']:
                    st.write(f"**昨日运行地点：** {stats['昨日运行地点']}")
                else:
                    st.write("**昨日运行地点：** 无")
                if stats['次日运行地点']:
                    st.write(f"**次日运行地点：** {stats['次日运行地点']}")
                else:
                    st.write("**次日运行地点：** （未上传次日计划文件）")

                # ===== 文件名使用今日日期 =====
                today = datetime.now().date()
                file_name = f"中南-深圳局-天成商务航空有限公司-{today.month}月{today.day}日飞行计划日报.xlsx"

                with open(output_path, 'rb') as f:
                    st.download_button(
                        label="📥 下载更新后的文件",
                        data=f.read(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                os.unlink(excel1_path)
                os.unlink(output_path)

            except Exception as e:
                st.error(f"处理出错：{e}")
                if 'excel1_path' in locals() and os.path.exists(excel1_path):
                    os.unlink(excel1_path)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.unlink(output_path)


# ==============================
# 功能 B：模板生成备案表
# ==============================
DEFAULT_ICAO_MAP = {
    "B65AP": "GLF4",
    "B652R": "GLF4",
    "B652S": "GLF4",
    "B8105": "GLEX",
    "B8309": "GLF5",
    "B652Q": "GLF4",
    "B3926": "LJ60",
    "B8160": "GLF5",
    "B8262": "GLF4",
    "B8292": "GLF5",
}

def find_header_row(df_raw):
    keywords = ["客户", "航班号", "飞机注册号", "出发日期"]
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
        if all(kw in row_str for kw in keywords):
            return idx
    return 0

def parse_uploaded_file(uploaded_file):
    df_all = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    header_row = find_header_row(df_all)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    return df

def map_usage(usage):
    usage = str(usage).strip()
    if "调机" in usage:
        return "调机飞行", "调机飞行"
    elif "载客" in usage or "共享租赁" in usage:
        return "公务飞行", "私用飞行"
    else:
        return "公务飞行", "私用飞行"

def parse_duration_to_minutes(duration_str):
    if pd.isna(duration_str) or duration_str == "" or duration_str is None:
        return None
    s = str(duration_str).strip()
    try:
        minutes = parse_duration(s)
        return minutes
    except:
        return None

def combine_date_time(date_val, time_val):
    if pd.isna(date_val) or pd.isna(time_val):
        return None
    try:
        date_str = str(date_val).strip()
        time_str = str(time_val).strip()
        if time_str.count(":") == 2:
            dt_str = f"{date_str} {time_str}"
        else:
            dt_str = f"{date_str} {time_str}:00"
        return pd.to_datetime(dt_str)
    except:
        return None

def run_feature_b():
    default_supervision = "深圳局"
    default_operator = "天成商务航空有限公司"
    icao_map = DEFAULT_ICAO_MAP.copy()

    if "template_wb" not in st.session_state:
        st.session_state.template_wb = None
    if "header_row" not in st.session_state:
        st.session_state.header_row = None
    if "data_start_row" not in st.session_state:
        st.session_state.data_start_row = None
    if "group_rows" not in st.session_state:
        st.session_state.group_rows = 4

    st.subheader("📂 模板管理")
    if st.session_state.template_wb is None:
        st.info("首次使用请上传模板文件。")
        template_file = sess_file_uploader("上传：桌面-Jetops申请一览-每日通航运行情况跟踪表（天成商务航空有限公司）20260708", type=["xlsx"], key="b_template_upload")
        if template_file:
            try:
                wb = load_workbook(template_file)
                ws = wb.active
                header_row = None
                for row in range(1, 5):
                    if any("所属监管局" in str(cell.value) for cell in ws[row]):
                        header_row = row
                        break
                if header_row is None:
                    st.error("模板中未找到表头行，请确认模板有“所属监管局”字段。")
                    st.stop()
                data_start_row = header_row + 1
                st.session_state.template_wb = wb
                st.session_state.header_row = header_row
                st.session_state.data_start_row = data_start_row
                st.success("✅ 模板加载成功！现在可以上传数据文件。")
            except Exception as e:
                st.error(f"模板加载失败：{e}")
    else:
        st.success("✅ 模板已加载（如需更换，请点击下方按钮重置）")
        if st.button("重新上传模板", key="b_reset_template"):
            st.session_state.template_wb = None
            st.session_state.header_row = None
            st.session_state.data_start_row = None
            st.rerun()

    st.subheader("📊 数据上传")
    data_file = sess_file_uploader("上传：航段数据导出（当日B机）", type=["xlsx"], key="b_data_upload")

    if data_file and st.session_state.template_wb is not None:
        try:
            df_raw = parse_uploaded_file(data_file)
            st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")
            if len(df_raw) > 20:
                st.warning(f"数据条数（{len(df_raw)}）超过模板预设的20行，多余数据将被忽略。")

            total_flights = len(df_raw)
            status_col = "航段状态" if "航段状态" in df_raw.columns else None
            actual_depart_col = "实际出发" if "实际出发" in df_raw.columns else None
            reg_col = "飞机注册号" if "飞机注册号" in df_raw.columns else None

            if status_col:
                landed = df_raw[df_raw[status_col].astype(str).str.contains("已执飞|已完成", na=False)]
                not_landed = df_raw[~df_raw[status_col].astype(str).str.contains("已执飞|已完成", na=False)]
            else:
                landed = pd.DataFrame()
                not_landed = df_raw

            if actual_depart_col:
                unlanded = not_landed[not_landed[actual_depart_col].notna()]
                not_executed = not_landed[not_landed[actual_depart_col].isna()]
            else:
                unlanded = pd.DataFrame()
                not_executed = not_landed

            landed_count = len(landed)
            unlanded_count = len(unlanded)
            not_executed_count = len(not_executed)

            if reg_col:
                all_regs = df_raw[reg_col].dropna().astype(str).unique()
                reg_list = sorted(all_regs)
            else:
                reg_list = []

            # ============ 汇报文案 ============
            parts = []
            if landed_count > 0:
                parts.append(f"{landed_count}班已落地")
            if unlanded_count > 0:
                parts.append(f"{unlanded_count}班未落地")
            if not_executed_count > 0:
                parts.append(f"{not_executed_count}班未起飞")

            if not parts:
                report = "下午好，天成今日无飞行计划"
            else:
                report = "下午好，天成今日" + "、".join(parts)

            if reg_list:
                plane_text = f"今日有飞行计划的飞机：{'、'.join(reg_list)}"
            else:
                plane_text = ""

            st.subheader("📊 飞行计划统计")
            col1, col2, col3 = st.columns(3)
            col1.metric("飞行计划总数", total_flights)
            col2.metric("已落地班次", landed_count)
            col3.metric("未落地班次", unlanded_count)
            col4, col5 = st.columns(2)
            col4.metric("未执行班次", not_executed_count)
            col5.metric("涉及飞机数", len(reg_list))
            if reg_list:
                st.write("**飞机注册号列表：**", "、".join(reg_list))

            # 组装汇报文案（中间空一行）
            full_report = report
            if plane_text:
                full_report += "\n\n" + plane_text
            st.text_area("📋 汇报文案（可复制）", full_report, height=140)

            wb = st.session_state.template_wb
            ws = wb.active
            header_row = st.session_state.header_row
            data_start_row = st.session_state.data_start_row
            group_rows = st.session_state.group_rows

            has_actual_depart = "实际出发" in df_raw.columns
            has_estimated_flight_time = "预计飞行时间" in df_raw.columns
            if not has_estimated_flight_time:
                st.warning("数据中没有“预计飞行时间”列，将使用原“预计到达”作为K列值。")

            records = []
            for idx, row in df_raw.iterrows():
                if idx >= 20:
                    break

                dt = pd.to_datetime(row["出发日期"])
                flight_date = f"{dt.year}/{dt.month}/{dt.day}"

                dep_city = str(row.get("出发城市", "")).strip()
                arr_city = str(row.get("到达城市", "")).strip()
                route = f"{dep_city}-{arr_city}" if dep_city and arr_city else f"{row['出发地']}-{row['到达地']}"

                run_type, oper_type = map_usage(row["用途"])

                if has_actual_depart:
                    actual_depart = format_time(row.get("实际出发", ""))
                    if actual_depart:
                        start_time = actual_depart
                    else:
                        start_time = format_time(row.get("计划出发", ""))
                else:
                    start_time = format_time(row.get("计划出发", ""))

                estimated_landing = ""
                if has_estimated_flight_time:
                    duration_str = row.get("预计飞行时间", "")
                    minutes = parse_duration_to_minutes(duration_str)
                    if minutes is not None:
                        date_val = row["出发日期"]
                        if has_actual_depart and pd.notna(row.get("实际出发")):
                            time_val = row["实际出发"]
                        else:
                            time_val = row["计划出发"]
                        if pd.notna(date_val) and pd.notna(time_val):
                            dt_obj = combine_date_time(date_val, time_val)
                            if dt_obj is not None:
                                new_dt = dt_obj + timedelta(minutes=minutes)
                                estimated_landing = new_dt.strftime("%H:%M:%S")
                    if not estimated_landing and "预计到达" in df_raw.columns:
                        estimated_landing = format_time(row.get("预计到达", ""))
                else:
                    if "预计到达" in df_raw.columns:
                        estimated_landing = format_time(row.get("预计到达", ""))

                actual_end = format_time(row.get("实际到达", "")) if "实际到达" in df_raw.columns else ""

                status = str(row.get("航段状态", "")).strip()
                is_landed = "是" if status in ["已执飞", "已完成"] else "否"

                reg = str(row["飞机注册号"]).strip().upper()
                icao_type = icao_map.get(reg, "")

                record = {
                    "A": default_supervision,
                    "B": default_operator,
                    "C": flight_date,
                    "D": run_type,
                    "E": oper_type,
                    "F": icao_type,
                    "G": reg,
                    "J": start_time,
                    "K": estimated_landing,
                    "L": is_landed,
                    "M": actual_end if is_landed == "是" else "",
                    "N": route,
                }
                records.append(record)

            for i, rec in enumerate(records):
                row_num = data_start_row + i * group_rows
                ws[f"A{row_num}"] = rec["A"]
                ws[f"B{row_num}"] = rec["B"]
                ws[f"C{row_num}"] = rec["C"]
                ws[f"D{row_num}"] = rec["D"]
                ws[f"E{row_num}"] = rec["E"]
                ws[f"F{row_num}"] = rec["F"]
                ws[f"G{row_num}"] = rec["G"]
                ws[f"J{row_num}"] = rec["J"]
                ws[f"K{row_num}"] = rec["K"]
                ws[f"L{row_num}"] = rec["L"]
                ws[f"M{row_num}"] = rec["M"]
                ws[f"N{row_num}"] = rec["N"]

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                label="⬇️ 下载填入数据的模板文件",
                data=output,
                file_name="每日通航运行情况跟踪表_生成.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="b_download"
            )

        except Exception as e:
            st.error(f"❌ 处理失败：{e}")
            st.exception(e)
    else:
        if st.session_state.template_wb is None:
            st.info("👆 请先上传模板。")
        else:
            st.info("👆 请上传航段数据 Excel 文件。")


# ==============================
# 功能 C：生成每日运行跟踪表
# ==============================
def run_feature_c():
    ICAO_MAP = {
        "B65AP": "GLF4",
        "B652R": "GLF4",
        "B652S": "GLF4",
        "B8105": "GLEX",
        "B8309": "GLF5",
        "B652Q": "GLF4",
        "B3926": "LJ60",
        "B8160": "GLF5",
        "B8262": "GLF4",
        "B8292": "GLF5",
    }

    st.markdown("上传 **天成商务航空每日运行跟踪** 模板和 **航段数据导出**，自动在模板最下方新增一行今日汇总数据。")

    template_file = sess_file_uploader("📂 上传：Operation-每日通航运行情况-天成商务航空每日运行跟踪", type=["xlsx"], key="c_template")
    data_file = sess_file_uploader("📂 上传：航段数据导出（当日B机）", type=["xlsx"], key="c_data")

    if template_file and data_file:
        with st.spinner("正在处理..."):
            try:
                keywords = ["客户", "航班号", "出发城市", "到达城市", "实际出发", "计划出发", "预计到达", "实际到达", "航段状态", "飞机注册号", "用途"]
                df = read_excel_with_auto_header(data_file, keywords)
                if df.empty:
                    st.error("航段数据为空或格式不正确。")
                    return

                required_cols = ["飞机注册号", "用途", "出发城市", "到达城市", "航段状态"]
                for col in required_cols:
                    if col not in df.columns:
                        st.error(f"航段数据缺少必要列：{col}")
                        return

                today = datetime.now().date()
                today_str = today.strftime("%Y/%m/%d")

                total_flights = len(df)
                actual_flights = total_flights

                usage_set = set()
                for u in df["用途"].dropna():
                    u_str = str(u).strip()
                    if "调机" in u_str:
                        usage_set.add("调机飞行")
                    else:
                        usage_set.add("公务飞行")
                run_types = "、".join(sorted(usage_set)) if usage_set else "公务飞行"

                status_col = "航段状态"
                has_actual_depart = "实际出发" in df.columns
                all_landed = all(str(s).strip() in ["已执飞", "已完成"] for s in df[status_col].dropna())
                if has_actual_depart:
                    any_started = any(pd.notna(s) for s in df["实际出发"])
                else:
                    any_started = any(str(s).strip() in ["已执飞", "已完成", "执飞中"] for s in df[status_col].dropna())

                if all_landed:
                    status_text = "已实施，全部结束"
                elif any_started:
                    status_text = "已实施，未全部结束"
                else:
                    status_text = "未实施"

                # ============ 开始时间 ============
                if has_actual_depart:
                    valid_depart = df[df["实际出发"].notna()]
                    if not valid_depart.empty:
                        times = valid_depart["实际出发"].apply(lambda x: str(x).strip())
                        earliest = min(times, key=lambda t: t if t else "99:99")
                        start_time = format_time(earliest)
                    else:
                        if "计划出发" in df.columns:
                            plan_times = df["计划出发"].dropna().apply(lambda x: str(x).strip())
                            if not plan_times.empty:
                                earliest_plan = min(plan_times, key=lambda t: t if t else "99:99")
                                start_time = format_time(earliest_plan)
                            else:
                                start_time = ""
                        else:
                            start_time = ""
                else:
                    if "计划出发" in df.columns:
                        plan_times = df["计划出发"].dropna().apply(lambda x: str(x).strip())
                        if not plan_times.empty:
                            earliest_plan = min(plan_times, key=lambda t: t if t else "99:99")
                            start_time = format_time(earliest_plan)
                        else:
                            start_time = ""
                    else:
                        start_time = ""

                # ============ 计划结束时间：取所有航段中最晚的结束时刻（考虑跨天） ============
                plan_end = ""
                if "预计到达" in df.columns:
                    end_datetimes = []
                    for _, row in df.iterrows():
                        arr_time_val = row.get("预计到达")
                        if pd.isna(arr_time_val) or str(arr_time_val).strip() == "":
                            continue
                        # 优先用"到达日期"，缺则回退"出发日期"
                        arr_date_val = None
                        if "到达日期" in df.columns:
                            arr_date_val = row.get("到达日期")
                        if arr_date_val is None or pd.isna(arr_date_val) or str(arr_date_val).strip() == "":
                            arr_date_val = row.get("出发日期")
                        if pd.isna(arr_date_val) or str(arr_date_val).strip() == "":
                            continue
                        try:
                            arr_date = pd.to_datetime(arr_date_val)
                            time_str = format_time(arr_time_val)
                            parts_t = time_str.split(":")
                            h = int(parts_t[0])
                            m = int(parts_t[1])
                            s = int(parts_t[2]) if len(parts_t) > 2 else 0
                            end_dt = arr_date.replace(hour=h, minute=m, second=s)
                            end_datetimes.append(end_dt)
                        except Exception:
                            continue
                    if end_datetimes:
                        latest_end = max(end_datetimes)
                        plan_end = latest_end.strftime("%H:%M:%S")

                # ============ 实际结束时间（只在全部落地时填） ============
                if all_landed and "实际到达" in df.columns:
                    actual_end_datetimes = []
                    for _, row in df.iterrows():
                        arr_time_val = row.get("实际到达")
                        if pd.isna(arr_time_val) or str(arr_time_val).strip() == "":
                            continue
                        arr_date_val = None
                        if "到达日期" in df.columns:
                            arr_date_val = row.get("到达日期")
                        if arr_date_val is None or pd.isna(arr_date_val) or str(arr_date_val).strip() == "":
                            arr_date_val = row.get("出发日期")
                        if pd.isna(arr_date_val) or str(arr_date_val).strip() == "":
                            continue
                        try:
                            arr_date = pd.to_datetime(arr_date_val)
                            time_str = format_time(arr_time_val)
                            parts_t = time_str.split(":")
                            h = int(parts_t[0])
                            m = int(parts_t[1])
                            s = int(parts_t[2]) if len(parts_t) > 2 else 0
                            end_dt = arr_date.replace(hour=h, minute=m, second=s)
                            actual_end_datetimes.append(end_dt)
                        except Exception:
                            continue
                    if actual_end_datetimes:
                        latest_actual = max(actual_end_datetimes)
                        actual_end = latest_actual.strftime("%H:%M:%S")
                    else:
                        actual_end = ""
                else:
                    actual_end = ""

                regs = df["飞机注册号"].dropna().astype(str).str.upper().unique()
                models = set()
                for reg in regs:
                    model = ICAO_MAP.get(reg, "")
                    if model:
                        models.add(model)
                model_str = "、".join(sorted(models)) if models else ""

                routes = []
                for _, row in df.iterrows():
                    dep = str(row.get("出发城市", "")).strip()
                    arr = str(row.get("到达城市", "")).strip()
                    if dep and arr:
                        routes.append(f"{dep}-{arr}")
                    elif dep:
                        routes.append(dep)
                    elif arr:
                        routes.append(arr)
                route_str = "、".join(routes)
                if len(route_str) > 100:
                    route_str = route_str[:100] + "..."

                supervision = "深圳局"
                category = "公务航空飞行"
                operator = "天成商务航空有限公司"
                yes = "是"

                wb = load_workbook(template_file)
                ws = wb.active

                target_row = None
                for row in range(2, ws.max_row + 2):
                    if ws.cell(row, 1).value is None or str(ws.cell(row, 1).value).strip() == "":
                        target_row = row
                        break
                if target_row is None:
                    target_row = ws.max_row + 1

                ws.cell(target_row, 1).value = supervision
                ws.cell(target_row, 2).value = today_str
                ws.cell(target_row, 3).value = category
                ws.cell(target_row, 4).value = operator
                ws.cell(target_row, 5).value = run_types
                ws.cell(target_row, 6).value = total_flights
                ws.cell(target_row, 7).value = actual_flights
                ws.cell(target_row, 8).value = status_text
                ws.cell(target_row, 9).value = start_time
                ws.cell(target_row, 10).value = plan_end
                ws.cell(target_row, 11).value = actual_end
                ws.cell(target_row, 12).value = model_str
                ws.cell(target_row, 13).value = route_str
                ws.cell(target_row, 15).value = yes
                ws.cell(target_row, 16).value = yes
                ws.cell(target_row, 17).value = yes

                font10 = Font(size=10)
                for col in range(1, 18):
                    ws.cell(row=target_row, column=col).font = font10

                align_right = Alignment(horizontal='right', vertical='center')
                for col in [2, 9, 10, 11]:
                    ws.cell(row=target_row, column=col).alignment = align_right

                align_left_no_wrap = Alignment(horizontal='left', vertical='center', wrap_text=False, shrink_to_fit=False)
                ws.cell(row=target_row, column=13).alignment = align_left_no_wrap
                ws.column_dimensions['M'].width = 50

                prev_row = target_row - 1
                if prev_row >= 1:
                    src_f_num = ws.cell(prev_row, 6).number_format
                    if src_f_num:
                        ws.cell(target_row, 6).number_format = src_f_num
                    else:
                        ws.cell(target_row, 6).number_format = '0'
                    src_g_num = ws.cell(prev_row, 7).number_format
                    if src_g_num:
                        ws.cell(target_row, 7).number_format = src_g_num
                    else:
                        ws.cell(target_row, 7).number_format = '0'
                else:
                    ws.cell(target_row, 6).number_format = '0'
                    ws.cell(target_row, 7).number_format = '0'

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                st.success("✅ 处理完成！已添加一行新数据。")
                col1, col2, col3 = st.columns(3)
                col1.metric("总架次", total_flights)
                col2.metric("状态", status_text)
                col3.metric("航空器型号数", len(models))

                st.write(f"**开始时间：** {start_time}")
                st.write(f"**计划结束：** {plan_end}")
                st.write(f"**实际结束：** {actual_end if actual_end else '未结束'}")
                st.write(f"**航线：** {route_str}")

                st.download_button(
                    label="⬇️ 下载更新后的模板",
                    data=output,
                    file_name="天成商务航空每日运行跟踪.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"处理失败：{e}")
                st.exception(e)









# ==============================
# 功能 D：通航脚本生成器（完整版，包含所有脚本生成函数）
# ==============================
def run_feature_d():
    st.markdown("上传 Excel 文件，自动生成浏览器控制台脚本，**先执行当日已飞完计划，再备案未来及漏备案计划**。")

    header_row = 1

    uploaded_file = sess_file_uploader("📂 上传：航段数据导出（只能是B机，提取的时候注意）", type=["xlsx", "xls"], key="d_upload")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=header_row)
            df.columns = df.columns.str.strip()
            df = df.dropna(how='all')
            st.success("文件上传成功！")

            required_cols = ["飞机注册号", "出发日期", "到达日期", "用途", "出发城市", "到达城市", "预计飞行时间", "实际到达"]
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                st.error(f"❌ 缺少必要列: {missing}")
                st.info(f"实际列名: {list(df.columns)}")
                return

            # 处理日期格式
            df['出发日期'] = pd.to_datetime(df['出发日期']).dt.date
            today = date.today()

            # 数据拆分
            # 1. 今日已飞完
            df_today_landed = df[(df['出发日期'] == today) & (df["实际到达"].notna() & (df["实际到达"].astype(str).str.strip() != ""))].copy()
            # 2. 今日未飞完
            df_today_unlanded = df[(df['出发日期'] == today) & (df["实际到达"].isna() | (df["实际到达"].astype(str).str.strip() == ""))].copy()
            # 3. 未来计划
            df_future = df[df['出发日期'] > today].copy()

            st.info(f"✅ 共读取 {len(df)} 条计划。今日已飞完: {len(df_today_landed)} 条，今日未飞完: {len(df_today_unlanded)} 条，未来计划: {len(df_future)} 条。")

            if len(df_today_landed) == 0 and len(df_today_unlanded) == 0 and len(df_future) == 0:
                st.warning("没有需要处理的计划。")
                return

            def get_daily_rec(row):
                rec_dict = row.to_dict()
                for k, v in rec_dict.items():
                    if pd.isna(v):
                        rec_dict[k] = ""
                return rec_dict

            def get_future_rec(row):
                purpose_raw = row.get("用途", "")
                if "维修" in purpose_raw or "调机" in purpose_raw:
                    purpose = "调机"
                else:
                    purpose = "自用飞行"
                start_date = str(row["出发日期"])
                end_date = str(row["到达日期"])
                
                # 内联解析时间，不依赖外部函数
                flight_time = str(row.get("预计飞行时间", "")).strip()
                hours, minutes = 0, 0
                if flight_time:
                    parts = flight_time.split(':')
                    if len(parts) >= 2:
                        try:
                            hours = int(float(parts[0]))
                            minutes = int(float(parts[1]))
                        except:
                            pass

                dep_city = str(row["出发城市"]).strip()
                arr_city = str(row["到达城市"]).strip()
                reg_raw = str(row["飞机注册号"]).strip()
                return {
                    "reg": reg_raw,
                    "start_date": start_date,
                    "end_date": end_date,
                    "purpose": purpose,
                    "dep_city": dep_city,
                    "arr_city": arr_city,
                    "flight_hours": hours,
                    "flight_minutes": minutes
                }

            col1, col2 = st.columns(2)
            
            filtered_daily_records = []
            filtered_future_records = []

            with col1:
                st.markdown("### 📋 执行的计划（今日已飞完）")
                if len(df_today_landed) > 0:
                    select_all_exec = st.checkbox("全选执行计划", value=True, key="select_all_exec")
                    for idx, row in df_today_landed.iterrows():
                        label = f"{row['飞机注册号']} {row['出发日期']} {row['出发城市']}->{row['到达城市']}"
                        
                        c1, c2 = st.columns([4, 1])
                        with c1:
                            do_exec = st.checkbox(label, value=select_all_exec, key=f"exec_{idx}")
                        with c2:
                            do_backup = st.checkbox("补备案", value=False, key=f"backup_{idx}", help="如果昨天漏备案，可勾选此项")
                        
                        if do_exec:
                            filtered_daily_records.append(get_daily_rec(row))
                        if do_backup:
                            filtered_future_records.append(get_future_rec(row))
                else:
                    st.info("暂无今日已飞完的计划")

            with col2:
                st.markdown("### 📝 备案的计划")
                
                # 今日未飞完
                if len(df_today_unlanded) > 0:
                    st.markdown("**今日未飞完（如昨天漏备案，请勾选补备案）**")
                    for idx, row in df_today_unlanded.iterrows():
                        label = f"{row['飞机注册号']} {row['出发日期']} {row['出发城市']}->{row['到达城市']}"
                        do_backup_today = st.checkbox(label, value=False, key=f"file_unlanded_{idx}")
                        if do_backup_today:
                            filtered_future_records.append(get_future_rec(row))
                
                # 未来计划
                if len(df_future) > 0:
                    st.markdown("**未来计划（默认全选备案）**")
                    select_all_file = st.checkbox("全选未来计划", value=True, key="select_all_file")
                    for idx, row in df_future.iterrows():
                        label = f"{row['飞机注册号']} {row['出发日期']} {row['出发城市']}->{row['到达城市']}"
                        do_file = st.checkbox(label, value=select_all_file, key=f"file_future_{idx}")
                        if do_file:
                            filtered_future_records.append(get_future_rec(row))

                if len(df_today_unlanded) == 0 and len(df_future) == 0:
                    st.info("暂无备案计划")

            st.success(f"✅ 最终脚本将包含 {len(filtered_daily_records)} 条执行计划，{len(filtered_future_records)} 条备案计划")

            custom_detail_map = {}
            city_map, detail_map = build_city_mappings(df, custom_detail_map)
            city_map_json = json.dumps(city_map, ensure_ascii=False, indent=4)
            detail_map_json = json.dumps(detail_map, ensure_ascii=False, indent=4)
            domestic_keywords_json = json.dumps(DOMESTIC_KEYWORDS)

            base_script = generate_base_script(city_map_json, detail_map_json, domestic_keywords_json)
            daily_script = generate_daily_script(filtered_daily_records, city_map_json, detail_map_json, domestic_keywords_json) if len(filtered_daily_records) > 0 else ""
            nextday_script = generate_nextday_script(filtered_future_records, city_map_json, detail_map_json, domestic_keywords_json) if len(filtered_future_records) > 0 else ""

            final_script = base_script + "\n\n" + daily_script + "\n\n" + nextday_script + """
(async () => {
    console.log("========== 开始执行综合流程 ==========");
    if (typeof runDailyPlans === 'function') {
        await runDailyPlans();
    } else {
        console.log("没有当日执行计划需要处理。");
    }
    if (typeof runNextDayPlans === 'function') {
        await runNextDayPlans();
    } else {
        console.log("没有备案计划需要处理。");
    }
    console.log("========== 综合流程全部完成 ==========");
})();
"""
            st.subheader("📋 复制以下代码到浏览器控制台（F12）运行")
            st.code(final_script, language="javascript")
            st.info("💡 提示：请确保已登录系统并停留在「经营活动信息管理」列表页，脚本将自动处理当日执行计划和备案计划。")
            st.download_button(
                label="💾 下载脚本文件 (.js)",
                data=final_script,
                file_name="flight_plan_combined.js",
                mime="application/javascript"
            )
        except Exception as e:
            st.error(f"处理文件时出错: {e}")
            import traceback
            st.exception(e)
    else:
        st.info("请上传 Excel 文件开始")












# ==============================
# 功能 E：值班连班统计
# ==============================
def run_feature_e():
    st.markdown("上传值班表（PDF或Excel），自动统计运管主班、运控白班/夜班、补贴天数、休息天数和在岗天数。")

    uploaded_file = sess_file_uploader("上传值班表（PDF或Excel）", type=["pdf", "xlsx", "xls"], key="e_upload")

    control_staff_input = st.text_input(
        "值班人员名单（空格分隔）",
        value="陈宇鸣 周贤民 吴迪 王浩宇 林泓辰 陈育盛 钟洪达",
        key="e_control"
    )

    if uploaded_file:
        control_staff = set(control_staff_input.strip().split())
        target_staff = control_staff

        schedules = []
        file_type = uploaded_file.type

        if file_type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]:
            try:
                df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            except Exception as e:
                st.error(f"读取Excel失败: {e}")
                st.stop()

            st.subheader("原始表格预览（前20行）")
            st.dataframe(df.head(20))

            header_idx = None
            for i, row in df.iterrows():
                row_str = " ".join([str(c) for c in row if pd.notna(c)])
                if "运行管理" in row_str:
                    header_idx = i
                    break

            if header_idx is None:
                st.error("未找到包含'运行管理'的表头行")
                st.stop()

            col_mapping = {}
            for idx, val in enumerate(df.iloc[header_idx]):
                val_str = str(val) if pd.notna(val) else ""
                if "运行管理" in val_str:
                    col_mapping[idx] = "management"
                elif "运行计划" in val_str:
                    col_mapping[idx] = "plan"
                elif "运行监控" in val_str:
                    col_mapping[idx] = "control"
                elif "运行保障" in val_str:
                    col_mapping[idx] = "control"
                elif "运行支援" in val_str:
                    col_mapping[idx] = "control"

            data_start = header_idx + 1
            current_date = None
            day_row = None

            for i in range(data_start, len(df)):
                row = df.iloc[i]
                second_cell = str(row[1]) if pd.notna(row[1]) else ""
                second_cell = second_cell.strip()

                if "白" in second_cell:
                    first_cell = str(row[0]) if pd.notna(row[0]) else ""
                    date_match = re.search(r"(\d+月\d+日|\d+日)", first_cell)
                    if date_match:
                        current_date = date_match.group(1)
                    else:
                        current_date = first_cell
                    day_row = row
                elif "晚" in second_cell and day_row is not None:
                    if current_date is None:
                        first_cell = str(row[0]) if pd.notna(row[0]) else ""
                        date_match = re.search(r"(\d+月\d+日|\d+日)", first_cell)
                        if date_match:
                            current_date = date_match.group(1)
                    if current_date:
                        day_people = set()
                        night_people = set()
                        for col_idx, role in col_mapping.items():
                            day_name = str(day_row[col_idx]).strip() if pd.notna(day_row[col_idx]) else ""
                            night_name = str(row[col_idx]).strip() if pd.notna(row[col_idx]) else ""
                            if day_name and day_name not in ["nan", "None", ""]:
                                if day_name in target_staff:
                                    day_people.add(day_name)
                            if night_name and night_name not in ["nan", "None", ""]:
                                if night_name in target_staff:
                                    night_people.add(night_name)
                        schedules.append({"date": current_date, "day": day_people, "night": night_people})
                        day_row = None
                        current_date = None

            if not schedules:
                st.error("未能从Excel解析到排班数据，请检查格式")
                st.stop()

        else:
            with pdfplumber.open(uploaded_file) as pdf:
                all_text = ""
                for page in pdf.pages:
                    all_text += page.extract_text() + "\n"

            lines = all_text.split("\n")
            day_shifts = []
            night_shifts = []

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if "白" in line and "晚" not in line:
                    date_match = re.search(r"(\d+月\d+日|\d+日)", line)
                    date_str = date_match.group(1) if date_match else ""
                    names = re.findall(r"[\u4e00-\u9fa5]{2,3}", line)
                    filtered_names = [n for n in names if n in target_staff and n not in ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "运行控制", "运行管理", "白班", "夜班", "带班主任"]]
                    if filtered_names:
                        day_shifts.append((date_str, filtered_names))
                elif "晚" in line:
                    date_match = re.search(r"(\d+月\d+日|\d+日)", line)
                    date_str = date_match.group(1) if date_match else ""
                    names = re.findall(r"[\u4e00-\u9fa5]{2,3}", line)
                    filtered_names = [n for n in names if n in target_staff and n not in ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "运行控制", "运行管理", "白班", "夜班", "带班主任"]]
                    if filtered_names:
                        night_shifts.append((date_str, filtered_names))

            min_len = min(len(day_shifts), len(night_shifts))
            for i in range(min_len):
                date_day, day_names = day_shifts[i]
                date_night, night_names = night_shifts[i]
                date_str = date_day if date_day else date_night
                schedules.append({
                    "date": date_str,
                    "day": set(day_names),
                    "night": set(night_names)
                })

            if not schedules:
                st.error("未识别到任何排班数据，请检查文件格式")
                st.stop()

        st.success(f"成功识别 {len(schedules)} 天的排班数据")

        import re as _re
        def parse_date_to_tuple(date_str):
            match = _re.search(r'(\d+)月(\d+)日', date_str)
            if match:
                return (int(match.group(1)), int(match.group(2)))
            match = _re.search(r'(\d+)日', date_str)
            if match:
                return (9, int(match.group(1)))
            return (9, 99)

        schedules.sort(key=lambda x: parse_date_to_tuple(x['date']))

        with st.expander("🔍 完整排班数据（请核对）"):
            st.write(f"**总天数：{len(schedules)} 天**")
            for sch in schedules:
                st.write(f"**{sch['date']}**")
                st.write(f"  - 白班：{', '.join(sorted(sch['day'])) if sch['day'] else '（无）'}")
                st.write(f"  - 夜班：{', '.join(sorted(sch['night'])) if sch['night'] else '（无）'}")

        all_persons = target_staff
        stats = {name: {"consecutive": 0, "pure_day": 0, "pure_night": 0, "total_night": 0, "rest_days": 0} for name in all_persons}
        attendance_records = {name: [] for name in all_persons}

        for sch in schedules:
            date_str = sch["date"]
            day_set = sch["day"]
            night_set = sch["night"]

            for name in all_persons:
                in_day = name in day_set
                in_night = name in night_set
                attendance_records[name].append(in_day or in_night)

                if in_day and in_night:
                    stats[name]["consecutive"] += 1
                elif in_day and not in_night:
                    stats[name]["pure_day"] += 1
                elif not in_day and in_night:
                    stats[name]["pure_night"] += 1

        for name in all_persons:
            stats[name]["total_night"] = stats[name]["consecutive"] + stats[name]["pure_night"]

        for name in all_persons:
            rest = 0
            cnt = 0
            for present in attendance_records[name]:
                if not present:
                    cnt += 1
                else:
                    if cnt >= 2:
                        rest += (cnt - 1)
                    cnt = 0
            if cnt >= 2:
                rest += (cnt - 1)
            stats[name]["rest_days"] = rest

        result_data = []
        for name in all_persons:
            consecutive = stats[name]["consecutive"]
            pure_day = stats[name]["pure_day"]
            pure_night = stats[name]["pure_night"]

            # 计算累计在岗时间
            main_shift_min = 15 * 60 + 30
            day_shift_min = 8 * 60 + 30
            night_shift_min = 15 * 60 + 30
            total_minutes = consecutive * main_shift_min + pure_day * day_shift_min + pure_night * night_shift_min
            hours = total_minutes // 60
            minutes = total_minutes % 60
            total_time_str = f"{hours}:{minutes:02d}"

            # 计算在岗天数 (运管主班1.5，夜班1.5，白班1)
            work_days = consecutive * 1.5 + pure_day * 1.0 + pure_night * 1.5
            # 格式化天数（如果是整数去掉小数点，如果带.5保留）
            work_days_str = f"{work_days:.1f}".rstrip('0').rstrip('.') if work_days % 1 != 0 else str(int(work_days))

            result_data.append({
                "姓名": name,
                "运管主班": consecutive,
                "运控白班": pure_day,
                "运控夜班": pure_night,
                "补贴天数": stats[name]["total_night"],
                "休息天数": stats[name]["rest_days"],
                "在岗天数": work_days_str,  # 新增列
                "累计在岗时间": total_time_str
            })

        result_df = pd.DataFrame(result_data).sort_values(by="运管主班", ascending=False)

        st.subheader("📌 值班人员统计表")
        st.dataframe(result_df, use_container_width=True, height=400)

        csv = result_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 下载完整统计表 (CSV)", csv, "shift_statistics.csv", "text/csv", key="e_download")
        
        # ============ 计算逻辑备注 ============
        st.markdown("---")
        st.markdown("""
        ### 📝 计算逻辑说明
        
        **1. 在岗天数计算逻辑：**
        *   **运管主班（连班）**：按 **1.5 天** 计算。
        *   **运控夜班（纯夜班）**：按 **1.5 天** 计算。
        *   **运控白班（纯白班）**：按 **1.0 天** 计算。
        
        **2. 累计在岗时间计算逻辑：**
        *   **运管主班（连班）**：单次按 **15小时30分钟** 计算。
        *   **运控夜班（纯夜班）**：单次按 **15小时30分钟** 计算。
        *   **运控白班（纯白班）**：单次按 **8小时30分钟** 计算。
        
        **3. 休息天数计算逻辑：**
        *   仅当**连续休息天数 ≥ 2天**时，才计为有效休息天数。
        *   连续休息 N 天，计为 **(N - 1) 天**（例如：连续休息 3 天算 2 天休息，单休 1 天不计入休息天数）。
        """)














# ==============================
# 功能 G：航路处理工具（HTML/JS 沙箱版）
# ==============================
def run_feature_g():
    G_HTML = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>航路处理工具</title>
    <style>
        body { font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif; margin: 12px; color:#333; font-size:16px; }
        textarea {
            width: 100%; height: 300px; box-sizing: border-box;
            font-family: Consolas, "Courier New", monospace; font-size: 15px;
            padding: 10px; border: 1px solid #ccc; border-radius: 6px;
            line-height: 1.6;
        }
        button {
            padding: 10px 20px; font-size: 16px; border-radius: 6px;
            border: 1px solid #ddd; background: #fff; cursor: pointer; margin-right: 8px;
        }
        button.primary { background: #ff4b4b; color: #fff; border-color: #ff4b4b; font-weight: bold; }
        button.primary:hover { background: #e63939; }
        button:hover { background: #f5f5f5; }
        button.primary:hover { background: #e63939; }
        .toolbar { margin: 12px 0; }
        .result-box {
            background: #f5f5f5; border: 1px solid #e0e0e0; border-radius: 6px;
            padding: 14px; font-family: Consolas, "Courier New", monospace;
            font-size: 16px; line-height: 1.7; white-space: pre-wrap; word-break: break-all;
            max-height: 500px; overflow-y: auto;
        }
        .success { color: #2e7d32; background: #e8f5e9; padding: 8px; border-radius: 4px; margin: 8px 0; font-size: 15px; }
        .error { color: #d32f2f; background: #ffebee; padding: 8px; border-radius: 4px; margin: 8px 0; font-size: 15px; }
        .info { color: #1976d2; background: #e3f2fd; padding: 8px; border-radius: 4px; margin: 8px 0; font-size: 15px; }
        .label { color: #555; font-size: 15px; margin: 6px 0; }
    </style>
</head>
<body>
    <div class="label">📋 请输入待处理的航路文本</div>
    <textarea id="inputText" placeholder="粘贴民航航线数据，支持多行表格格式/纯中文描述格式..."></textarea>

    <div class="toolbar">
        <button class="primary" id="processBtn">⚙️ 处理</button>
        <button id="clearBtn">🗑️ 清空</button>
        <button id="copyBtn">📋 复制结果</button>
        <span id="copyStatus" style="margin-left:8px; color:#2e7d32; font-size:15px;"></span>
    </div>

    <div id="status"></div>

    <div id="resultSection" style="display:none;">
        <div class="label">📊 处理结果</div>
        <div class="result-box" id="resultBox"></div>
    </div>

    <script>
        const INPUT_CACHE_KEY = 'route_input_cache_v1';

        // ===== localStorage =====
        function saveInput(text) {
            try { localStorage.setItem(INPUT_CACHE_KEY, text); } catch(e) {}
        }
        function loadInput() {
            try { return localStorage.getItem(INPUT_CACHE_KEY) || ''; } catch(e) { return ''; }
        }
        function escapeHtml(s) {
            return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        // ===== 基础工具（对应 Python 同名函数） =====
        function parseCoord(coordStr) {
            const letter = coordStr[0];
            const numPart = coordStr.slice(1);
            if (letter === 'N') {
                let deg = parseInt(numPart.slice(0, 2));
                let minute = parseInt(numPart.slice(2, 4));
                let secPart = numPart.slice(4);
                let secInt;
                if (secPart.indexOf('.') !== -1) {
                    secInt = Math.round(parseFloat(secPart));
                } else {
                    secInt = parseInt(secPart);
                }
                if (secInt >= 60) {
                    secInt -= 60;
                    minute += 1;
                    if (minute >= 60) { minute -= 60; deg += 1; }
                }
                return String(deg).padStart(2, '0') + String(minute).padStart(2, '0') + String(secInt).padStart(2, '0');
            } else if (letter === 'E') {
                let deg = parseInt(numPart.slice(0, 3));
                let minute = parseInt(numPart.slice(3, 5));
                let secPart = numPart.slice(5);
                let secInt;
                if (secPart.indexOf('.') !== -1) {
                    secInt = Math.round(parseFloat(secPart));
                } else {
                    secInt = parseInt(secPart);
                }
                if (secInt >= 60) {
                    secInt -= 60;
                    minute += 1;
                    if (minute >= 60) { minute -= 60; deg += 1; }
                }
                return String(deg).padStart(3, '0') + String(minute).padStart(2, '0') + String(secInt).padStart(2, '0');
            }
            throw new Error('未知的坐标前缀: ' + letter);
        }

        function baseName(s) { return s.split('@')[0]; }

        function isOpenPoint(s) {
            const base = baseName(s);
            if (/^[A-Z]{2,5}$/.test(base)) return true;
            if (/^P[A-Z]+$/.test(base)) return true;
            return false;
        }

        function isPPoint(s) {
            const base = baseName(s);
            return /^P\d+$/.test(base);
        }

        function cleanRoute(r) {
            if (r.startsWith('#')) return r.slice(1);
            return r;
        }

        function isOpenRoute(rt) {
            return rt && (rt[0] !== 'H' && rt[0] !== 'J' && rt[0] !== 'V');
        }

        function isClosedRoute(rt) {
            return rt.startsWith('H') || rt.startsWith('J') || rt.startsWith('V');
        }

        // ===== 表格格式提取 =====
        function extractTable(text) {
            let tokens = text.trim().split(/\s+/);
            let startIdx = 0;
            for (let i = 0; i < tokens.length; i++) {
                if (/^\d+$/.test(tokens[i]) && parseInt(tokens[i]) >= 1 && parseInt(tokens[i]) <= 40) {
                    startIdx = i;
                    break;
                }
            }
            tokens = tokens.slice(startIdx);

            const lines = [];
            let i = 0;
            while (i < tokens.length) {
                if (/^\d+$/.test(tokens[i])) {
                    const line = [tokens[i]];
                    i++;
                    while (i < tokens.length && !/^\d+$/.test(tokens[i])) {
                        line.push(tokens[i]);
                        i++;
                    }
                    lines.push(line);
                } else {
                    i++;
                }
            }

            const points = [];
            const routes = [];
            for (const line of lines) {
                let latIdx = -1;
                for (let idx = 0; idx < line.length; idx++) {
                    const tok = line[idx];
                    if (tok.startsWith('N') && /^\d+(\.\d+)?$/.test(tok.slice(1))) {
                        latIdx = idx;
                        break;
                    }
                }
                if (latIdx === -1) continue;
                const lonIdx = latIdx + 1;
                if (lonIdx >= line.length || !line[lonIdx].startsWith('E')) continue;
                const latStr = line[latIdx];
                const lonStr = line[lonIdx];

                let route = null;
                if (lonIdx + 1 < line.length) {
                    const nextTok = line[lonIdx + 1];
                    if (/^[A-Z][A-Z0-9]*$/.test(nextTok) && !/^\d/.test(nextTok[0])) {
                        route = nextTok;
                    }
                }

                let pointName = null;
                for (let j = latIdx - 1; j > 0; j--) {
                    const tok = line[j];
                    if (isOpenPoint(tok) || isPPoint(tok)) {
                        pointName = tok;
                        break;
                    }
                }
                if (pointName === null) continue;

                let pointDisplay;
                if (isPPoint(pointName)) {
                    const latInt = parseCoord(latStr);
                    const lonInt = parseCoord(lonStr);
                    pointDisplay = pointName + '@' + latInt + 'N' + lonInt + 'E';
                } else {
                    pointDisplay = pointName;
                }

                points.push(pointDisplay);
                if (route !== null) routes.push(route);
            }

            const seq = [];
            for (let i = 0; i < points.length; i++) {
                seq.push(points[i]);
                if (i < routes.length) seq.push(routes[i]);
            }
            return seq;
        }

        // ===== 中文描述格式提取 =====
        function extractChinese(text) {
            // 去掉中文标点（保留英文括号和字母数字）
            text = text.replace(/[\u4e00-\u9fa5，、。；：""''（）【】]/g, ' ');
            const words = text.split(/\s+/).filter(w => w);
            const seq = [];
            for (const w of words) {
                if (w.indexOf('(') !== -1 && w.indexOf(')') !== -1) {
                    const m = w.match(/\(([A-Z]+)\)/);
                    if (m) {
                        const point = m[1];
                        const prefix = w.slice(0, w.indexOf('('));
                        const mRoute = prefix.match(/([A-Z]\d+)$/);
                        if (mRoute) seq.push(mRoute[1]);
                        seq.push(point);
                    }
                } else if (/^[A-Z]\d+[A-Z]{2,5}$/.test(w) || /^[A-Z]\d+P\d+$/.test(w)) {
                    const m = w.match(/^([A-Z]\d+)([A-Z]{2,5}|P\d+)$/);
                    if (m) {
                        seq.push(m[1]);
                        seq.push(m[2]);
                    }
                } else if (/^[A-Z]\d+$/.test(w)) {
                    seq.push(w);
                } else if (isOpenPoint(w) || isPPoint(w)) {
                    seq.push(w);
                }
            }
            return seq;
        }

        // ===== 步骤 =====
        function step1Extract(text) {
            if (/N\d{5,6}(\.\d+)?\s+E\d{6,7}(\.\d+)?/.test(text)) {
                return { seq: extractTable(text), fmt: 'table' };
            } else {
                return { seq: extractChinese(text), fmt: 'chinese' };
            }
        }

        function step2Reduce(seq) {
            let L = seq.slice();
            let changed = true;
            while (changed) {
                changed = false;
                const n = L.length;
                const candidates = [];
                for (let i = 0; i < n; i += 2) {
                    if (!isOpenPoint(L[i])) continue;
                    if (i + 1 >= n) continue;
                    const firstRoute = cleanRoute(L[i + 1]);
                    if (!isOpenRoute(firstRoute)) continue;
                    for (let j = i + 2; j < n; j += 2) {
                        let allSame = true;
                        for (let k = i + 1; k < j; k += 2) {
                            const rt = cleanRoute(L[k]);
                            if (rt !== firstRoute || !isOpenRoute(rt)) {
                                allSame = false;
                                break;
                            }
                        }
                        if (!allSame) break;
                        if (isOpenPoint(L[j])) {
                            const length = Math.floor((j - i) / 2);
                            if (length >= 2) {
                                candidates.push([i, j, length]);
                            }
                        }
                    }
                }
                if (candidates.length === 0) break;
                candidates.sort((a, b) => b[2] - a[2]);
                const [bestI, bestJ] = candidates[0];
                const newSegment = [L[bestI], L[bestI + 1], L[bestJ]];
                L = L.slice(0, bestI).concat(newSegment).concat(L.slice(bestJ + 1));
                changed = true;
            }
            return L;
        }

        function step3AddHash(seq) {
            const pts = seq.filter((_, i) => i % 2 === 0);
            const rts = seq.filter((_, i) => i % 2 === 1);
            const res = [pts[0]];
            for (let i = 0; i < rts.length; i++) {
                const rt = rts[i];
                const left = pts[i];
                const right = pts[i + 1];
                let needHash = false;
                if (isClosedRoute(rt)) needHash = true;
                else if (isPPoint(left) || isPPoint(right)) needHash = true;
                res.push(needHash ? '#' + rt : rt);
                res.push(right);
            }
            return res;
        }

        // ===== 主流程 =====
        function process() {
            const status = document.getElementById('status');
            const inputText = document.getElementById('inputText').value;
            if (!inputText.trim()) {
                status.innerHTML = '<div class="error">请输入待处理的航路文本</div>';
                document.getElementById('resultSection').style.display = 'none';
                return;
            }
            try {
                const { seq: seq1, fmt } = step1Extract(inputText);
                let seq = seq1;
                if (fmt === 'table') {
                    seq = step2Reduce(seq);
                    seq = step3AddHash(seq);
                }
                const result = seq.length > 0 ? seq.join(' ') : '⚠️ 未提取到有效航路数据';

                document.getElementById('resultBox').textContent = result;
                document.getElementById('resultSection').style.display = 'block';
                status.innerHTML = '<div class="success">✅ 处理完成</div>';
            } catch (e) {
                status.innerHTML = '<div class="error">❌ 处理失败：' + escapeHtml(e.message) + '</div>';
                console.error(e);
            }
        }

        // ===== 事件绑定 =====
        const inputEl = document.getElementById('inputText');
        let saveTimer = null;
        inputEl.addEventListener('input', () => {
            clearTimeout(saveTimer);
            saveTimer = setTimeout(() => saveInput(inputEl.value), 500);
        });

        document.getElementById('processBtn').addEventListener('click', process);

        document.getElementById('clearBtn').addEventListener('click', () => {
            inputEl.value = '';
            saveInput('');
            document.getElementById('resultSection').style.display = 'none';
            document.getElementById('status').innerHTML = '';
            document.getElementById('copyStatus').textContent = '';
        });

        document.getElementById('copyBtn').addEventListener('click', async () => {
            const text = document.getElementById('resultBox').textContent;
            if (!text) { return; }
            const statusEl = document.getElementById('copyStatus');
            try {
                await navigator.clipboard.writeText(text);
                statusEl.textContent = '✅ 已复制';
                setTimeout(() => { statusEl.textContent = ''; }, 1500);
            } catch (e) {
                const ta = document.createElement('textarea');
                ta.value = text;
                ta.style.position = 'fixed';
                ta.style.left = '-9999px';
                document.body.appendChild(ta);
                ta.select();
                try {
                    document.execCommand('copy');
                    statusEl.textContent = '✅ 已复制';
                    setTimeout(() => { statusEl.textContent = ''; }, 1500);
                } catch (e2) {
                    statusEl.textContent = '❌ 复制失败';
                    statusEl.style.color = '#d32f2f';
                }
                document.body.removeChild(ta);
            }
        });

        // ===== 页面加载：恢复上次输入 =====
        window.addEventListener('DOMContentLoaded', () => {
            const saved = loadInput();
            if (saved) inputEl.value = saved;
        });
    </script>
</body>
</html>
"""
    components.html(G_HTML, height=900, scrolling=True)













# ==============================
# 功能 H：检查单/脚本生成器 + 世界时行程（合并版）
# ==============================
def run_feature_chk():
    CHK_HTML = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>检查单 & 世界时行程</title>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <style>
        body { font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif; margin: 12px; color:#333; font-size:16px; }
        h3 { margin: 18px 0 8px 0; font-size: 19px; }
        input[type=file] { padding: 6px; font-size: 15px; }
        button { padding: 10px 18px; font-size: 16px; border-radius: 6px; border: 1px solid #ddd; background:#fff; cursor: pointer; margin-right: 6px; transition: all 0.2s; }
        button:hover { background:#f5f5f5; }
        button.primary { background:#ff4b4b; color:#fff; border-color:#ff4b4b; font-weight: bold; }
        button.primary:hover { background:#e63939; }
        button.active { background:#ff9800; color:#fff; border-color:#ff9800; font-weight: bold; }
        button.active:hover { background:#f57c00; }
        
        .code-block { background:#f5f5f5; padding:12px; border-radius:4px; font-family: Consolas, "Courier New", monospace; font-size:16px; white-space: pre-wrap; word-break: break-all; max-height: 400px; overflow-y:auto; border:1px solid #e0e0e0; }
        .status { color:#555; font-size:15px; margin-left:8px; }
        .error { color:#d32f2f; background:#ffebee; padding:8px; border-radius:4px; margin:6px 0; font-size:15px; }
        .success { color:#2e7d32; background:#e8f5e9; padding:8px; border-radius:4px; margin:6px 0; font-size:15px; }
        .info { color:#1976d2; background:#e3f2fd; padding:8px; border-radius:4px; margin:6px 0; font-size:15px; }
        details { margin: 10px 0; padding: 10px; border: 1px solid #eee; border-radius: 4px; background:#fafafa; }
        summary { cursor: pointer; font-weight: bold; padding: 6px 0; font-size: 18px; }
        .pp-aircraft { font-weight: bold; font-size: 17px; margin: 10px 0 6px 0; }
        table { border-collapse: collapse; }
        .toolbar { margin: 10px 0 14px 0; display: flex; align-items: center; flex-wrap: wrap; gap: 8px; }
        
        /* 世界时样式 */
        .wt-block { margin-bottom: 20px; }
        .wt-title { font-weight: bold; font-size: 19px; margin-bottom: 8px; }
        .wt-list { background:#f7f7f7; border: 1px solid #ddd; border-radius: 6px; padding: 10px 14px; font-family: Consolas, "Courier New", monospace; font-size: 16px; line-height: 1.9; color:#222; }
        .wt-line { padding: 3px 0; }
    </style>
</head>
<body>
    <input type="file" id="fileInput" accept=".xlsx,.xls">
    <div id="status"></div>

    <div id="result" style="display:none;">
        <div class="toolbar">
            <button class="primary" id="copyBtn">📋 一键复制（今明后）检查单</button>
            <button id="toggleAllBtn">👁️ 显示所有计划</button>
            <button id="switchViewBtn" style="margin-left: 10px;">🌐 世界时行程</button>
            <span class="status" id="copyStatus"></span>
        </div>

        <!-- 检查单视图 -->
        <div id="checklistView">
            <h3>📦 PRELIM / PACKAGE</h3>
            <div id="prelimPackage"></div>

            <h3>📜 JavaScript 填充脚本</h3>
            <div style="margin-bottom: 8px;">
                <button id="copyJsBtn">📋 复制脚本</button>
                <span class="status" id="copyJsStatus"></span>
            </div>
            <div class="code-block" id="jsScript"></div>
        </div>

        <!-- 世界时视图 -->
        <div id="worldtimeView" style="display:none;">
            <div style="margin-bottom: 12px;">
                <button id="toggleDomesticBtn" class="primary">👁️ 显示全部计划（含纯内地）</button>
            </div>
            <div id="worldtimeList"></div>
        </div>
    </div>

    <script>
        const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
        const CHK_PREFERRED = ["B3926", "B652R", "B8105", "B8160", "B8262", "B8292", "B8309",
            "N2QE", "N328LM", "N550DR", "N577QT", "N7777U", "N777ZH",
            "N88AY", "T7178HT", "T7CJK", "VPCSZ", "VPCVA",
            "B652Q", "B652S", "B65AP", "MLLIN"];
        const chkPriority = {};
        CHK_PREFERRED.forEach((ac, i) => chkPriority[ac] = i);
        const chkDefaultPri = CHK_PREFERRED.length;

        // 原版世界时专属排序优先级
        const WT_PRIORITY = ['B652Q', 'B65AP', 'B652S', 'MLLIN', 'N88AY', 'B652R'];

        // 状态变量
        let currentChkRows = [];
        let currentChkFlights = [];
        let currentChkRawItems = [];
        let currentWorldTimeFlights = [];
        let currentJsScript = '';
        
        let showAllPlans = false; // 检查单日期展开状态
        let showDomestic = false; // 世界时纯内地展示状态（默认 false = 隐藏纯内地）
        let currentView = 'checklist'; // 'checklist' 或 'worldtime'

        // ============ 工具函数 ============
        function pad2(n) { return String(n).padStart(2, '0'); }

        function parseDate(v) {
            if (v === null || v === undefined || v === '') return null;
            if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
            if (typeof v === 'number') {
                const ms = Math.round((v - 25569) * 86400 * 1000);
                const d = new Date(ms);
                return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
            }
            const s = String(v).trim();
            let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
            if (m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
            const d = new Date(s);
            if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return null;
        }

        function parseTime(v) {
            if (v === null || v === undefined || v === '') return null;
            if (v instanceof Date) return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
            if (typeof v === 'number') {
                const totalMin = Math.round(v * 24 * 60);
                return pad2(Math.floor(totalMin / 60) % 24) + ':' + pad2(totalMin % 60);
            }
            const s = String(v).trim();
            const m = s.match(/(\d{1,2}):(\d{2})/);
            if (m) return pad2(parseInt(m[1])) + ':' + m[2];
            return null;
        }

        function getUTCdate(dateVal, timeVal) {
            const d = parseDate(dateVal);
            if (!d) return null;
            const t = parseTime(timeVal);
            let bj;
            if (t) {
                const parts = t.split(':');
                bj = new Date(d.getFullYear(), d.getMonth(), d.getDate(), parseInt(parts[0]), parseInt(parts[1]));
            } else {
                bj = d;
            }
            const utc = new Date(bj.getTime() - 8 * 3600 * 1000);
            return pad2(utc.getDate()) + MONTHS[utc.getMonth()];
        }

        function getBJdate(dateVal) {
            const d = parseDate(dateVal);
            if (!d) return null;
            return pad2(d.getDate()) + MONTHS[d.getMonth()];
        }

        function toUTCLabel(dateVal, timeStr) {
            const d = parseDate(dateVal);
            if (!d || !timeStr) return null;
            const parts = timeStr.split(':');
            const h = parseInt(parts[0]);
            const m = parseInt(parts[1]);

            let total = h * 60 + m - 8 * 60;
            let dayOff = 0;
            while (total < 0) { total += 24 * 60; dayOff--; }
            while (total >= 24 * 60) { total -= 24 * 60; dayOff++; }
            const uh = Math.floor(total / 60);
            const um = total % 60;

            const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            dd.setDate(dd.getDate() + dayOff);

            return {
                day: dd.getDate(),
                month: dd.getMonth() + 1,
                hours: uh,
                minutes: um,
                sortKey: dd.getFullYear() * 100000000 + (dd.getMonth() + 1) * 1000000 + dd.getDate() * 10000 + uh * 100 + um
            };
        }

        function formatUTCLabel(u) {
            if (!u) return '';
            return pad2(u.day) + MONTHS[u.month - 1] + ' ' + pad2(u.hours) + pad2(u.minutes) + 'Z';
        }

        // 纯内地机场判断
        function isMainlandAirport(icao) {
            if (!icao) return false;
            const s = String(icao).trim().toUpperCase();
            if (s.length !== 4) return false;
            if (s.charAt(0) !== 'Z') return false;
            if (s.indexOf('ZM') === 0) return false;
            return true;
        }

        function isDomesticLine(line) {
            const m = line.match(/ETD\s+(\S+)\s+.*?ETA\s+(\S+)/);
            if (!m) return false;
            return isMainlandAirport(m[1]) && isMainlandAirport(m[2]);
        }

        function escapeHtml(s) {
            return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        // ============ 主流程 ============
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    processWorkbook(ev.target.result, file.name);
                } catch (err) {
                    showError('处理失败：' + err.message);
                    console.error(err);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function processWorkbook(data, fileName) {
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

            let headerIdx = -1;
            for (let i = 0; i < Math.min(rows.length, 10); i++) {
                const vals = rows[i].map(v => String(v).trim());
                if (vals.includes('飞机注册号') && vals.includes('出发地') && vals.includes('到达地')) {
                    headerIdx = i;
                    break;
                }
            }
            if (headerIdx === -1) {
                showError('未在文件中找到标题行（需包含：飞机注册号、出发地、到达地）');
                return;
            }

            const headers = rows[headerIdx].map(h => String(h).trim());
            function findCol(cands) {
                for (const c of cands) {
                    const idx = headers.indexOf(c);
                    if (idx !== -1) return idx;
                }
                for (const c of cands) {
                    for (let i = 0; i < headers.length; i++) {
                        if (headers[i].includes(c)) return i;
                    }
                }
                return -1;
            }

            const colAircraft = findCol(['飞机注册号', '注册号', '机号']);
            const colOrigin = findCol(['出发地', '起飞机场']);
            const colDest = findCol(['到达地', '目的地机场']);
            const colDate = findCol(['出发日期', '起飞日期']);
            const colTime = findCol(['计划出发', '起飞时间']);
            const colPurpose = findCol(['用途']);
            const colDepCity = findCol(['出发城市']);
            const colArrCity = findCol(['到达城市']);
            const colArrDate = findCol(['到达日期']);
            const colArrTime = findCol(['预计到达', '到达时间']);

            if (colAircraft === -1 || colOrigin === -1 || colDest === -1) {
                showError('Excel 缺少必要列。当前列名：' + headers.join(', '));
                return;
            }

            const chkFlights = [];
            const chkRawItems = [];
            const worldTimePlans = {}; 

            for (let i = headerIdx + 1; i < rows.length; i++) {
                const r = rows[i];
                if (!r || r.length === 0) continue;
                const get = (idx) => (idx >= 0 && idx < r.length) ? r[idx] : '';

                const aircraft = get(colAircraft);
                if (aircraft === '' || aircraft === null || aircraft === undefined) continue;
                const acStr = String(aircraft).trim();
                if (/[\u4e00-\u9fa5]/.test(acStr)) continue;

                const origin = get(colOrigin);
                const dest = get(colDest);
                if (origin === '' || dest === '') continue;
                const orgStr = String(origin).trim().toUpperCase();
                const dstStr = String(dest).trim().toUpperCase();

                const dtObj = parseDate(get(colDate));
                const timestamp = dtObj ? dtObj.getTime() : 0;

                const depDate = get(colDate);
                const depTime = get(colTime);
                const arrDate = get(colArrDate);
                const arrTime = get(colArrTime);
                const purpose = get(colPurpose);

                const depTimeStr = parseTime(depTime);
                const arrTimeStr = parseTime(arrTime);

                // 1. 收集检查单用数据
                if (orgStr.length === 4 && /^[A-Z]+$/.test(orgStr) &&
                    dstStr.length === 4 && /^[A-Z]+$/.test(dstStr)) {
                    chkFlights.push({
                        aircraft: acStr, origin: orgStr, dest: dstStr,
                        date: getUTCdate(depDate, depTime), date_bj: getBJdate(depDate),
                        timestamp: timestamp
                    });
                }

                // 2. 收集提取文本用数据
                if (depTimeStr && arrTimeStr) {
                    const depCityStr = get(colDepCity) ? String(get(colDepCity)).trim() : '';
                    const arrCityStr = get(colArrCity) ? String(get(colArrCity)).trim() : '';
                    const purposeStr = purpose ? String(purpose).trim() : '';

                    const depDtObj = parseDate(depDate);
                    const arrDtObj = parseDate(arrDate);
                    let dayDiff = 0;
                    if (depDtObj && arrDtObj) {
                        dayDiff = Math.round((arrDtObj.getTime() - depDtObj.getTime()) / 86400000);
                    }
                    const plus = dayDiff > 0 ? ' +' + dayDiff : '';

                    const line1 = acStr + ' ' + depTimeStr + ' - ' + arrTimeStr + plus;
                    const line2 = depCityStr + ' - ' + arrCityStr;
                    const content = purposeStr.indexOf('调机') !== -1 ? 'F\n' + line1 + '\n' + line2 : line1 + '\n' + line2;

                    chkRawItems.push({
                        aircraft: acStr, depDt: depDtObj, depTimeStr: depTimeStr, content: content,
                    });
                }

                // 3. 收集世界时数据
                const depUtc = toUTCLabel(depDate, depTimeStr);
                const arrUtc = toUTCLabel(arrDate, arrTimeStr);
                if (depUtc && arrUtc) {
                    const use = String(purpose || '');
                    const flightType = use.indexOf('调机') !== -1 ? 'FERRY' : 'PAX';
                    
                    const wtLine = 'ETD ' + orgStr + ' ' + formatUTCLabel(depUtc) +
                                 ' // ETA ' + dstStr + ' ' + formatUTCLabel(arrUtc) +
                                 '  ' + flightType;
                    
                    if (!worldTimePlans[acStr]) worldTimePlans[acStr] = [];
                    worldTimePlans[acStr].push({ sortKey: depUtc.sortKey, line: wtLine });
                }
            }

            if (chkFlights.length === 0 && chkRawItems.length === 0 && Object.keys(worldTimePlans).length === 0) {
                showError('❌ 未读取到有效航段数据');
                return;
            }

            // ---- 排序检查单数据 ----
            chkRawItems.sort((a, b) => {
                const ta = a.depDt ? a.depDt.getTime() : Number.MAX_SAFE_INTEGER;
                const tb = b.depDt ? b.depDt.getTime() : Number.MAX_SAFE_INTEGER;
                if (ta !== tb) return ta - tb;
                const pa = chkPriority[a.aircraft] !== undefined ? chkPriority[a.aircraft] : chkDefaultPri;
                const pb = chkPriority[b.aircraft] !== undefined ? chkPriority[b.aircraft] : chkDefaultPri;
                if (pa !== pb) return pa - pb;
                return (a.depTimeStr || '99:99').localeCompare(b.depTimeStr || '99:99');
            });

            const chkRows = [];
            let prevDate = null, prevAc = null, first = true;
            for (const it of chkRawItems) {
                const curDate = it.depDt ? it.depDt.getTime() : null;
                const curAc = it.aircraft;
                if (!first && (curDate !== prevDate || curAc !== prevAc)) {
                    chkRows.push({ type: 'blank' });
                }
                // 注意这里：加入了 timestamp，让复制时可以判断日期
                chkRows.push({ type: 'data', content: it.content, timestamp: curDate || 0 });
                prevDate = curDate;
                prevAc = curAc;
                first = false;
            }

            // ---- 排序世界时数据 ----
            const worldTimeFlights = [];
            for (const ac in worldTimePlans) {
                worldTimePlans[ac].sort((a, b) => a.sortKey - b.sortKey);
                const lines = [ac];
                worldTimePlans[ac].forEach(it => lines.push(it.line));
                worldTimeFlights.push({ reg: ac, text: lines.join('\n') });
            }
            
            worldTimeFlights.sort((a, b) => {
                const ia = WT_PRIORITY.indexOf(a.reg);
                const ib = WT_PRIORITY.indexOf(b.reg);
                if (ia !== -1 && ib !== -1) return ia - ib;
                if (ia !== -1) return -1;
                if (ib !== -1) return 1;
                if (a.reg === 'N/A' && b.reg !== 'N/A') return 1;
                if (b.reg === 'N/A' && a.reg !== 'N/A') return -1;
                return a.reg.localeCompare(b.reg);
            });

            currentChkRows = chkRows;
            currentChkFlights = chkFlights;
            currentChkRawItems = chkRawItems;
            currentWorldTimeFlights = worldTimeFlights;

            renderAll();

            const now = new Date();
            const timestamp = now.getFullYear() + '-' + pad2(now.getMonth()+1) + '-' + pad2(now.getDate()) +
                              ' ' + pad2(now.getHours()) + ':' + pad2(now.getMinutes()) + ':' + pad2(now.getSeconds());
            
            const status = document.getElementById('status');
            status.innerHTML = '<div class="success">✅ 文件读取成功：' + escapeHtml(fileName) +
                '（' + timestamp + '，已自动缓存）</div>';
        }

        // ============ 渲染控制 ============
        function renderAll() {
            document.getElementById('result').style.display = 'block';
            renderPrelimPackage(currentChkFlights, showAllPlans);
            generateJsScript(currentChkFlights);
            renderWorldTime(currentWorldTimeFlights, showDomestic);
            updateUIState();
        }

        function updateUIState() {
            const btnSwitch = document.getElementById('switchViewBtn');
            const viewChk = document.getElementById('checklistView');
            const viewWt = document.getElementById('worldtimeView');
            
            if (currentView === 'checklist') {
                btnSwitch.textContent = '🌐 世界时行程';
                btnSwitch.className = '';
                viewChk.style.display = 'block';
                viewWt.style.display = 'none';
            } else {
                btnSwitch.textContent = '🌐 当前：世界时行程';
                btnSwitch.className = 'active';
                viewChk.style.display = 'none';
                viewWt.style.display = 'block';
            }
            
            const btnDomestic = document.getElementById('toggleDomesticBtn');
            if (btnDomestic) {
                btnDomestic.textContent = showDomestic ? '🔒 隐藏纯内地计划' : '👁️ 显示全部计划（含纯内地）';
            }
        }

        // ============ 检查单渲染 ============
        function renderPrelimPackage(flights, showAll) {
            const container = document.getElementById('prelimPackage');
            if (!flights || flights.length === 0) {
                container.innerHTML = '<div class="info">无 PRELIM/PACKAGE 数据</div>';
                return;
            }
            
            const dateGroups = {};
            for (const f of flights) {
                if (!f.date_bj) continue;
                const ts = f.timestamp || 0;
                if (!dateGroups[ts]) {
                    dateGroups[ts] = { displayDate: f.date_bj, aircrafts: {} };
                }
                const ac = f.aircraft;
                if (!dateGroups[ts].aircrafts[ac]) dateGroups[ts].aircrafts[ac] = [];
                dateGroups[ts].aircrafts[ac].push(f);
            }

            const today = new Date(); today.setHours(0, 0, 0, 0);
            const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
            const dayAfter = new Date(today); dayAfter.setDate(today.getDate() + 2);

            const todayTime = today.getTime();
            const tomorrowTime = tomorrow.getTime();
            const dayAfterTime = dayAfter.getTime();

            const sortedTimestamps = Object.keys(dateGroups).map(Number).sort((a, b) => a - b);
            let html = '';

            for (const ts of sortedTimestamps) {
                if (!showAll && ts > dayAfterTime) continue;

                let showPrelim = true;
                let showPackage = true;

                if (ts === todayTime) {
                    showPrelim = false;
                } else if (ts === tomorrowTime) {
                    showPrelim = true; showPackage = true;
                } else if (ts >= dayAfterTime) {
                    showPrelim = true; showPackage = false;
                } 

                const group = dateGroups[ts];
                let groupHtml = '';
                let hasContent = false;
                
                const acGroups = group.aircrafts;
                const sortedAcs = Object.keys(acGroups).sort((a, b) => {
                    const pa = chkPriority[a] !== undefined ? chkPriority[a] : chkDefaultPri;
                    const pb = chkPriority[b] !== undefined ? chkPriority[b] : chkDefaultPri;
                    return pa - pb;
                });

                for (const ac of sortedAcs) {
                    const items = acGroups[ac];
                    const prelims = showPrelim ? items.map(f => 'PRELIM ' + f.aircraft + ' ' + f.origin + '-' + f.dest + ' ' + f.date) : [];
                    const packages = showPackage ? items.map(f => 'PACKAGE ' + f.aircraft + ' ' + f.origin + '-' + f.dest + ' ' + f.date) : [];
                    
                    let block = '';
                    if (prelims.length > 0) block += prelims.join('\n');
                    if (packages.length > 0) {
                        if (block) block += '\n\n';
                        block += packages.join('\n');
                    }
                    if (!block) continue;

                    hasContent = true;
                    groupHtml += '<div class="pp-aircraft">' + ac + '</div><div class="code-block">' + escapeHtml(block) + '</div>';
                }
                
                if (hasContent) {
                    html += '<details open><summary>📅 ' + group.displayDate + '</summary>';
                    html += groupHtml;
                    html += '</details>';
                }
            }
            container.innerHTML = html;
        }

        // ============ 世界时渲染 ============
        function renderWorldTime(flights, showAllDomestic) {
            const container = document.getElementById('worldtimeList');
            if (!flights || flights.length === 0) {
                container.innerHTML = '<div class="info">无世界时行程数据</div>';
                return;
            }

            let html = '';
            let hasValidData = false;

            for (const item of flights) {
                const lines = item.text.split('\n');
                const reg = lines[0];
                let routes = lines.slice(1);

                if (!showAllDomestic) {
                    routes = routes.filter(line => !isDomesticLine(line));
                }

                if (routes.length === 0) continue;
                hasValidData = true;

                html += '<div class="wt-block">';
                html += '<div class="wt-title">✈️ ' + escapeHtml(reg) + '</div>';
                html += '<div class="wt-list">';
                routes.forEach(line => {
                    html += '<div class="wt-line">' + escapeHtml(line) + '</div>';
                });
                html += '</div>';
                
                const copyText = reg + '\n' + routes.join('\n');
                const b64Text = btoa(unescape(encodeURIComponent(copyText)));
                html += '<button style="margin-top:6px;" onclick="copyTextFromB64(\''+b64Text+'\', this)">📋 复制该飞机</button>';
                
                html += '</div>';
            }

            if (!hasValidData) {
                container.innerHTML = '<div class="info">📭 所有计划均为纯内地航段，已全部隐藏。点击上方按钮可查看全部。</div>';
            } else {
                container.innerHTML = html;
            }
        }

        function copyTextFromB64(b64, btn) {
            const text = decodeURIComponent(escape(atob(b64)));
            navigator.clipboard.writeText(text).then(() => {
                const original = btn.textContent;
                btn.textContent = '✅ 已复制';
                setTimeout(() => btn.textContent = original, 1500);
            }).catch(() => {
                const ta = document.createElement('textarea');
                ta.value = text;
                document.body.appendChild(ta);
                ta.select();
                try { document.execCommand('copy'); } catch (e) {}
                document.body.removeChild(ta);
                const original = btn.textContent;
                btn.textContent = '✅ 已复制';
                setTimeout(() => btn.textContent = original, 1500);
            });
        }

        // ============ 复制检查单逻辑 ============
        async function copyChecklist() {
            const status = document.getElementById('copyStatus');
            const TD_STYLE = "text-align:center; vertical-align:middle; font-family:'Times New Roman', Times, serif; font-size:16pt; border:1px solid #000000; padding:2px 6px;";
            
            // 计算今天、明天、后天的边界时间戳
            const today = new Date(); today.setHours(0, 0, 0, 0);
            const dayAfter = new Date(today); dayAfter.setDate(today.getDate() + 2);
            const todayTime = today.getTime();
            const dayAfterEnd = dayAfter.getTime() + 86399999; // 后天 23:59:59.999

            const copyRows = [];
            let pendingBlank = null;

            for (const row of currentChkRows) {
                if (row.type === 'blank') {
                    pendingBlank = row;
                } else {
                    const ts = row.timestamp || 0;
                    // 只取今天、明天、后天的数据
                    if (ts >= todayTime && ts <= dayAfterEnd) {
                        if (pendingBlank) {
                            copyRows.push(pendingBlank);
                            pendingBlank = null;
                        }
                        copyRows.push(row);
                    } else {
                        pendingBlank = null; // 超出范围，丢弃暂存的空行
                    }
                }
            }

            if (copyRows.length === 0) {
                status.textContent = '⚠️ 今天、明天、后天没有检查单可复制';
                status.style.color = '#d32f2f';
                return;
            }

            let htmlTable = '<table style="border-collapse: collapse;">';
            for (const row of copyRows) {
                if (row.type === 'blank') {
                    htmlTable += '<tr><td style="' + TD_STYLE + '">&nbsp;</td></tr>';
                } else {
                    const content = row.content.split('\n').map(escapeHtml).join('<br>');
                    htmlTable += '<tr><td style="' + TD_STYLE + '">' + content + '</td></tr>';
                }
            }
            htmlTable += '</table>';

            const plain = copyRows.map(r => r.type === 'blank' ? '' : r.content).join('\n\n');

            try {
                await navigator.clipboard.write([
                    new ClipboardItem({
                        'text/html': new Blob([htmlTable], { type: 'text/html' }),
                        'text/plain': new Blob([plain], { type: 'text/plain' })
                    })
                ]);
                status.textContent = '✅ 已复制（今明后，保留格式）';
                status.style.color = '#2e7d32';
            } catch (e) {
                const div = document.createElement('div');
                div.innerHTML = htmlTable;
                div.style.position = 'fixed';
                div.style.left = '-9999px';
                document.body.appendChild(div);
                const range = document.createRange();
                range.selectNodeContents(div);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
                document.execCommand('copy');
                sel.removeAllRanges();
                document.body.removeChild(div);
                status.textContent = '✅ 已复制（降级模式）';
                status.style.color = '#2e7d32';
            }
        }

        function copyJsScript() {
            const status = document.getElementById('copyJsStatus');
            navigator.clipboard.writeText(currentJsScript).then(() => {
                status.textContent = '✅ 已复制';
                status.style.color = '#2e7d32';
            }).catch(e => {
                const ta = document.createElement('textarea');
                ta.value = currentJsScript;
                document.body.appendChild(ta);
                ta.select();
                try { document.execCommand('copy'); status.textContent = '✅ 已复制'; status.style.color = '#2e7d32'; }
                catch (e2) { status.textContent = '❌ 复制失败'; status.style.color = '#d32f2f'; }
                document.body.removeChild(ta);
            });
        }

        function generateJsScript(flights) {
            if (!flights || flights.length === 0) {
                currentJsScript = '// 无数据，未生成填充脚本';
                document.getElementById('jsScript').textContent = currentJsScript;
                return;
            }
            const template = [
'var flightData = __FLIGHT_DATA__;',
'var currentIndex = 0;',
'',
'function getElementByXpath(path) {',
'    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;',
'}',
'',
'function fillFlight(index) {',
'    if (index < 0 || index >= flightData.length) {',
'        console.error("索引超出范围 (0 ~ " + (flightData.length - 1) + ")");',
'        return false;',
'    }',
'    var data = flightData[index];',
'    console.log("🛫 填充第 " + (index+1) + "/" + flightData.length + " 条: " + data.aircraft + "  " + data.origin + " -> " + data.dest);',
'',
'    var sel = document.getElementById("Aircraft");',
'    if (!sel) { console.error("未找到 id=\'Aircraft\'"); return false; }',
'    var found = false;',
'    for (var opt of sel.options) {',
'        if (opt.value === data.aircraft) { opt.selected = true; found = true; break; }',
'    }',
'    if (!found) { sel.value = data.aircraft; }',
'    sel.dispatchEvent(new Event("change", { bubbles: true }));',
'',
'    var originInput = getElementByXpath("/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/input[1]");',
'    if (!originInput) { console.error("未找到起飞机场输入框"); return false; }',
'    originInput.focus(); originInput.value = data.origin;',
'    originInput.dispatchEvent(new Event("input", { bubbles: true }));',
'    originInput.dispatchEvent(new Event("change", { bubbles: true }));',
'    originInput.dispatchEvent(new Event("blur", { bubbles: true }));',
'    originInput.blur();',
'',
'    var destInput = getElementByXpath("/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[1]/input[1]");',
'    if (!destInput) { console.error("未找到目的地机场输入框"); return false; }',
'    destInput.focus(); destInput.value = data.dest;',
'    destInput.dispatchEvent(new Event("input", { bubbles: true }));',
'    destInput.dispatchEvent(new Event("change", { bubbles: true }));',
'    destInput.dispatchEvent(new Event("blur", { bubbles: true }));',
'    destInput.blur();',
'',
'    console.log("✅ 填充完成，请手动点击提交按钮！");',
'    return true;',
'}',
'',
'function fillNext() {',
'    if (currentIndex >= flightData.length) { console.log("🎉 全部填充完毕！"); return; }',
'    if (fillFlight(currentIndex)) { currentIndex++; }',
'}',
'',
'function resetIndex() { currentIndex = 0; console.log("🔄 索引已重置为 0"); }',
'',
'console.log("✅ 脚本加载成功，共 " + flightData.length + " 条航段");',
'console.log("📌 输入 fillNext() 填充下一条");',
''].join('\n');
            const dataJson = JSON.stringify(flights, null, 2);
            currentJsScript = template.replace('__FLIGHT_DATA__', dataJson);
            document.getElementById('jsScript').textContent = currentJsScript;
        }

        function showError(msg) {
            const status = document.getElementById('status');
            status.innerHTML = '<div class="error">' + escapeHtml(msg) + '</div>';
        }

        // ============ 事件绑定 ============
        window.addEventListener('DOMContentLoaded', () => {
            document.getElementById('copyBtn').addEventListener('click', copyChecklist);
            document.getElementById('copyJsBtn').addEventListener('click', copyJsScript);
            
            document.getElementById('toggleAllBtn').addEventListener('click', function() {
                showAllPlans = !showAllPlans;
                this.textContent = showAllPlans ? '🔒 隐藏未来计划' : '👁️ 显示所有计划';
                renderPrelimPackage(currentChkFlights, showAllPlans);
            });

            document.getElementById('switchViewBtn').addEventListener('click', function() {
                currentView = currentView === 'checklist' ? 'worldtime' : 'checklist';
                updateUIState();
            });

            document.getElementById('toggleDomesticBtn').addEventListener('click', function() {
                showDomestic = !showDomestic;
                renderWorldTime(currentWorldTimeFlights, showDomestic);
                updateUIState();
            });
        });
    </script>
</body>
</html>
"""
    components.html(CHK_HTML, height=1600, scrolling=True)
















# ==============================
# 功能 I：WX AND NOTAM 邮件生成（HTML/JS 沙箱版 + 完整持久化）
# ==============================
def run_feature_wx_mail():
    WX_HTML = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>飞行任务邮件生成器</title>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <style>
        body { font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif; margin: 12px; color:#333; font-size:16px; }
        h2 { margin: 6px 0 12px 0; font-size:22px; }
        h3 { margin: 16px 0 10px 0; font-size: 19px; }
        textarea { width: 100%; height: 170px; box-sizing: border-box; font-family: Consolas, monospace; font-size:15px; padding:8px; }
        #mailList { margin-top: 12px; }
        .mail-row {
            display: flex; align-items: stretch; margin: 8px 0; gap: 8px;
        }
        .mail-link {
            flex: 1; padding: 12px 14px; box-sizing: border-box;
            border-radius: 4px; text-decoration: none; color: #0066cc;
            border: 1px solid #ccc; background: #f0f0f0;
            transition: background 0.3s; font-size:16px;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }
        .mail-link:hover { filter: brightness(0.95); }
        .del-btn {
            padding: 0 18px; font-size: 17px; border-radius: 4px;
            border: 1px solid #ddd; background: #fff; color: #888;
            cursor: pointer; transition: all 0.2s; flex-shrink: 0;
        }
        .del-btn:hover { background: #ffebee; color: #d32f2f; border-color: #ef9a9a; }
        .error { color: red; padding:8px; background:#ffebee; border-radius:4px; margin:6px 0; font-size:15px; }
        .success { color:#2e7d32; padding:8px; background:#e8f5e9; border-radius:4px; margin:6px 0; font-size:15px; }
        .info { color:#1976d2; padding:8px; background:#e3f2fd; border-radius:4px; margin:6px 0; font-size:15px; }
        button.primary { padding: 10px 20px; font-size:16px; border-radius:6px; border:1px solid #ff4b4b; background:#ff4b4b; color:#fff; cursor:pointer; margin-right:8px; font-weight:bold; }
        button.primary:hover { background:#e63939; }
        button { padding: 10px 20px; font-size:16px; border-radius:6px; border:1px solid #ddd; background:#fff; cursor:pointer; margin-right:8px; }
        button:hover { background:#f5f5f5; }
        input[type=file] { padding:6px; font-size:15px; }
        .toolbar { margin: 12px 0; }
        #cacheStatus:empty { display: none; }
    </style>
</head>
<body>
    <h2>✉️ 飞行任务邮件生成器</h2>
    <div id="cacheStatus"></div>

    <p>上传航段表（Excel 或 CSV）：<span id="fileStatus" style="color:#2e7d32; font-size:15px; margin-left:8px;"></span></p>
    <input type="file" id="flightFile" accept=".csv,.xlsx,.xls">

    <p>粘贴文本飞行计划：</p>
    <textarea id="planText" placeholder="B652Q 06:00 - 07:55
北京大兴 - 上海虹桥
P035,P039,C051

B65AP 07:30 - 09:00
日本东京 羽田 - 日本福冈
P032,P036,C050,M021
..."></textarea>

    <div class="toolbar">
        <button class="primary" onclick="generate()">生成邮件链接</button>
        <button onclick="renderMails()">🔄 刷新列表</button>
        <button onclick="clearAllMails()">🗑️ 清空全部</button>
    </div>
    <div id="output"></div>

    <h3>📬 待处理邮件列表</h3>
    <div id="mailList"></div>

    <script>
        const PILOT_CSV = `工号,姓名,邮箱
P001,庚凡,gengfan@amber-aviation.com
P002,张永一,zhangyongyi@amber-aviation.com
P003,梅峰,fmei@amber-aviation.com
P004,王斌,wangbin@amber-aviation.com
P019,"HEALY, Darran William",darranhealy@amber-aviation.com
P020,"BEEBE, Thaddeus John",thaddeusbeebe@amber-aviation.com
P032,林毅,ericlin@amber-aviation.com
P035,"Peter Robert, JACKSON",prjackson@amber-aviation.com
P036,王少雄,warrenwang@amber-aviation.com
P038,苗旺旺,johnmiao@amber-aviation.com
P039,"Yiftah, RAUCH",yiftahrauch@amber-aviation.com
P044,李辛欣,rockli@amber-aviation.com
P046,赵岩松,zhyszhao@amber-aviation.com
P051,彭罡,eugene.peng@humbleholding.com
P052,胡君量,brian.wu@humbleholding.com
P053,"Bruce Roderick, WAINES",brwaines@amber-aviation.com
P054,"Rodolfo, BONETTI",rbonetti@amber-aviation.com
P056,"Keith Robert, SHERREN",krsherren@amber-aviation.com
P057,"Oliver Viktor, RACZ",ovracz@amber-aviation.com
P059,蔡国俊,kctsai@amber-aviation.com
P061,李庆宏,qhli@amber-aviation.com
P065,宋炜,wsong@amber-aviation.com
P068,昝昭君,zjzan@amber-aviation.com
P069,"ROEDER, SIMONE ELKE",simoneroeder@amber-aviation.com
P070,"Herve Daniel, STAMM",hdstamm@amber-aviation.com
P071,孙浩,jasonsun@amber-aviation.com
P072,朱正宇,zyzhu@amber-aviation.com
P074,金尚明,smjin@amber-aviation.com
P075,"Eduard Pascal, Roski",eduardroski@amber-aviation.com
P077,刘凯,andyliu@amber-aviation.com
P078,张帆,fzhang@amber-aviation.com
P079,魏思远,wesleywei@amber-aviation.com
P080,刘爽,sliu@amber-aviation.com
P081,吴鹏,richardwu@amber-aviation.com
P082,刘汇川,frankliu@amber-aviation.com
P083,尤欣,xyou@amber-aviation.com
P084,李亚民,ymli@amber-aviation.com
P085,赵镭,lzhao@amber-aviation.com
P086,张贺新,hxzhang@amber-aviation.com
P087,孙赫,hesun@amber-aviation.com
P088,马坚,harryma@amber-aviation.com
P089,李晓龙,xlli@amber-aviation.com
P090,黄海东,hdhuang@amber-aviation.com
P091,马洪双,mikema@amber-aviation.com
PJZ001,张哲,zzhang@amber-aviation.com
PJZ002,郭春旭,charlesguo@amber-aviation.com
PJZ004,王国勤,leowang@amber-aviation.com
PJZ005,王莹,evawang@amber-aviation.com
PJZ007,徐卓,frankxu@amber-aviation.com
PJZ008,杨华,ariayang@amber-aviation.com
W070,王彦海,wang_yanhai@163.com
W213,沈志伟,cshum@tagaviation.com
W267,"Nathon Andrew G, NORBERG",naten7@hotmail.com
W268,"Daniel, RICHTER",pilotlocalizer@gmail.com
W270,杨涛,yang_tao2005@aliyun.com
W272,"Andrew Nigel, KING",Andrew.king@aero.bombardier.com`;

        const CLICKED_KEY = 'clickedMails';
        const MAILS_KEY = 'wx_generated_mails_v1';
        const FLIGHT_DB_KEY = 'wx_flight_rows_v1';   // 航段表（新增）
        const PLAN_TEXT_KEY = 'wx_plan_text_v1';     // 文本计划（新增）

        let pilotMap = {};
        let flightRows = [];
        let clickedSet = new Set();
        let generatedMails = [];

        try {
            clickedSet = new Set(JSON.parse(localStorage.getItem(CLICKED_KEY) || '[]'));
        } catch (e) { clickedSet = new Set(); }

        function pad2(n) { return String(n).padStart(2, '0'); }

        function nowTimestamp() {
            const now = new Date();
            return now.getFullYear() + '-' + pad2(now.getMonth()+1) + '-' + pad2(now.getDate()) +
                   ' ' + pad2(now.getHours()) + ':' + pad2(now.getMinutes()) + ':' + pad2(now.getSeconds());
        }

        // ===== 存储封装 =====
        function saveMails() {
            try { localStorage.setItem(MAILS_KEY, JSON.stringify(generatedMails)); } catch (e) {}
        }
        function loadMails() {
            try {
                const raw = localStorage.getItem(MAILS_KEY);
                return raw ? JSON.parse(raw) : [];
            } catch (e) { return []; }
        }
        function saveFlightCache(fileName) {
            try {
                const payload = {
                    timestamp: nowTimestamp(),
                    filename: fileName,
                    rows: flightRows
                };
                localStorage.setItem(FLIGHT_DB_KEY, JSON.stringify(payload));
                return payload.timestamp;
            } catch (e) { console.error('保存航段缓存失败：', e); return null; }
        }
        function loadFlightCache() {
            try {
                const raw = localStorage.getItem(FLIGHT_DB_KEY);
                return raw ? JSON.parse(raw) : null;
            } catch (e) { return null; }
        }
        function savePlanText(text) {
            try { localStorage.setItem(PLAN_TEXT_KEY, text); } catch (e) {}
        }
        function loadPlanText() {
            try { return localStorage.getItem(PLAN_TEXT_KEY) || ''; } catch (e) { return ''; }
        }

        // ===== 内置名单解析 =====
        function parseCSV(text) {
            const rows = [];
            let row = [], current = '', inQuotes = false;
            for (let i = 0; i < text.length; i++) {
                const c = text[i];
                if (c === '"') {
                    if (inQuotes && text[i+1] === '"') { current += '"'; i++; }
                    else inQuotes = !inQuotes;
                } else if (c === ',' && !inQuotes) {
                    row.push(current); current = '';
                } else if ((c === '\n' || c === '\r') && !inQuotes) {
                    if (c === '\r' && text[i+1] === '\n') i++;
                    row.push(current); rows.push(row);
                    row = []; current = '';
                } else current += c;
            }
            if (current || row.length) { row.push(current); rows.push(row); }
            return rows;
        }

        function initPilotMap() {
            const rows = parseCSV(PILOT_CSV);
            const headers = rows[0].map(h => h.trim());
            const idIdx = headers.indexOf('工号');
            const emailIdx = headers.indexOf('邮箱');
            if (idIdx === -1 || emailIdx === -1) { alert('内置名单解析失败'); return; }
            pilotMap = {};
            for (let i = 1; i < rows.length; i++) {
                const r = rows[i];
                if (!r || r.length <= Math.max(idIdx, emailIdx)) continue;
                const id = r[idIdx].trim();
                const email = r[emailIdx].trim();
                if (id && email) pilotMap[id] = email;
            }
        }

        // ===== 文件读取 =====
        function readFile(file, callback) {
            const reader = new FileReader();
            const ext = file.name.split('.').pop().toLowerCase();
            if (ext === 'csv') {
                reader.onload = e => callback(e.target.result, 'csv');
                reader.readAsText(file, 'UTF-8');
            } else {
                reader.onload = e => callback(e.target.result, 'excel');
                reader.readAsArrayBuffer(file);
            }
        }

        function parseData(data, type) {
            if (type === 'csv') return parseCSV(data);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            return XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        }

        function normalizeDate(v) {
            if (v instanceof Date) {
                return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0') + '-' + String(v.getDate()).padStart(2,'0');
            }
            if (typeof v === 'number') {
                const date = new Date(Math.round((v - 25569) * 86400 * 1000));
                return date.getUTCFullYear() + '-' + String(date.getUTCMonth()+1).padStart(2,'0') + '-' + String(date.getUTCDate()).padStart(2,'0');
            }
            const s = String(v).trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
            return s;
        }

        function normalizeTime(v) {
            if (v instanceof Date) {
                return String(v.getHours()).padStart(2,'0') + ':' + String(v.getMinutes()).padStart(2,'0');
            }
            if (typeof v === 'number') {
                const totalMinutes = Math.round(v * 24 * 60);
                return String(Math.floor(totalMinutes / 60) % 24).padStart(2,'0') + ':' + String(totalMinutes % 60).padStart(2,'0');
            }
            const s = String(v).trim();
            const m = s.match(/(\d{1,2}):(\d{2})/);
            if (m) return m[1].padStart(2,'0') + ':' + m[2];
            return s;
        }

        // 把原始二维数组解析成 flightRows
        function buildFlightRows(rows) {
            let headerIdx = -1;
            for (let i = 0; i < rows.length; i++) {
                const r = rows[i].map(c => String(c).trim());
                if (r.includes('飞机注册号') && r.includes('出发地') && r.includes('计划出发')) { headerIdx = i; break; }
            }
            if (headerIdx === -1) return null;
            const headers = rows[headerIdx].map(h => String(h).trim());
            const col = {
                flight: headers.indexOf('飞机注册号'),
                depICAO: headers.indexOf('出发地'),
                arrICAO: headers.indexOf('到达地'),
                date: headers.indexOf('出发日期'),
                depTime: headers.indexOf('计划出发'),
                arrTime: headers.indexOf('预计到达')
            };
            if (Object.values(col).some(v => v === -1)) return null;
            const out = [];
            for (let i = headerIdx + 1; i < rows.length; i++) {
                const r = rows[i];
                if (!r || r.length <= col.flight) continue;
                const flight = String(r[col.flight] || '').trim();
                if (!flight) continue;
                out.push({
                    flight: flight,
                    depICAO: String(r[col.depICAO] || '').trim(),
                    arrICAO: String(r[col.arrICAO] || '').trim(),
                    date: normalizeDate(r[col.date]),
                    depTime: normalizeTime(r[col.depTime]),
                    arrTime: normalizeTime(r[col.arrTime])
                });
            }
            return out;
        }

        // ===== 上传文件 =====
        document.getElementById('flightFile').addEventListener('change', function(e) {
            const file = e.target.files[0]; if (!file) return;
            readFile(file, (data, type) => {
                const rows = parseData(data, type);
                const built = buildFlightRows(rows);
                if (built === null) {
                    document.getElementById('fileStatus').textContent = '❌ 解析失败：缺少必要列';
                    document.getElementById('fileStatus').style.color = '#d32f2f';
                    return;
                }
                flightRows = built;
                const ts = saveFlightCache(file.name);
                const statusEl = document.getElementById('fileStatus');
                statusEl.style.color = '#2e7d32';
                statusEl.textContent = '✅ 已加载 ' + built.length + ' 条（' + (ts || '') + '）';
                document.getElementById('cacheStatus').innerHTML = '';
            });
        });

        // ===== 文本计划：自动保存 =====
        let planTextTimer = null;
        document.getElementById('planText').addEventListener('input', function(e) {
            clearTimeout(planTextTimer);
            planTextTimer = setTimeout(() => {
                savePlanText(e.target.value);
            }, 500);
        });

        // ===== 计划解析 =====
        function parsePlan(text) {
            const lines = text.split(/\r?\n/);
            const result = [];
            let cur = null, expectRoute = false, expectCrew = false;
            const flightRegex = /^([A-Z0-9]+)\s+(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/;
            for (let line of lines) {
                line = line.trim(); if (!line) continue;
                const m = line.match(flightRegex);
                if (m) {
                    if (cur) result.push(cur);
                    cur = { flight: m[1], depTime: m[2], arrTime: m[3], crew: [] };
                    expectRoute = true; expectCrew = false;
                } else if (cur && expectRoute && line.includes('-')) {
                    expectRoute = false; expectCrew = true;
                } else if (cur && expectCrew && line.includes(',')) {
                    cur.crew = line.split(',').map(s => s.trim()).filter(s => s);
                    expectCrew = false;
                }
            }
            if (cur) result.push(cur);
            return result;
        }

        function convertDate(d) {
            const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
            const p = String(d).split('-');
            if (p.length !== 3) return '';
            return p[2].padStart(2,'0') + months[parseInt(p[1],10)-1];
        }

        function isExpired(date, depTime) {
            const depDateTime = new Date(date + 'T' + depTime + ':00');
            return new Date() >= depDateTime;
        }

        // ===== 生成邮件 =====
        function generate() {
            const outputDiv = document.getElementById('output');
            outputDiv.innerHTML = '';

            if (flightRows.length === 0) { outputDiv.innerHTML = '<p class="error">请先上传航段表</p>'; return; }
            const planText = document.getElementById('planText').value;
            const planFlights = parsePlan(planText);
            if (planFlights.length === 0) { outputDiv.innerHTML = '<p class="error">未解析到文本飞行计划</p>'; return; }

            const existingIds = new Set(generatedMails.map(m => m.mailId));
            let added = 0, skipped = 0;

            planFlights.forEach(pf => {
                const matches = flightRows.filter(f => f.flight === pf.flight && f.depTime === pf.depTime);
                const match = matches[0];
                if (!match) {
                    outputDiv.innerHTML += '<p class="error">航班 ' + pf.flight + ' ' + pf.depTime + ' 在航段表中未找到</p>';
                    return;
                }
                if (isExpired(match.date, pf.depTime)) return;

                const crewCodes = pf.crew.filter(c => /^(P|W|PJZ)/i.test(c));
                const recipients = crewCodes.slice(0, 2);
                if (recipients.length === 0) {
                    outputDiv.innerHTML += '<p class="error">航班 ' + pf.flight + ' 未找到飞行员</p>';
                    return;
                }
                const emails = recipients.map(id => pilotMap[id]).filter(e => e);
                if (emails.length === 0) {
                    outputDiv.innerHTML += '<p class="error">航班 ' + pf.flight + ' 的飞行员没有邮箱记录</p>';
                    return;
                }

                const dateText = convertDate(match.date);
                const subject = 'WX AND NOTAM ' + pf.flight + ' ' + match.depICAO + '-' + match.arrICAO + ' ' + dateText;
                const mailto = 'mailto:' + emails.join(';') + '?subject=' + encodeURIComponent(subject);

                const mailId = pf.flight + '_' + match.date + '_' + pf.depTime;
                if (existingIds.has(mailId)) { skipped++; return; }

                generatedMails.push({
                    mailId: mailId,
                    depDateTime: match.date + 'T' + pf.depTime + ':00',
                    mailto: mailto,
                    text: pf.flight + ' ' + match.depICAO + '-' + match.arrICAO + ' ' + pf.depTime + '-' + pf.arrTime + ' → ' + recipients.join(', ')
                });
                existingIds.add(mailId);
                added++;
            });

            saveMails();
            renderMails();

            if (added > 0) {
                outputDiv.innerHTML = '<p class="success">✅ 新增 ' + added + ' 条邮件' + (skipped > 0 ? '（' + skipped + ' 条已存在，自动跳过）' : '') + '</p>';
            } else if (skipped > 0) {
                outputDiv.innerHTML = '<p class="info">ℹ️ 所有邮件都已存在列表中，未新增</p>';
            }
        }

        // ===== 渲染列表 =====
        function renderMails() {
            const container = document.getElementById('mailList');
            container.innerHTML = '';

            const now = new Date();
            const alive = [];
            for (const m of generatedMails) {
                const depDt = new Date(m.depDateTime);
                if (now >= depDt) continue;
                alive.push(m);
            }
            if (alive.length !== generatedMails.length) {
                generatedMails = alive;
                saveMails();
            }

            if (generatedMails.length === 0) {
                container.innerHTML = '<div class="info">📭 当前没有待处理的邮件。上传航段表并粘贴计划后点击「生成邮件链接」。</div>';
                return;
            }

            for (const m of generatedMails) {
                const row = document.createElement('div');
                row.className = 'mail-row';

                const link = document.createElement('a');
                link.href = m.mailto;
                link.className = 'mail-link';
                link.textContent = m.text;
                link.title = m.text;

                const depDt = new Date(m.depDateTime);
                const threeH = new Date(depDt.getTime() - 3 * 60 * 60 * 1000);
                if (clickedSet.has(m.mailId)) {
                    link.style.backgroundColor = '#c8e6c9';
                    link.style.borderColor = '#4caf50';
                } else if (now >= threeH) {
                    link.style.backgroundColor = '#fff9c4';
                    link.style.borderColor = '#fbc02d';
                } else {
                    link.style.backgroundColor = '#f0f0f0';
                    link.style.borderColor = '#ccc';
                }

                link.addEventListener('click', function() {
                    clickedSet.add(m.mailId);
                    try { localStorage.setItem(CLICKED_KEY, JSON.stringify([...clickedSet])); } catch(e) {}
                    setTimeout(renderMails, 100);
                });

                row.appendChild(link);

                const del = document.createElement('button');
                del.className = 'del-btn';
                del.textContent = '✕';
                del.title = '删除此条邮件';
                del.onclick = () => {
                    generatedMails = generatedMails.filter(x => x.mailId !== m.mailId);
                    saveMails();
                    renderMails();
                };
                row.appendChild(del);

                container.appendChild(row);
            }
        }

        function clearAllMails() {
            if (!confirm('确定清空所有待处理邮件吗？')) return;
            generatedMails = [];
            saveMails();
            renderMails();
        }

        setInterval(() => {
            renderMails();
        }, 60000);

        // ===== 页面加载：恢复上次状态 =====
        window.addEventListener('DOMContentLoaded', () => {
            initPilotMap();
            generatedMails = loadMails();
            renderMails();

            // 1) 恢复航段表
            const flightCache = loadFlightCache();
            if (flightCache && flightCache.rows && flightCache.rows.length > 0) {
                flightRows = flightCache.rows;
                document.getElementById('cacheStatus').innerHTML =
                    '<div class="info">💾 已恢复上次上传：' +
                    escapeHtml(flightCache.filename || '未知文件') +
                    '（' + escapeHtml(flightCache.timestamp || '') + '），共 ' +
                    flightCache.rows.length + ' 条航段</div>';
                document.getElementById('fileStatus').textContent =
                    '✅ 已加载 ' + flightCache.rows.length + ' 条';
                document.getElementById('fileStatus').style.color = '#2e7d32';
            }

            // 2) 恢复文本计划
            const savedPlan = loadPlanText();
            if (savedPlan) {
                document.getElementById('planText').value = savedPlan;
            }
        });

        function escapeHtml(s) {
            return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }
    </script>
</body>
</html>
"""
    components.html(WX_HTML, height=1200, scrolling=True)

# ==============================
# 主界面
# ==============================
st.set_page_config(page_title="监控+计划", layout="wide")

# ===== Tab 栏固定顶部 + 前3个一行、后5个一行，左对齐 =====
st.markdown("""
<style>
/* Tab 栏固定在顶部 */
div[data-testid="stTabs"] > div:first-child {
    position: sticky;
    top: 0;
    z-index: 999999;
    background: white;
    padding-top: 8px;
    padding-bottom: 8px;
    border-bottom: 1px solid #eee;
}
/* Tab 容器：换行 + 左对齐 */
div[data-testid="stTabs"] div[role="tablist"] {
    display: flex !important;
    flex-wrap: wrap !important;
    justify-content: flex-start !important;
    gap: 4px !important;
    overflow-x: visible !important;
    white-space: normal !important;
}
/* Tab 按内容宽度自适应，不拉伸 */
div[data-testid="stTabs"] div[role="tablist"] > * {
    flex: 0 0 auto !important;
    margin-bottom: 4px !important;
}
/* 第3个 Tab 之后强制换行（因为去掉了世界时，只剩前3个） */
div[data-testid="stTabs"] div[role="tablist"] > *:nth-child(3) {
    margin-right: 100% !important;
}
</style>
""", unsafe_allow_html=True)

st.title("监控+计划")

tabs = st.tabs([
    "✉️ WX邮件生成",
    "📝 检查单&填写脚本",
    "✈️ 航路处理",
    "📋 每日飞行数据-10：00发",
    "📄 每日通航运行情况跟踪表-16：30发",
    "📊 天成商务航空每日运行跟踪-16：30发",
    "📜 通航脚本",
    "📊 值班连班统计"
])

with tabs[0]:
    run_feature_wx_mail()
with tabs[1]:
    run_feature_chk()
with tabs[2]:
    run_feature_g()
with tabs[3]:
    run_feature_a()
with tabs[4]:
    run_feature_b()
with tabs[5]:
    run_feature_c()
with tabs[6]:
    run_feature_d()
with tabs[7]:
    run_feature_e()

st.caption("💡 邮件生成每天多次用，检查单与世界时行程已合并（自动填表），航路处理常驻第3位；其余功能保持不变。")
