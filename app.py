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
# 功能 D：通航脚本生成器（ICAO 维表版）
# 只替换本段；其余功能不动
# ==============================
COUNTRIES = [
    "香港", "澳门", "台湾", "蒙古", "朝鲜", "韩国", "日本", "菲律宾", "越南", "老挝",
    "柬埔寨", "缅甸", "泰国", "马来西亚", "文莱", "新加坡", "印度尼西亚", "东帝汶",
    "尼泊尔", "不丹", "孟加拉国", "印度", "巴基斯坦", "斯里兰卡", "马尔代夫",
    "哈萨克斯坦", "吉尔吉斯斯坦", "塔吉克斯坦", "乌兹别克斯坦", "土库曼斯坦",
    "阿富汗", "伊拉克", "伊朗", "叙利亚", "约旦", "黎巴嫩", "以色列", "巴勒斯坦",
    "沙特阿拉伯", "巴林", "卡塔尔", "科威特", "阿联酋", "阿曼", "也门", "格鲁吉亚",
    "亚美尼亚", "阿塞拜疆", "土耳其", "塞浦路斯", "芬兰", "瑞典", "挪威", "冰岛",
    "丹麦", "法罗群岛（丹）", "爱沙尼亚", "拉脱维亚", "立陶宛", "白俄罗斯", "俄罗斯",
    "乌克兰", "摩尔多瓦", "波兰", "捷克", "斯洛伐克", "匈牙利", "德国", "奥地利",
    "瑞士", "列支敦士登", "英国", "爱尔兰", "荷兰", "比利时", "卢森堡", "法国",
    "摩纳哥", "罗马尼亚", "保加利亚", "塞尔维亚", "马其顿", "阿尔巴尼亚", "希腊",
    "斯洛文尼亚", "克罗地亚", "波斯尼亚和墨塞哥维那", "意大利", "梵蒂冈", "圣马力诺",
    "马耳他", "西班牙", "葡萄牙", "安道尔", "埃及", "利比亚", "苏丹", "突尼斯",
    "阿尔及利亚", "摩洛哥", "亚速尔群岛（葡）", "马德拉群岛（葡）", "埃塞俄比亚",
    "厄立特里亚", "索马里", "吉布提", "肯尼亚", "坦桑尼亚", "乌干达", "卢旺达",
    "布隆迪", "塞舌尔", "乍得", "中非", "喀麦隆", "赤道几内亚", "加蓬", "刚果（布）",
    "刚果（金）", "圣多美及普林西比", "毛里塔尼亚", "西撒哈拉", "塞内加尔", "冈比亚",
    "马里", "布基纳法索", "几内亚", "几内亚比绍", "佛得角", "塞拉利昂", "利比里亚",
    "科特迪瓦", "加纳", "多哥", "贝宁", "尼日尔", "加那利群岛（西）", "赞比亚",
    "安哥拉", "津巴布韦", "马拉维", "莫桑比克", "博茨瓦纳", "纳米比亚", "南非",
    "斯威士兰", "莱索托", "马达加斯加", "科摩罗", "毛里求斯", "留尼旺（法）",
    "圣赫勒拿（英）", "澳大利亚", "新西兰", "巴布亚新几内亚", "所罗门群岛",
    "瓦努阿图", "密克罗尼西亚", "马绍尔群岛", "帕劳", "瑙鲁", "基里巴斯", "图瓦卢",
    "萨摩亚", "斐济群岛", "汤加", "库克群岛（新）", "关岛（美）", "新喀里多尼亚（法）",
    "法属波利尼西亚", "皮特凯恩岛（英）", "瓦利斯与富图纳（法）", "纽埃（新）",
    "托克劳（新）", "美属萨摩亚", "北马里亚纳（美）", "加拿大", "美国", "墨西哥",
    "格陵兰（丹）", "危地马拉", "伯利兹", "萨尔瓦多", "洪都拉斯", "尼加拉瓜",
    "哥斯达黎加", "巴拿马", "巴哈马", "古巴", "牙买加", "海地", "多米尼加共和国",
    "安提瓜和巴布达", "圣基茨和尼维斯", "多米尼克", "圣卢西亚", "圣文森特和格林纳丁斯",
    "格林纳达", "巴巴多斯", "特立尼达和多巴哥", "波多黎各（美）", "英属维尔京群岛",
    "美属维尔京群岛", "安圭拉（英）", "蒙特塞拉特（英）", "瓜德罗普（法）",
    "马提尼克（法）", "荷属安的列斯", "阿鲁巴（荷）", "特克斯和凯科斯群岛（英）",
    "开曼群岛（英）", "百慕大（英）", "哥伦比亚", "委内瑞拉", "圭亚那", "法属圭亚那",
    "苏里南", "厄瓜多尔", "秘鲁", "玻利维亚", "智利", "阿根廷", "乌拉圭", "巴拉圭",
    "其他"
]

ABBR_TO_COUNTRY = {
    "印尼": "印度尼西亚", "台北": "台湾", "高雄": "台湾", "台中": "台湾",
    "花莲": "台湾", "台东": "台湾", "嘉义": "台湾", "台南": "台湾",
}

TAIWAN_CITIES = [
    "台北松山", "台北桃园", "高雄小港", "台中清泉岗", "花莲", "台东", "嘉义", "台南", "马公", "金门", "马祖"
]

# ============================================================
# 机场维表：ICAO → (省份, 机场简称, 机场全名, [第三级候选…])
# 境内一律按 ICAO 认定，不再靠中文城市名做白名单
# ============================================================
AIRPORT_TABLE = {
    # ---- 上海 ----
    "ZSPD": ("上海", "浦东", "上海浦东机场", ["浦东新区", "浦东"]),
    "ZSSS": ("上海", "虹桥", "上海虹桥机场", ["闵行区", "虹桥"]),
    # ---- 云南 ----
    "ZPBS": ("云南", "云瑞", "保山云瑞机场", ["保山", "云瑞"]),
    "ZPCW": ("云南", "佤山", "沧源佤山机场", ["沧源", "临沧", "佤山"]),
    "ZPDL": ("云南", "凤仪", "大理凤仪机场", ["大理", "凤仪"]),
    "ZPDQ": ("云南", "香格里拉", "迪庆香格里拉机场", ["迪庆", "香格里拉"]),
    "ZPJH": ("云南", "嘎洒", "西双版纳嘎洒机场", ["西双版纳", "嘎洒"]),
    "ZPJM": ("云南", "景迈", "澜沧景迈机场", ["澜沧", "普洱", "景迈"]),
    "ZPLC": ("云南", "博尚", "临沧博尚机场", ["临沧", "博尚"]),
    "ZPLJ": ("云南", "三义", "丽江三义机场", ["丽江", "三义"]),
    "ZPMS": ("云南", "芒市", "德宏芒市机场", ["德宏", "芒市"]),
    "ZPNL": ("云南", "泸沽湖", "宁蒗泸沽湖机场", ["宁蒗", "丽江", "泸沽湖"]),
    "ZPPP": ("云南", "长水", "昆明长水机场", ["昆明", "长水"]),
    "ZPSM": ("云南", "思茅", "普洱思茅机场", ["普洱", "思茅"]),
    "ZPTC": ("云南", "驼峰", "腾冲驼峰机场", ["腾冲", "保山", "驼峰"]),
    "ZPWS": ("云南", "砚山", "文山砚山机场", ["文山", "砚山"]),
    "ZPZT": ("云南", "", "昭通机场", ["昭通"]),
    # ---- 内蒙古 ----
    "ZBAL": ("内蒙古", "巴彦浩特", "阿拉善左旗巴彦浩特机场", ["阿拉善左旗", "阿拉善", "巴彦浩特"]),
    "ZBAR": ("内蒙古", "巴丹吉林", "阿拉善右旗巴丹吉林机场", ["阿拉善右旗", "阿拉善", "巴丹吉林"]),
    "ZBCF": ("内蒙古", "玉龙", "赤峰玉龙机场", ["赤峰", "玉龙"]),
    "ZBDS": ("内蒙古", "伊霍洛", "鄂尔多斯伊霍洛机场", ["鄂尔多斯", "伊霍洛"]),
    "ZBEN": ("内蒙古", "桃来", "额济纳旗桃来机场", ["额济纳旗", "阿拉善", "桃来"]),
    "ZBER": ("内蒙古", "赛乌素", "二连浩特赛乌素机场", ["二连浩特", "锡林郭勒", "赛乌素"]),
    "ZBES": ("内蒙古", "伊尔施", "阿尔山伊尔施机场", ["阿尔山", "兴安盟", "伊尔施"]),
    "ZBHH": ("内蒙古", "白塔", "呼和浩特白塔机场", ["呼和浩特", "白塔"]),
    "ZBHZ": ("内蒙古", "霍林河", "霍林郭勒霍林河机场", ["霍林郭勒", "通辽", "霍林河"]),
    "ZBLA": ("内蒙古", "海拉尔", "呼伦贝尔海拉尔机场", ["呼伦贝尔", "海拉尔"]),
    "ZBMZ": ("内蒙古", "西郊", "满洲里西郊机场", ["满洲里", "呼伦贝尔", "西郊"]),
    "ZBOW": ("内蒙古", "东河", "包头东河机场", ["包头", "东河"]),
    "ZBTL": ("内蒙古", "", "通辽机场", ["通辽"]),
    "ZBUC": ("内蒙古", "集宁", "乌兰察布集宁机场", ["乌兰察布", "集宁"]),
    "ZBUH": ("内蒙古", "", "乌海机场", ["乌海"]),
    "ZBUL": ("内蒙古", "义勒力特", "乌兰浩特义勒力特机场", ["乌兰浩特", "兴安盟", "义勒力特"]),
    "ZBXH": ("内蒙古", "", "锡林浩特机场", ["锡林浩特"]),
    "ZBYZ": ("内蒙古", "天吉泰", "巴彦淖尔天吉泰机场", ["巴彦淖尔", "天吉泰"]),
    "ZBZL": ("内蒙古", "成吉思汗", "扎兰屯成吉思汗机场", ["扎兰屯", "呼伦贝尔", "成吉思汗"]),
    # ---- 北京 ----
    "ZBAA": ("北京", "首都", "北京首都机场", ["顺义区", "首都"]),
    "ZBAD": ("北京", "大兴", "北京大兴机场", ["大兴区", "大兴"]),
    # ---- 吉林 ----
    "ZYBA": ("吉林", "长安", "白城长安机场", ["白城", "长安"]),
    "ZYBS": ("吉林", "长白山", "白山长白山机场", ["白山", "长白山"]),
    "ZYCC": ("吉林", "龙嘉", "长春龙嘉机场", ["长春", "龙嘉"]),
    "ZYSQ": ("吉林", "查干湖", "松原查干湖机场", ["松原", "查干湖"]),
    "ZYTN": ("吉林", "三源浦", "通化三源浦机场", ["通化", "三源浦"]),
    "ZYYJ": ("吉林", "朝阳川", "延吉朝阳川机场", ["延吉", "延边", "朝阳川"]),
    # ---- 四川 ----
    "ZUBZ": ("四川", "恩阳", "巴中恩阳机场", ["巴中", "恩阳"]),
    "ZUDA": ("四川", "金垭", "达州金垭机场", ["达州", "金垭"]),
    "ZUDC": ("四川", "亚丁", "稻城亚丁机场", ["稻城", "甘孜", "亚丁"]),
    "ZUGH": ("四川", "", "广汉机场", ["广汉", "德阳"]),
    "ZUGU": ("四川", "盘龙", "广元盘龙机场", ["广元", "盘龙"]),
    "ZUGZ": ("四川", "格萨尔", "甘孜格萨尔机场", ["甘孜", "格萨尔"]),
    "ZUHY": ("四川", "红原", "阿坝红原机场", ["阿坝", "红原"]),
    "ZUJZ": ("四川", "黄龙", "九寨黄龙机场", ["九寨", "阿坝", "黄龙"]),
    "ZUKD": ("四川", "", "康定机场", ["康定", "甘孜"]),
    "ZULA": ("四川", "古城", "阆中古城机场", ["阆中", "南充", "古城"]),
    "ZULF": ("四川", "", "乐山机场", ["乐山"]),
    "ZULZ": ("四川", "云龙", "泸州云龙机场", ["泸州", "云龙"]),
    "ZUMY": ("四川", "南郊", "绵阳南郊机场", ["绵阳", "南郊"]),
    "ZUNC": ("四川", "高坪", "南充高坪机场", ["南充", "高坪"]),
    "ZUTF": ("四川", "天府", "成都天府机场", ["成都", "天府"]),
    "ZUUU": ("四川", "双流", "成都双流机场", ["成都", "双流"]),
    "ZUXC": ("四川", "青山", "西昌青山机场", ["西昌", "凉山", "青山"]),
    "ZUXJ": ("四川", "", "新津机场", ["新津", "成都"]),
    "ZUYB": ("四川", "五粮液", "宜宾五粮液机场", ["宜宾", "五粮液"]),
    "ZUZH": ("四川", "保安营", "攀枝花保安营机场", ["攀枝花", "保安营"]),
    # ---- 天津 ----
    "ZBTJ": ("天津", "滨海", "天津滨海机场", ["滨海新区", "滨海"]),
    # ---- 宁夏 ----
    "ZLGY": ("宁夏", "六盘山", "固原六盘山机场", ["固原", "六盘山"]),
    "ZLIC": ("宁夏", "河东", "银川河东机场", ["银川", "河东"]),
    "ZLZW": ("宁夏", "沙坡头", "中卫沙坡头机场", ["中卫", "沙坡头"]),
    # ---- 安徽 ----
    "ZSAQ": ("安徽", "", "安庆机场", ["安庆"]),
    "ZSBA": ("安徽", "滕湖", "蚌埠滕湖机场", ["蚌埠", "滕湖"]),
    "ZSBO": ("安徽", "", "亳州机场", ["亳州"]),
    "ZSFY": ("安徽", "", "阜阳机场", ["阜阳"]),
    "ZSJH": ("安徽", "九华山", "池州九华山机场", ["池州", "九华山"]),
    "ZSOF": ("安徽", "新桥", "合肥新桥机场", ["合肥", "新桥"]),
    "ZSTX": ("安徽", "屯溪", "黄山屯溪机场", ["黄山", "屯溪"]),
    "ZSWA": ("安徽", "宣州", "芜湖宣州机场", ["芜湖", "宣州"]),
    # ---- 山东 ----
    "ZSDY": ("山东", "胜利", "东营胜利机场", ["东营", "胜利"]),
    "ZSHZ": ("山东", "牡丹", "菏泽牡丹机场", ["菏泽", "牡丹"]),
    "ZSJG": ("山东", "大安", "济宁大安机场", ["济宁", "大安"]),
    "ZSJN": ("山东", "遥墙", "济南遥墙机场", ["济南", "遥墙"]),
    "ZSLY": ("山东", "启阳", "临沂启阳机场", ["临沂", "启阳"]),
    "ZSQD": ("山东", "胶东", "青岛胶东机场", ["青岛", "胶东"]),
    "ZSRZ": ("山东", "山字河", "日照山字河机场", ["日照", "山字河"]),
    "ZSWF": ("山东", "", "潍坊机场", ["潍坊"]),
    "ZSWH": ("山东", "大水泊", "威海大水泊机场", ["威海", "大水泊"]),
    "ZSYT": ("山东", "蓬莱", "烟台蓬莱机场", ["烟台", "蓬莱"]),
    # ---- 山西 ----
    "ZBCZ": ("山西", "王村", "长治王村机场", ["长治", "王村"]),
    "ZBDT": ("山西", "云冈", "大同云冈机场", ["大同", "云冈"]),
    "ZBLF": ("山西", "尧都", "临汾尧都机场", ["临汾", "尧都"]),
    "ZBLL": ("山西", "大武", "吕梁大武机场", ["吕梁", "大武"]),
    "ZBSG": ("山西", "滋润", "朔州滋润机场", ["朔州", "滋润"]),
    "ZBXZ": ("山西", "五台山", "忻州五台山机场", ["忻州", "五台山"]),
    "ZBYC": ("山西", "盐湖", "运城盐湖机场", ["运城", "盐湖"]),
    "ZBYN": ("山西", "武宿", "太原武宿机场", ["太原", "武宿"]),
    # ---- 广东 ----
    "ZGGG": ("广东", "白云", "广州白云机场", ["广州", "白云"]),
    "ZGHZ": ("广东", "平潭", "惠州平潭机场", ["惠州", "平潭"]),
    "ZGLD": ("广东", "", "罗定机场", ["罗定", "云浮"]),
    "ZGMX": ("广东", "梅县", "梅州梅县机场", ["梅州", "梅县"]),
    "ZGOW": ("广东", "潮汕", "揭阳潮汕机场", ["揭阳", "潮汕"]),
    "ZGSD": ("广东", "金湾", "珠海金湾机场", ["珠海", "金湾"]),
    "ZGSG": ("广东", "丹霞", "韶关丹霞机场", ["韶关", "丹霞"]),
    "ZGSZ": ("广东", "宝安", "深圳宝安机场", ["深圳", "宝安"]),
    "ZGUH": ("广东", "九洲", "珠海九洲机场", ["珠海", "九洲"]),
    "ZGYJ": ("广东", "合山", "阳江合山机场", ["阳江", "合山"]),
    "ZGZJ": ("广东", "吴川", "湛江吴川机场", ["湛江", "吴川"]),
    # ---- 广西 ----
    "ZGBH": ("广西", "福成", "北海福成机场", ["北海", "福成"]),
    "ZGBS": ("广西", "巴马", "百色巴马机场", ["百色", "巴马", "田阳"]),
    "ZGHC": ("广西", "金城江", "河池金城江机场", ["河池", "金城江"]),
    "ZGKL": ("广西", "两江", "桂林两江机场", ["桂林", "两江"]),
    "ZGNN": ("广西", "吴圩", "南宁吴圩机场", ["南宁", "吴圩"]),
    "ZGWZ": ("广西", "西江", "梧州西江机场", ["梧州", "西江"]),
    "ZGYL": ("广西", "福绵", "玉林福绵机场", ["玉林", "福绵"]),
    "ZGZH": ("广西", "白莲", "柳州白莲机场", ["柳州", "白莲"]),
    # ---- 新疆 ----
    "ZWAK": ("新疆", "红旗坡", "阿克苏红旗坡机场", ["阿克苏", "红旗坡"]),
    "ZWAL": ("新疆", "塔里木", "阿拉尔塔里木机场", ["阿拉尔", "塔里木"]),
    "ZWAT": ("新疆", "雪都", "阿勒泰雪都机场", ["阿勒泰", "雪都"]),
    "ZWBL": ("新疆", "阿拉山口", "博乐阿拉山口机场", ["博乐", "阿拉山口"]),
    "ZWCM": ("新疆", "玉都", "且末玉都机场", ["且末", "巴音郭楞", "玉都"]),
    "ZWFY": ("新疆", "可可托海", "富蕴可可托海机场", ["富蕴", "阿勒泰", "可可托海"]),
    "ZWHJ": ("新疆", "巴音布鲁克", "和静巴音布鲁克机场", ["和静", "巴音郭楞", "巴音布鲁克"]),
    "ZWHM": ("新疆", "伊州", "哈密伊州机场", ["哈密", "伊州"]),
    "ZWHZ": ("新疆", "花园", "石河子花园机场", ["石河子", "花园"]),
    "ZWKC": ("新疆", "龟兹", "库车龟兹机场", ["库车", "阿克苏", "龟兹"]),
    "ZWKL": ("新疆", "梨城", "库尔勒梨城机场", ["库尔勒", "梨城"]),
    "ZWKM": ("新疆", "古海", "克拉玛依古海机场", ["克拉玛依", "古海"]),
    "ZWKN": ("新疆", "喀纳斯", "布尔津喀纳斯机场", ["布尔津", "阿勒泰", "喀纳斯"]),
    "ZWLK": ("新疆", "大河", "巴里坤大河机场", ["巴里坤", "哈密", "大河"]),
    "ZWLT": ("新疆", "轮台", "巴音郭楞轮台机场", ["巴音郭楞", "轮台"]),
    "ZWNL": ("新疆", "那拉提", "新源那拉提机场", ["新源", "伊犁", "那拉提"]),
    "ZWQT": ("新疆", "江布拉克", "奇台江布拉克机场", ["奇台", "昌吉", "江布拉克"]),
    "ZWRQ": ("新疆", "楼兰", "若羌楼兰机场", ["若羌", "巴音郭楞", "楼兰"]),
    "ZWSC": ("新疆", "叶尔羌", "莎车叶尔羌机场", ["莎车", "喀什", "叶尔羌"]),
    "ZWSH": ("新疆", "徕宁", "喀什徕宁机场", ["喀什", "徕宁"]),
    "ZWTC": ("新疆", "千泉", "塔城千泉机场", ["塔城", "千泉"]),
    "ZWTK": ("新疆", "红其拉甫", "塔什库尔干红其拉甫机场", ["塔什库尔干", "喀什", "红其拉甫"]),
    "ZWTL": ("新疆", "交河", "吐鲁番交河机场", ["吐鲁番", "交河"]),
    "ZWTN": ("新疆", "昆冈", "和田昆冈机场", ["和田", "昆冈"]),
    "ZWTS": ("新疆", "唐王城", "图木舒克唐王城机场", ["图木舒克", "唐王城"]),
    "ZWWW": ("新疆", "天山", "乌鲁木齐天山机场", ["乌鲁木齐", "天山"]),
    "ZWYN": ("新疆", "伊宁", "伊犁伊宁机场", ["伊犁", "伊宁"]),
    "ZWYT": ("新疆", "万方", "于田万方机场", ["于田", "万方"]),
    "ZWZS": ("新疆", "天马", "昭苏天马机场", ["昭苏", "伊犁", "天马"]),
    # ---- 江苏 ----
    "ZSCG": ("江苏", "奔牛", "常州奔牛机场", ["常州", "奔牛"]),
    "ZSLG": ("江苏", "花果山", "连云港花果山机场", ["连云港", "花果山"]),
    "ZSNJ": ("江苏", "禄口", "南京禄口机场", ["南京", "禄口"]),
    "ZSNT": ("江苏", "兴东", "南通兴东机场", ["南通", "兴东"]),
    "ZSSH": ("江苏", "涟水", "淮安涟水机场", ["淮安", "涟水"]),
    "ZSWX": ("江苏", "硕放", "无锡硕放机场", ["无锡", "硕放"]),
    "ZSXZ": ("江苏", "观音", "徐州观音机场", ["徐州", "观音"]),
    "ZSYA": ("江苏", "泰州", "扬州泰州机场", ["扬州", "泰州"]),
    "ZSYN": ("江苏", "南洋", "盐城南洋机场", ["盐城", "南洋"]),
    # ---- 江西 ----
    "ZSCN": ("江西", "昌北", "南昌昌北机场", ["南昌", "昌北"]),
    "ZSGS": ("江西", "", "井冈山机场", ["井冈山", "吉安"]),
    "ZSGZ": ("江西", "黄金", "赣州黄金机场", ["赣州", "黄金"]),
    "ZSJD": ("江西", "罗家", "景德镇罗家机场", ["景德镇", "罗家"]),
    "ZSJJ": ("江西", "庐山", "九江庐山机场", ["九江", "庐山"]),
    "ZSRJ": ("江西", "瑞金", "赣州瑞金机场", ["赣州", "瑞金"]),
    "ZSSR": ("江西", "三清山", "上饶三清山机场", ["上饶", "三清山"]),
    "ZSYC": ("江西", "明月山", "宜春明月山机场", ["宜春", "明月山"]),
    # ---- 河北 ----
    "ZBCD": ("河北", "普宁", "承德普宁机场", ["承德", "普宁"]),
    "ZBDH": ("河北", "北戴河", "秦皇岛北戴河机场", ["秦皇岛", "北戴河"]),
    "ZBHD": ("河北", "", "邯郸机场", ["邯郸"]),
    "ZBSJ": ("河北", "正定", "石家庄正定机场", ["石家庄", "正定"]),
    "ZBSN": ("河北", "三女河", "唐山三女河机场", ["唐山", "三女河"]),
    "ZBXT": ("河北", "褡裢", "邢台褡裢机场", ["邢台", "褡裢"]),
    "ZBZJ": ("河北", "宁远", "张家口宁远机场", ["张家口", "宁远"]),
    # ---- 河南 ----
    "ZHCC": ("河南", "新郑", "郑州新郑机场", ["郑州", "新郑"]),
    "ZHLY": ("河南", "", "洛阳机场", ["洛阳"]),
    "ZHNY": ("河南", "姜营", "南阳姜营机场", ["南阳", "姜营"]),
    "ZHQQ": ("河南", "红旗渠", "安阳红旗渠机场", ["安阳", "红旗渠"]),
    "ZHXY": ("河南", "明港", "信阳明港机场", ["信阳", "明港"]),
    # ---- 浙江 ----
    "ZSHC": ("浙江", "萧山", "杭州萧山机场", ["杭州", "萧山"]),
    "ZSJU": ("浙江", "", "衢州机场", ["衢州"]),
    "ZSJX": ("浙江", "南湖", "嘉兴南湖机场", ["嘉兴", "南湖"]),
    "ZSLI": ("浙江", "", "丽水机场", ["丽水"]),
    "ZSLQ": ("浙江", "路桥", "台州路桥机场", ["台州", "路桥"]),
    "ZSNB": ("浙江", "栎社", "宁波栎社机场", ["宁波", "栎社"]),
    "ZSWZ": ("浙江", "龙湾", "温州龙湾机场", ["温州", "龙湾"]),
    "ZSYW": ("浙江", "", "义乌机场", ["义乌", "金华"]),
    "ZSZS": ("浙江", "普陀山", "舟山普陀山机场", ["舟山", "普陀山"]),
    # ---- 海南 ----
    "ZJHK": ("海南", "美兰", "海口美兰机场", ["海口", "美兰"]),
    "ZJQH": ("海南", "博鳌", "琼海博鳌机场", ["琼海", "博鳌"]),
    "ZJSY": ("海南", "凤凰", "三亚凤凰机场", ["三亚", "凤凰"]),
    "ZJYX": ("海南", "永兴", "三沙永兴机场", ["三沙", "永兴"]),
    # ---- 湖北 ----
    "ZHEC": ("湖北", "花湖", "鄂州花湖机场", ["鄂州", "花湖"]),
    "ZHES": ("湖北", "许家坪", "恩施许家坪机场", ["恩施", "许家坪"]),
    "ZHHH": ("湖北", "天河", "武汉天河机场", ["武汉", "天河"]),
    "ZHJZ": ("湖北", "沙市", "荆州沙市机场", ["荆州", "沙市"]),
    "ZHSN": ("湖北", "红坪", "神农架红坪机场", ["神农架", "红坪"]),
    "ZHSY": ("湖北", "武当山", "十堰武当山机场", ["十堰", "武当山"]),
    "ZHXF": ("湖北", "刘集", "襄阳刘集机场", ["襄阳", "刘集"]),
    "ZHYC": ("湖北", "三峡", "宜昌三峡机场", ["宜昌", "三峡"]),
    # ---- 湖南 ----
    "ZGCD": ("湖南", "桃花源", "常德桃花源机场", ["常德", "桃花源"]),
    "ZGCJ": ("湖南", "芷江", "怀化芷江机场", ["怀化", "芷江"]),
    "ZGCZ": ("湖南", "北湖", "郴州北湖机场", ["郴州", "北湖"]),
    "ZGDY": ("湖南", "荷花", "张家界荷花机场", ["张家界", "荷花"]),
    "ZGHA": ("湖南", "黄花", "长沙黄花机场", ["长沙", "黄花"]),
    "ZGHY": ("湖南", "南岳", "衡阳南岳机场", ["衡阳", "南岳"]),
    "ZGLG": ("湖南", "零陵", "永州零陵机场", ["永州", "零陵"]),
    "ZGSY": ("湖南", "武冈", "邵阳武冈机场", ["邵阳", "武冈"]),
    "ZGXX": ("湖南", "边城", "湘西边城机场", ["湘西", "边城"]),
    "ZGYY": ("湖南", "三荷", "岳阳三荷机场", ["岳阳", "三荷"]),
    # ---- 甘肃 ----
    "ZLDH": ("甘肃", "莫高", "敦煌莫高机场", ["敦煌", "酒泉", "莫高"]),
    "ZLJC": ("甘肃", "金川", "金昌金川机场", ["金昌", "金川"]),
    "ZLJQ": ("甘肃", "酒泉", "嘉峪关酒泉机场", ["嘉峪关", "酒泉"]),
    "ZLLL": ("甘肃", "中川", "兰州中川机场", ["兰州", "中川"]),
    "ZLLN": ("甘肃", "成县", "陇南成县机场", ["陇南", "成县"]),
    "ZLQY": ("甘肃", "西峰", "庆阳西峰机场", ["庆阳", "西峰"]),
    "ZLTS": ("甘肃", "麦积山", "天水麦积山机场", ["天水", "麦积山"]),
    "ZLXH": ("甘肃", "夏河", "甘南夏河机场", ["甘南", "夏河"]),
    "ZLZY": ("甘肃", "甘州", "张掖甘州机场", ["张掖", "甘州"]),
    # ---- 福建 ----
    "ZSAM": ("福建", "高崎", "厦门高崎机场", ["厦门", "高崎"]),
    "ZSFZ": ("福建", "长乐", "福州长乐机场", ["福州", "长乐"]),
    "ZSLO": ("福建", "冠豸山", "连城冠豸山机场", ["连城", "龙岩", "冠豸山"]),
    "ZSQZ": ("福建", "晋江", "泉州晋江机场", ["泉州", "晋江"]),
    "ZSSM": ("福建", "沙县", "三明沙县机场", ["三明", "沙县"]),
    "ZSWY": ("福建", "", "武夷山机场", ["武夷山", "南平"]),
    # ---- 西藏 ----
    "ZUAL": ("西藏", "昆莎", "阿里昆莎机场", ["阿里", "昆莎"]),
    "ZUBD": ("西藏", "邦达", "昌都邦达机场", ["昌都", "邦达"]),
    "ZUDR": ("西藏", "定日", "日喀则定日机场", ["日喀则", "定日"]),
    "ZULS": ("西藏", "贡嘎", "拉萨贡嘎机场", ["拉萨", "贡嘎"]),
    "ZUNZ": ("西藏", "米林", "林芝米林机场", ["林芝", "米林"]),
    "ZUPL": ("西藏", "普兰", "阿里普兰机场", ["阿里", "普兰"]),
    "ZURK": ("西藏", "和平", "日喀则和平机场", ["日喀则", "和平"]),
    "ZUSH": ("西藏", "隆子", "山南隆子机场", ["山南", "隆子"]),
    # ---- 贵州 ----
    "ZUAS": ("贵州", "黄果树", "安顺黄果树机场", ["安顺", "黄果树"]),
    "ZUBJ": ("贵州", "飞雄", "毕节飞雄机场", ["毕节", "飞雄"]),
    "ZUDJ": ("贵州", "德江", "铜仁德江机场", ["铜仁", "德江"]),
    "ZUGY": ("贵州", "龙洞堡", "贵阳龙洞堡机场", ["贵阳", "龙洞堡"]),
    "ZUKJ": ("贵州", "黄平", "凯里黄平机场", ["凯里", "黔东南", "黄平"]),
    "ZULB": ("贵州", "荔波", "黔南荔波机场", ["黔南", "荔波"]),
    "ZUMT": ("贵州", "茅台", "遵义茅台机场", ["遵义", "茅台"]),
    "ZUNP": ("贵州", "黎平", "黔东南黎平机场", ["黔东南", "黎平"]),
    "ZUPS": ("贵州", "月照", "六盘水月照机场", ["六盘水", "月照"]),
    "ZUTR": ("贵州", "凤凰", "铜仁凤凰机场", ["铜仁", "凤凰"]),
    "ZUYI": ("贵州", "万峰林", "兴义万峰林机场", ["兴义", "万峰林"]),
    "ZUZY": ("贵州", "新舟", "遵义新舟机场", ["遵义", "新舟"]),
    # ---- 辽宁 ----
    "ZYAS": ("辽宁", "腾鳌", "鞍山腾鳌机场", ["鞍山", "腾鳌"]),
    "ZYCH": ("辽宁", "", "长海机场", ["长海", "大连"]),
    "ZYCY": ("辽宁", "", "朝阳机场", ["朝阳"]),
    "ZYDD": ("辽宁", "浪头", "丹东浪头机场", ["丹东", "浪头"]),
    "ZYJZ": ("辽宁", "锦州湾", "锦州锦州湾机场", ["锦州", "锦州湾"]),
    "ZYTL": ("辽宁", "周水子", "大连周水子机场", ["大连", "周水子"]),
    "ZYTX": ("辽宁", "桃仙", "沈阳桃仙机场", ["沈阳", "桃仙"]),
    "ZYYK": ("辽宁", "兰旗", "营口兰旗机场", ["营口", "兰旗"]),
    # ---- 重庆 ----
    "ZUCK": ("重庆", "江北", "重庆江北机场", ["江北区", "江北", "重庆"]),
    "ZUQJ": ("重庆", "武陵山", "黔江武陵山机场", ["黔江区", "武陵山"]),
    "ZUWL": ("重庆", "仙女山", "重庆仙女山机场", ["武隆区", "仙女山"]),
    "ZUWS": ("重庆", "巫山", "重庆巫山机场", ["巫山县", "巫山"]),
    "ZUWX": ("重庆", "五桥", "万州五桥机场", ["万州区", "五桥"]),
    # ---- 陕西 ----
    "ZLAK": ("陕西", "富强", "安康富强机场", ["安康", "富强"]),
    "ZLHZ": ("陕西", "城固", "汉中城固机场", ["汉中", "城固"]),
    "ZLXY": ("陕西", "咸阳", "西安咸阳机场", ["西安", "咸阳"]),
    "ZLYA": ("陕西", "南泥湾", "延安南泥湾机场", ["延安", "南泥湾"]),
    "ZLYL": ("陕西", "榆阳", "榆林榆阳机场", ["榆林", "榆阳"]),
    # ---- 青海 ----
    "ZLDL": ("青海", "德令哈", "海西德令哈机场", ["海西", "德令哈"]),
    "ZLGL": ("青海", "玛沁", "果洛玛沁机场", ["果洛", "玛沁"]),
    "ZLGM": ("青海", "", "格尔木机场", ["格尔木", "海西"]),
    "ZLHB": ("青海", "祁连", "海北祁连机场", ["海北", "祁连"]),
    "ZLHX": ("青海", "茫崖", "海西茫崖机场", ["海西", "茫崖"]),
    "ZLXN": ("青海", "曹家堡", "西宁曹家堡机场", ["西宁", "曹家堡"]),
    "ZLYS": ("青海", "巴塘", "玉树巴塘机场", ["玉树", "巴塘"]),
    # ---- 黑龙江 ----
    "ZYDQ": ("黑龙江", "萨尔图", "大庆萨尔图机场", ["大庆", "萨尔图"]),
    "ZYDU": ("黑龙江", "德都", "五大连池德都机场", ["五大连池", "黑河", "德都"]),
    "ZYFY": ("黑龙江", "东极", "抚远东极机场", ["抚远", "佳木斯", "东极"]),
    "ZYHB": ("黑龙江", "太平", "哈尔滨太平机场", ["哈尔滨", "太平"]),
    "ZYHE": ("黑龙江", "瑷珲", "黑河瑷珲机场", ["黑河", "瑷珲"]),
    "ZYJD": ("黑龙江", "鄂伦春", "大兴安岭鄂伦春机场", ["大兴安岭", "鄂伦春"]),
    "ZYJM": ("黑龙江", "松江", "佳木斯松江机场", ["佳木斯", "松江"]),
    "ZYJS": ("黑龙江", "湿地", "建三江湿地机场", ["建三江", "佳木斯", "湿地"]),
    "ZYJX": ("黑龙江", "兴凯湖", "鸡西兴凯湖机场", ["鸡西", "兴凯湖"]),
    "ZYLD": ("黑龙江", "林都", "伊春林都机场", ["伊春", "林都"]),
    "ZYMD": ("黑龙江", "海浪", "牡丹江海浪机场", ["牡丹江", "海浪"]),
    "ZYMH": ("黑龙江", "古莲", "漠河古莲机场", ["漠河", "大兴安岭", "古莲"]),
    "ZYQQ": ("黑龙江", "三家子", "齐齐哈尔三家子机场", ["齐齐哈尔", "三家子"]),
    "ZYSD": ("黑龙江", "东宁", "绥芬河东宁机场", ["绥芬河", "牡丹江", "东宁"]),
}

MAINLAND_EXCLUDE_PREFIX = ("ZK", "ZM")   # 朝鲜、蒙古


def is_mainland_icao(icao) -> bool:
    s = str(icao or "").strip().upper()
    return len(s) == 4 and s[0] == "Z" and s[:2] not in MAINLAND_EXCLUDE_PREFIX


def resolve_segment(icao, city_text=""):
    """
    返回 dict：
      境内已知  {zone:'境内', province:'云南', candidates:[...], airport:'芒市', airport_name:'德宏芒市机场', error:None}
      境内未知  {zone:'境内', province:None, ..., error:'ICAO 不在维表'}
      境外      {zone:'境外', country:'日本', ...}
    只要 province 为空或 error 非空，调用方必须跳过该计划
    """
    city = str(city_text or "").strip()
    s = str(icao or "").strip().upper()

    if is_mainland_icao(s):
        t = AIRPORT_TABLE.get(s)
        if t:
            return {"zone": "境内", "icao": s, "province": t[0], "airport": t[1],
                    "airport_name": t[2], "candidates": list(t[3]), "error": None}
        return {"zone": "境内", "icao": s, "province": None, "airport": "",
                "airport_name": "", "candidates": [city] if city else [],
                "error": f"ICAO {s} 不在维表，请补 AIRPORT_TABLE"}

    country = extract_country(city) if city else ""
    return {"zone": "境外", "icao": s, "province": None, "airport": "",
            "airport_name": (city + "机场") if city else "",
            "candidates": [country] if country else [],
            "error": None if country else "境外段城市无法识别国家"}


def airport_preview(rows):
    """rows: [(icao, city), ...] → (已识别列表, 异常列表)"""
    ok, bad = [], []
    seen = set()
    for icao, city in rows:
        s = str(icao or "").strip().upper()
        if s in seen:
            continue
        seen.add(s)
        info = resolve_segment(icao, city)
        if info["zone"] == "境内" and info["province"]:
            ok.append({"ICAO": info["icao"], "原城市串": city, "省份": info["province"],
                       "第三级候选": " / ".join(info["candidates"]), "机场全名": info["airport_name"]})
        else:
            bad.append({"ICAO": s or "(空)", "原城市串": city,
                        "问题": info["error"] or "境外段（走国家选择）"})
    return ok, bad


def build_foreign_map(df):
    """只给境外城市建 城市→国家 的映射；境内全部走 AIRPORT_TABLE"""
    foreign = {}
    for _, row in df.iterrows():
        for icao_col, city_col in (("出发地", "出发城市"), ("到达地", "到达城市")):
            icao = str(row.get(icao_col) or "").strip().upper()
            city = str(row.get(city_col) or "").strip()
            if city and not is_mainland_icao(icao):
                foreign[city] = extract_country(city)
    return foreign


def parse_flight_time(time_str):
    try:
        parts = time_str.split(':')
        hours = int(parts[0])
        minutes = int(parts[1])
        return hours, minutes
    except:
        return 0, 0


def extract_country(city_name):
    if city_name.startswith("印尼"):
        return "印度尼西亚"
    parts = re.split(r'[\s\-]', city_name)
    if parts:
        first_part = parts[0]
        if first_part in ABBR_TO_COUNTRY:
            return ABBR_TO_COUNTRY[first_part]
    for tw_city in TAIWAN_CITIES:
        if tw_city in city_name:
            return "台湾"
    for country in COUNTRIES:
        if country in city_name:
            return country
    if parts:
        return parts[0]
    return city_name


# ============================================================
# 基础脚本（JS）
# ============================================================
_JS_HEAD = r"""
// ================= 公共辅助函数（基础脚本） =================

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function getMainDoc() {
    let iframe = document.querySelector('#main');
    if (iframe) {
        let doc = iframe.contentDocument;
        if (doc && doc.querySelector('body')) {
            return doc;
        }
    }
    console.log('未找到 iframe #main，使用顶层文档');
    return document;
}

async function waitForElement(selector, timeout = 15000, isXPath = true) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
        const doc = await getMainDoc();
        if (!doc) return null;
        let el;
        if (isXPath) {
            el = doc.evaluate(selector, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        } else {
            el = doc.querySelector(selector);
        }
        if (el) return el;
        await sleep(300);
    }
    console.warn(`⚠️ 等待元素超时: ${selector}`);
    return null;
}

// 下拉选项常常比页面晚一到几秒才出来，等它加载后再判断，别急着判空
async function ensureOptions(selectEl, getEl, timeout = 12000, minCount = 2) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
        if (selectEl && selectEl.options && selectEl.options.length >= minCount) return selectEl;
        await sleep(400);
        if (getEl) {
            try {
                const el2 = await getEl();
                if (el2) selectEl = el2;
            } catch (e) {}
        }
    }
    return (selectEl && selectEl.options && selectEl.options.length >= minCount) ? selectEl : null;
}

// 慢的时候给一次重试机会（8 秒后重试），别一失败就整批停
async function withRetry(fn, attempts = 2, gap = 8000, label = "") {
    for (let i = 0; i < attempts; i++) {
        let ok = false;
        try { ok = await fn(); } catch (e) { console.error(`${label} 抛异常`, e && e.message); ok = false; }
        if (ok) return true;
        if (i + 1 < attempts) {
            console.warn(`⚠️ ${label} 第 ${i + 1} 次未成功，${gap / 1000} 秒后重试一次（多为网页反应慢）`);
            await sleep(gap);
        }
    }
    return false;
}

function setDateInput(inputEl, dateStr) {
    if (!inputEl) return false;
    inputEl.value = dateStr;
    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    inputEl.blur();
    return true;
}

function setNumberInput(inputEl, value) {
    if (!inputEl) return false;
    inputEl.value = value;
    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    inputEl.blur();
    return true;
}

function formatRegNumber(reg) {
    if (reg.includes('-')) return reg;
    const match = reg.match(/^([A-Z]+)(\d+[A-Z]*)$/);
    if (match) return `${match[1]}-${match[2]}`;
    return reg;
}
"""

# ---- 机场映射与下拉联动（ICAO 优先，匹配不到即阻断）----
_JS_MAPPING = r"""
const MAINLAND_EXCLUDE_PREFIX = ["ZK", "ZM"];   // 朝鲜、蒙古

function normIcao(v) {
    return String(v == null ? "" : v).trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function isMainlandIcao(icao) {
    const s = normIcao(icao);
    return s.length === 4 && s.charAt(0) === "Z" && MAINLAND_EXCLUDE_PREFIX.indexOf(s.slice(0, 2)) === -1;
}

// 返回统一结构；province 为空 = 不可填，调用方必须跳过
function getLocationInfo(icao, cityRaw) {
    const city = String(cityRaw == null ? "" : cityRaw).trim();
    const code = normIcao(icao);

    if (isMainlandIcao(code)) {
        const e = AIRPORT_MAP[code];
        if (e) {
            return {
                zone: "境内", icao: code, province: e.p, candidates: (e.c || []).slice(),
                airport: e.a || "", airportName: e.n || "", raw: city, error: null
            };
        }
        return {
            zone: "境内", icao: code, province: null, candidates: city ? [city] : [],
            airport: "", airportName: "", raw: city, error: "ICAO " + code + " 不在维表"
        };
    }

    // 非 Z 段：境外，沿用原国家识别（CITY_MAP 只存境外城市→国家）
    let country = CITY_MAP[city];
    if (!country && city) {
        const parts = city.split(/[\s\-]/);
        country = parts[0] === "印尼" ? "印度尼西亚" : parts[0];
    }
    return {
        zone: "境外", icao: code, province: null, country: country || null,
        candidates: country ? [country] : [], airport: "", airportName: "",
        raw: city, error: country ? null : "境外段城市无法识别国家"
    };
}

// 去掉“国际/国内/机场”和空格后再比，机场全名写法差异（广州白云机场 vs 广州白云国际机场）也能对上
function normName(v) {
    return String(v == null ? "" : v).replace(/\s/g, "").replace(/(国际|国内|机场)/g, "");
}

// 下拉框匹配：未命中就是 0 分，绝不选第一个充数
async function setSelectValue(selectEl, valueText) {
    if (!selectEl) return false;
    const target = String(valueText == null ? "" : valueText).trim();
    if (!target) return false;

    // 慢的时候下拉选项比页面晚出来，先等选项加载（最多 12 秒）
    if (!selectEl.options || selectEl.options.length <= 1) {
        selectEl = await ensureOptions(selectEl, null, 12000, 2);
        if (!selectEl || !selectEl.options || selectEl.options.length <= 1) {
            console.warn(`  ⚠️ 下拉框选项 12 秒内没加载出来：目标 "${target}"`);
            return false;
        }
    }

    let bestIndex = -1, bestScore = 0, bestText = "";

    for (let i = 0; i < selectEl.options.length; i++) {
        const optText = (selectEl.options[i].text || "").trim();
        if (!optText || optText === "请选择") continue;
        if (optText === "其它" && target !== "其它") continue;   // 其它是占位项，但目标本身就是“其它”时必须可选

        let score = 0;
        if (optText === target) {
            score = 1000;
        } else if (optText === target + "机场") {
            score = 900;
        } else if (optText.indexOf(target) !== -1) {
            score = Math.max(1, 100 - (optText.length - target.length));
            if (optText.indexOf("直升机场") !== -1 || optText.indexOf("通用机场") !== -1) score -= 200;
            if (optText.indexOf("机场") !== -1) score += 50;
        } else if (optText.length >= 3 && target.indexOf(optText) !== -1) {
            score = Math.max(1, 40 - (target.length - optText.length));
            if (optText.indexOf("直升") !== -1 || optText.indexOf("通用") !== -1) score -= 200;
        } else {
            const nt = normName(target), no = normName(optText);
            if (nt && no === nt) {
                score = 800;
            } else if (nt && no.indexOf(nt) !== -1) {
                score = 400;
            } else if (nt && no.length >= 2 && nt.indexOf(no) !== -1) {
                score = 200;
            }
        }

        if (score > bestScore) { bestScore = score; bestIndex = i; bestText = optText; }
    }

    if (bestIndex !== -1 && bestScore > 0) {
        selectEl.selectedIndex = bestIndex;
        selectEl.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(300);
        console.log(`  ✅ 选择 "${bestText}"（目标 ${target}）`);
        return true;
    }

    console.warn(`  ⚠️ 下拉框无匹配项：目标 "${target}"`);
    return false;
}

// 航段（备案表单）：境内/境外 → 省 → 市/区
async function fillSegmentSelects(container, info) {
    if (!container) { console.error("未找到航段容器"); return false; }
    const sel = (i) => container.querySelectorAll("select")[i];
    if (!sel(0) || !sel(1)) { console.error("航段容器内 select 数量不足 2"); return false; }

    if (info.zone === "境外") {
        const ok1 = await setSelectValue(sel(0), "境外");
        await sleep(600);
        const ok2 = ok1 ? await setSelectValue(sel(1), info.candidates[0]) : false;
        if (!ok1 || !ok2) {
            console.error(`❌ ${info.icao} 境外国家匹配失败：${info.candidates[0] || "(空)"}`);
            return false;
        }
        return true;
    }

    if (!info.province) {
        console.error(`❌ ${info.icao} 省份未识别：${info.error || ""}`);
        return false;
    }

    const ok1 = await setSelectValue(sel(0), "境内");
    if (!ok1) {
        console.error(`❌ ${info.icao} 无法选择“境内”`);
        return false;
    }
    await sleep(600);
    const ok2 = await setSelectValue(sel(1), info.province);
    if (!ok2) {
        console.error(`❌ ${info.icao} 省级匹配失败：${info.province}`);
        return false;
    }

    // 第三级：省选完它才重建，耐心等 15 秒
    let s3 = null;
    const start = Date.now();
    while (Date.now() - start < 15000) {
        const cur = sel(2);
        if (!cur) {
            if (Date.now() - start > 6000) break;   // 确实没有第三级，止于省
        } else if (cur.options && cur.options.length > 1) {
            s3 = cur;
            break;
        }
        await sleep(400);
    }

    if (!sel(2)) {
        console.log(`  ℹ️ ${info.icao} 无第三级下拉，止于 ${info.province}`);
        return true;
    }
    if (!s3) {
        console.error(`❌ ${info.icao} 第三级下拉 15 秒内一直是空的（省=${info.province}），可能省份选错`);
        return false;
    }
    if (!s3.options || s3.options.length <= 1) {
        console.error(`❌ ${info.icao} 第三级下拉为空（省=${info.province}）`);
        return false;
    }

    for (const cand of info.candidates) {
        if (await setSelectValue(s3, cand)) {
            console.log(`  ✅ ${info.icao} 第三级 = ${cand}`);
            return true;
        }
        s3 = sel(2) || s3;
    }
    console.error(`❌ ${info.icao} 第三级无法匹配，候选=${info.candidates.join(" / ")}，请补 AIRPORT_TABLE 候选`);
    return false;
}

async function fillFirstSegmentSelects(icao, city) {
    const container = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[1]/div[1]/div/div', 15000, true);
    if (!container) { console.error("未找到第一个航段容器"); return false; }
    return fillSegmentSelects(container, getLocationInfo(icao, city));
}

async function fillSecondSegmentSelects(icao, city) {
    const container = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[2]/div[1]/div/div', 15000, true);
    if (!container) { console.error("未找到第二个航段容器"); return false; }
    return fillSegmentSelects(container, getLocationInfo(icao, city));
}

// 作业区（执行表单）：起飞机场 / 降落机场
async function fillAirportBlock(blockIndex, info, label) {
    let firstSelectXPath, secondSelectXPath;
    if (blockIndex === 1) {
        firstSelectXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[1]/div[2]/div/div[1]/div/select[1]';
        secondSelectXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[1]/div[2]/div/div[1]/div/select[2]';
    } else {
        firstSelectXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[2]/div[2]/div/div[1]/div/select[1]';
        secondSelectXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[2]/div[2]/div/div[1]/div/select[2]';
    }

    console.log(`⏳ 处理 ${label} 区块（${info.icao}）...`);
    const firstSelect = await waitForElement(firstSelectXPath, 15000);
    if (!firstSelect) { console.error(`❌ 未找到 ${label} 第一个选择框`); return false; }
    const okFirst = await setSelectValue(firstSelect, blockIndex === 1 ? "起飞机场" : "降落机场");

    const secondSelect = await waitForElement(secondSelectXPath, 15000);
    if (!secondSelect) { console.error(`❌ 未找到 ${label} 第二个选择框`); return false; }

    if (info.zone === "境内") {
        if (!info.province) { console.error(`❌ ${info.icao} 省份未识别，无法匹配机场`); return false; }
        const cands = [info.airportName, info.airport, info.candidates[0]].filter(Boolean);
        for (const cand of cands) {
            if (await setSelectValue(secondSelect, cand)) {
                console.log(`  ✅ ${label} 机场 = ${cand}`);
                return okFirst;
            }
        }
        console.error(`❌ ${label} 未匹配机场，候选=${cands.join(" / ")}`);
        return false;
    }

    // 境外：先在机场下拉里选“其它”，再填 [境外][境外] + 机场名称
    if (!(await setSelectValue(secondSelect, "其它"))) {
        console.error(`❌ ${label} 机场下拉里没有“其它”选项（选项写法可能变了），该计划请手动处理`);
        return false;
    }
    await sleep(1200);

    // 选完“其它”后，同一区块会多出两个“境外”下拉和一个机场名称输入框
    let box = firstSelect;
    while (box && !(box.contains && box.contains(secondSelect))) box = box.parentElement;

    let z1 = null, z2 = null, nameInput = null;
    if (box) {
        const allSel = box.querySelectorAll('select');
        if (allSel.length >= 4) { z1 = allSel[2]; z2 = allSel[3]; }
        const inputs = Array.from(box.querySelectorAll('input'))
            .filter(i => i.type !== 'checkbox' && i.type !== 'radio' && !i.readOnly);
        if (inputs.length) nameInput = inputs[inputs.length - 1];
    }

    // 兜底：按原来的绝对路径再找一遍（两种页面层级写法都试）
    if (!z1 || !z2) {
        const seg = blockIndex === 2 ? 'div[11]/div[2]/div[2]/div' : 'div[11]/div[1]/div[2]/div';
        const bases = [`/html/body/div[1]/div/div[3]/div/div[2]/form/${seg}/div[2]/div`,
                       `/html/body/div/div[1]/div[3]/div/div[2]/form/${seg}/div[2]/div`];
        for (const b of bases) {
            if (!z1) z1 = await waitForElement(b + '/select[1]', 3000);
            if (!z2) z2 = await waitForElement(b + '/select[2]', 3000);
            if (!nameInput) nameInput = await waitForElement(b + '/input', 3000);
            if (z1 && z2) break;
        }
    }

    if (!z1 || !z2) {
        console.error(`❌ ${label} 选了“其它”但没找到境外的两级下拉，该计划请手动处理`);
        return false;
    }

    if (nameInput) {
        const cityStr = info.raw || "";
        const guess = cityStr.replace(/^[^\s\-]+[\s\-]+/, "").trim() || cityStr;
        const nameVal = info.airportName || guess;
        if (nameVal) {
            setNumberInput(nameInput, nameVal);
            console.log(`  ℹ️ ${label} 境外机场名称填：${nameVal}${info.airportName ? "" : "（按城市串推断，请核对）"}`);
        }
    }

    const okZ = (await setSelectValue(z1, "境外")) && (await setSelectValue(z2, "境外"));
    if (nameInput && info.airportName) {
        setNumberInput(nameInput, info.airportName);
    }
    return okFirst && okZ;
}

// 运行前先看清楚映射，别等填完才发现不对
function logAirportMapping(records) {
    const rows = [];
    (records || []).forEach(r => {
        const dep = getLocationInfo(r.dep_icao || r["出发地"], r.dep_city || r["出发城市"]);
        const arr = getLocationInfo(r.arr_icao || r["到达地"], r.arr_city || r["到达城市"]);
        rows.push({
            航段: (r.reg || r["飞机注册号"]) + " " + (r.dep_city || r["出发城市"]) + "→" + (r.arr_city || r["到达城市"]),
            起飞: dep.icao, 起飞映射: dep.zone + "/" + (dep.province || dep.country || ""),
            起飞候选: (dep.candidates || []).join(" / "), 起飞问题: dep.error || "",
            降落: arr.icao, 降落映射: arr.zone + "/" + (arr.province || arr.country || ""),
            降落候选: (arr.candidates || []).join(" / "), 降落问题: arr.error || ""
        });
    });
    console.table(rows);
    const bad = rows.filter(x => x.起飞问题 || x.降落问题);
    if (bad.length) console.error(`⚠️ 有 ${bad.length} 条航段映射不完整，这些计划会自动跳过`);
    return bad.length;
}
"""

# ---- 时间、飞机、日期、返回列表等（与原逻辑一致）----
_JS_TAIL = r"""
async function fillFirstSegmentTime(record) {
    const hourInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[1]/div[2]/div/input[1]', 15000, true);
    const minuteInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[1]/div[2]/div/input[2]', 15000, true);
    const flightsInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[1]/div[3]/div/input', 15000, true);
    if (hourInput && minuteInput && flightsInput) {
        hourInput.value = record.flight_hours.toString().padStart(2, '0');
        minuteInput.value = record.flight_minutes.toString().padStart(2, '0');
        flightsInput.value = 1;
        hourInput.dispatchEvent(new Event('input', { bubbles: true }));
        minuteInput.dispatchEvent(new Event('input', { bubbles: true }));
        flightsInput.dispatchEvent(new Event('input', { bubbles: true }));
        console.log(`已填入第一个航段时间: ${record.flight_hours}:${record.flight_minutes}, 架次: 1`);
        return true;
    }
    console.warn('无法填入第一个航段时间和架次');
    return false;
}

async function fillSecondSegmentTime() {
    const hourInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[2]/div[2]/div/input[1]', 15000, true);
    const minuteInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[2]/div[2]/div/input[2]', 15000, true);
    const flightsInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[2]/div[3]/div/input', 15000, true);
    if (hourInput && minuteInput && flightsInput) {
        hourInput.value = '00';
        minuteInput.value = '00';
        flightsInput.value = 0;
        hourInput.dispatchEvent(new Event('input', { bubbles: true }));
        minuteInput.dispatchEvent(new Event('input', { bubbles: true }));
        flightsInput.dispatchEvent(new Event('input', { bubbles: true }));
        console.log('已填入第二个航段时间: 00:00, 架次: 0');
        return true;
    }
    console.warn('无法填入第二个航段时间和架次');
    return false;
}

async function selectAircraft(reg) {
    const regForSelect = formatRegNumber(reg);
    console.log(`尝试选择飞机: ${regForSelect}`);

    const airbox = await waitForElement('//*[@id="airbox"]', 10000, true);
    if (!airbox) {
        console.warn('未找到飞机选择对话框');
        return false;
    }
    const doc = await getMainDoc();

    let span = doc.evaluate(`//span[text()='${regForSelect}']`, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (!span) {
        span = doc.evaluate(`//span[contains(text(), '${regForSelect}')]`, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    }
    if (!span) {
        console.warn(`未找到包含注册号 ${regForSelect} 的 span 元素`);
        return false;
    }
    const li = span.closest('li');
    if (!li) {
        console.warn(`未找到注册号 ${regForSelect} 对应的 li`);
        return false;
    }
    const checkbox = li.querySelector('input[type="checkbox"]');
    if (!checkbox) {
        console.warn(`未找到注册号 ${regForSelect} 的复选框`);
        return false;
    }
    checkbox.click();
    console.log('已勾选飞机复选框');
    await sleep(300);

    let closeBtn = doc.evaluate('//*[@id="close7"]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (!closeBtn) {
        closeBtn = doc.evaluate('//button[contains(text(), "关闭")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    }
    if (!closeBtn) {
        closeBtn = doc.evaluate('//button[contains(text(), "确定")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    }
    if (closeBtn) {
        closeBtn.click();
        console.log('已关闭选择对话框');
        await sleep(500);
        return true;
    } else {
        console.warn('未找到关闭按钮，尝试按 ESC 关闭');
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
        await sleep(500);
        return false;
    }
}

async function waitForDialogConfirmButton(timeout = 15000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
        let btn = document.evaluate('//a[contains(text(), "确定")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!btn) btn = document.evaluate('//button[contains(text(), "确定")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (btn) return btn;
        try {
            const doc = await getMainDoc();
            btn = doc.evaluate('//a[contains(text(), "确定")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (!btn) btn = doc.evaluate('//button[contains(text(), "确定")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (btn) return btn;
        } catch(e) {}
        await sleep(300);
    }
    return null;
}

async function handleSecondFlightTime() {
    const hourXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[2]/div[2]/div/div[5]/div/input[1]';
    const minuteXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[2]/div[2]/div/div[5]/div/input[2]';
    const countXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[2]/div[2]/div/div[6]/div/input';
    const hourInput = await waitForElement(hourXPath, 10000);
    const minuteInput = await waitForElement(minuteXPath, 10000);
    const countInput = await waitForElement(countXPath, 10000);
    if (hourInput && minuteInput && countInput) {
        console.log('⏱️ 填入第二个飞行时间: 00:00，飞行架次数: 0');
        setNumberInput(hourInput, '0');
        setNumberInput(minuteInput, '0');
        setNumberInput(countInput, '0');
        return true;
    } else {
        console.warn('⚠️ 未找到第二个飞行时间输入框');
        return false;
    }
}

async function handleDetailArea(depInfo, arrInfo) {
    const detailXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[16]/div/input';
    const detailInput = await waitForElement(detailXPath, 10000);
    if (!detailInput) { console.error('❌ 未找到详细作业区输入框'); return false; }
    const depName = depInfo.airportName || (depInfo.candidates[0] || "") + "机场";
    const arrName = arrInfo.airportName || (arrInfo.candidates[0] || "") + "机场";
    const detailText = `${depName}-${arrName}`;
    console.log(`📝 填入详细作业区: ${detailText}`);
    setNumberInput(detailInput, detailText);
    return true;
}

async function waitForReturnToList(timeout = 300000) {
    const start = Date.now();
    console.log('⏳ 等待您手动点击“提交”后返回列表页...');
    while (Date.now() - start < timeout) {
        const doc = await getMainDoc();
        if (!doc) return false;
        const rows = doc.querySelectorAll(ROW_SELECTOR);
        if (rows.length > 0) {
            console.log('✅ 已返回列表页');
            return true;
        }
        await sleep(2000);
    }
    console.error('❌ 等待超时，未返回列表页');
    return false;
}

async function ensureListPage(passiveTimeout = 90000) {
    let btn = await waitForElement('input.query.yuanjiao', 10000, false);
    if (btn) return true;

    console.log('当前不在列表页，先尝试点掉遗留的按钮...');
    const doc = await getMainDoc();
    if (doc) {
        let closeBtn = doc.evaluate('//button[contains(text(), "取消")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!closeBtn) closeBtn = doc.evaluate('//button[contains(text(), "关闭")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!closeBtn) closeBtn = doc.evaluate('//input[@type="button" and contains(@value, "返回")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!closeBtn) closeBtn = doc.evaluate('//a[contains(text(), "返回")]', doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (closeBtn) {
            closeBtn.click();
            console.log('已点击返回/关闭按钮，等待列表页...');
            await sleep(1200);
        }
    }

    // 表单页偶尔没有可点的返回按钮，这里不硬失败，给时间手动收尾
    const start = Date.now();
    console.log(`⚠️ 当前还停在表单页，请手动点“返回/取消”回到列表页（最多等 ${Math.round(passiveTimeout / 1000)} 秒）...`);
    while (Date.now() - start < passiveTimeout) {
        await sleep(2000);
        const backBtn = await waitForElement('input.query.yuanjiao', 2000, false);
        if (backBtn) {
            console.log('✅ 已回到列表页，继续执行');
            return true;
        }
    }
    console.error('❌ 等待超时仍未回到列表页');
    return false;
}
"""


_BJ_TZ = timezone(timedelta(hours=8))   # Streamlit Cloud 服务器多为 UTC，这里固定按北京时间显示


def gen_time_text():
    """脚本生成时间（北京时间，精确到秒）"""
    return datetime.now(_BJ_TZ).strftime("%Y-%m-%d %H:%M:%S")


def generate_base_script(city_map_json):
    airport_json = json.dumps(
        {icao: {"p": t[0], "a": t[1], "n": t[2], "c": t[3]} for icao, t in AIRPORT_TABLE.items()},
        ensure_ascii=False, separators=(",", ":")
    )
    header = (f"// ============================================================\n"
              f"// 生成时间：{gen_time_text()}（北京时间 UTC+8）\n"
              f"// 机场维表 {len(AIRPORT_TABLE)} 条；境内按 ICAO 判定，匹配不到即跳过该计划\n"
              f"// ============================================================\n")
    return (header
            + _JS_HEAD
            + "\nconst CITY_MAP = " + city_map_json + ";\n"
            + "\nconst AIRPORT_MAP = " + airport_json + ";\n"
            + _JS_MAPPING
            + _JS_TAIL)


def generate_daily_script(records):
    js_data = json.dumps(records, ensure_ascii=False, indent=4)
    script = f"""
// ================= 自动生成的飞行计划脚本（当日计划专用） =================
// 当日待处理计划数: {len(records)}
// =========================================================

const IFRAME_ID = 'main';
const ROW_SELECTOR = 'table tbody:nth-of-type(2) tr';
const REG_SELECTOR = 'td:nth-child(6) div';
const SEGMENT_SELECTOR = 'td:nth-child(7) div';
const DATE_SELECTOR = 'td:nth-child(9)';

const excelData = {js_data};

function normalizeReg(reg) {{
    return reg.replace(/[-\\s]/g, '').trim();
}}

function getPlanKey(plan) {{
    return `${{plan["飞机注册号"]}}_${{plan["出发日期"]}}_${{plan["出发城市"]}}_${{plan["到达城市"]}}`;
}}

async function waitForTable() {{
    const start = Date.now();
    while (Date.now() - start < 10000) {{
        const doc = await getMainDoc();
        if (!doc) return null;
        const rows = doc.querySelectorAll(ROW_SELECTOR);
        if (rows.length > 0) return rows;
        await sleep(300);
    }}
    return null;
}}

function extractCityKeywords(segText) {{
    let parts = segText.split(',');
    if (parts.length < 2) return null;
    let depPart = parts[0].trim();
    let arrPart = parts[1].trim();
    depPart = depPart.replace(/^(境内|境外)-/, '');
    arrPart = arrPart.replace(/^(境内|境外)-/, '');
    let extract = (part) => {{
        let segments = part.split('-');
        let keywords = [];
        for (let seg of segments) {{
            let words = seg.split(/\\s+/);
            for (let w of words) if (w) keywords.push(w);
        }}
        return keywords;
    }};
    return {{ depKeywords: extract(depPart), arrKeywords: extract(arrPart) }};
}}

async function tryMatchPlan(plan) {{
    const rows = await waitForTable();
    if (!rows) return null;
    const regTarget = normalizeReg(plan["飞机注册号"]);
    const dateTarget = plan["出发日期"];
    const depCity = plan["出发城市"];
    const arrCity = plan["到达城市"];
    const depInfo = getLocationInfo(plan["出发地"], depCity);
    const arrInfo = getLocationInfo(plan["到达地"], arrCity);
    const depKey = depInfo.zone === "境外" ? (depInfo.candidates[0] || depCity) : (depInfo.airportName || depCity);
    const arrKey = arrInfo.zone === "境外" ? (arrInfo.candidates[0] || arrCity) : (arrInfo.airportName || arrCity);

    for (let i = 0; i < rows.length; i++) {{
        const row = rows[i];
        const regEl = row.querySelector(REG_SELECTOR);
        const regNo = regEl ? normalizeReg(regEl.innerText.trim()) : null;
        if (regNo !== regTarget) continue;
        const dateCell = row.querySelector(DATE_SELECTOR);
        const webDate = dateCell ? dateCell.innerText.trim() : '';
        if (webDate !== dateTarget) continue;
        const segEl = row.querySelector(SEGMENT_SELECTOR);
        if (!segEl) continue;
        const segText = segEl.innerText.trim();
        const kw = extractCityKeywords(segText);
        if (!kw) continue;
        const depMatch = kw.depKeywords.some(k => depKey.includes(k) || k.includes(depCity));
        const arrMatch = kw.arrKeywords.some(k => arrKey.includes(k) || k.includes(arrCity));
        if (depMatch && arrMatch) {{
            return row;
        }}
    }}
    return null;
}}

async function waitForManualExecute(plan) {{
    console.log(`%c⏳ 等待手动操作：请点击计划【${{plan["飞机注册号"]}} - ${{plan["出发城市"]}} → ${{plan["到达城市"]}}】对应的“执行”按钮`, 'background: #ff9800; color: white; font-size: 14px; padding: 4px 8px; border-radius: 4px');
    console.log(`%c💡 按 空格键 或 回车键 可跳过此计划，继续处理下一个`, 'background: #2196F3; color: white; font-size: 12px; padding: 2px 6px; border-radius: 4px');
    console.log(`计划详情: 日期=${{plan["出发日期"]}}, 用途=${{plan["用途"]}}, 飞行时间=${{plan["实际飞行时间"]}}`);

    let skipPressed = false;
    const keyHandler = (e) => {{
        if (e.key === ' ' || e.key === 'Enter') {{
            e.preventDefault();
            skipPressed = true;
            document.removeEventListener('keydown', keyHandler);
            console.log(`⏭️ 检测到空格/回车，跳过计划 ${{plan["飞机注册号"]}}`);
        }}
    }};
    document.addEventListener('keydown', keyHandler);

    const startDateXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[9]/div/input';
    let startInput = null;
    let attempts = 0;
    while (attempts < 180 && !skipPressed) {{
        await sleep(500);
        if (skipPressed) break;
        const doc = await getMainDoc();
        if (!doc) continue;
        startInput = doc.evaluate(startDateXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (startInput) break;
        attempts++;
    }}
    document.removeEventListener('keydown', keyHandler);

    if (skipPressed) {{
        console.log(`⏭️ 已跳过计划：${{plan["飞机注册号"]}}`);
        return null;
    }}
    if (!startInput) {{
        console.error('❌ 等待超时，未检测到表单加载，将跳过此计划');
        return null;
    }}
    console.log('✅ 检测到表单已加载，继续填写');
    return await getMainDoc();
}}

async function processExistingPlanWithManual(row, plan) {{
    console.log(`\\n🔧 开始处理计划（手动执行后自动填写）：机号 ${{plan["飞机注册号"]}}`);

    const doc = await waitForManualExecute(plan);
    if (!doc) {{
        console.log(`⏭️ 计划已被跳过或超时，继续下一个。`);
        return false;
    }}

    const startDateXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[9]/div/input';
    const endDateXPath   = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[10]/div/input';
    console.log('⏳ 等待日期输入框...');
    const startInput = await waitForElement(startDateXPath, 15000);
    const endInput   = await waitForElement(endDateXPath, 15000);
    if (!startInput || !endInput) {{
        console.error('❌ 未找到日期输入框，可能表单未正确加载');
        return false;
    }}
    const startDate = plan["出发日期"];
    const endDate   = plan["到达日期"];
    console.log(`📅 填入作业开始日期: ${{startDate}}`);
    console.log(`📅 填入作业结束日期: ${{endDate}}`);
    setDateInput(startInput, startDate);
    setDateInput(endInput, endDate);

    const depInfo = getLocationInfo(plan["出发地"], plan["出发城市"]);
    const arrInfo = getLocationInfo(plan["到达地"], plan["到达城市"]);
    console.log(`🗺️ 起飞映射: ${{depInfo.icao}} -> ${{depInfo.zone}}/${{depInfo.province || depInfo.country || ''}}`);
    console.log(`🗺️ 降落映射: ${{arrInfo.icao}} -> ${{arrInfo.zone}}/${{arrInfo.province || arrInfo.country || ''}}`);

    if (!(await withRetry(() => fillAirportBlock(1, depInfo, "起飞"), 2, 10000, "起飞机场"))) return false;
    if (!(await withRetry(() => fillAirportBlock(2, arrInfo, "降落"), 2, 10000, "降落机场"))) return false;

    const actualFlightTime = plan["实际飞行时间"];
    if (actualFlightTime && actualFlightTime.includes(':')) {{
        const flightHourXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[1]/div[2]/div/div[5]/div/input[1]';
        const flightMinuteXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[1]/div[2]/div/div[5]/div/input[2]';
        const flightCountXPath = '/html/body/div[1]/div/div[3]/div/div[2]/form/div[11]/div[1]/div[2]/div/div[6]/div/input';
        const hourInput = await waitForElement(flightHourXPath, 10000);
        const minuteInput = await waitForElement(flightMinuteXPath, 10000);
        const countInput = await waitForElement(flightCountXPath, 10000);
        if (hourInput && minuteInput && countInput) {{
            const [hour, minute] = actualFlightTime.split(':');
            console.log(`⏱️ 填入第一个飞行时间: ${{hour}}:${{minute}}，飞行架次数: 1`);
            setNumberInput(hourInput, hour);
            setNumberInput(minuteInput, minute);
            setNumberInput(countInput, '1');
        }} else console.warn('⚠️ 未找到第一个飞行时间输入框');
    }} else console.warn(`⚠️ 实际飞行时间 "${{actualFlightTime}}" 格式不正确`);

    await handleSecondFlightTime();
    await handleDetailArea(depInfo, arrInfo);

    console.log('✅ 所有字段填写完成，请手动点击“提交”按钮。');
    const returned = await waitForReturnToList();
    if (!returned) {{
        console.error('❌ 未能检测到返回列表页，请手动返回后继续运行脚本。');
        return false;
    }}
    return true;
}}

async function runDailyPlans() {{
    console.log('🚀 开始执行当日计划自动化流程...');
    logAirportMapping(excelData);
    let processedCount = 0;
    const processedKeys = new Set();

    for (let plan of excelData) {{
        const key = getPlanKey(plan);
        if (processedKeys.has(key)) continue;
        processedKeys.add(key);
        processedCount++;
        console.log(`\\n========== 处理第 ${{processedCount}} 个当日计划 ==========`);
        console.log(`计划信息: 机号=${{plan["飞机注册号"]}}, 出发=${{plan["出发城市"]}}, 到达=${{plan["到达城市"]}}, 日期=${{plan["出发日期"]}}`);

        if (!plan["出发地"] || !plan["到达地"]) {{
            console.error(`⚠️ 计划缺少 ICAO（出发地/到达地），无法判定境内境外，跳过`);
            continue;
        }}

        let matchedRow = await tryMatchPlan(plan);
        if (matchedRow) {{
            console.log(`✅ 匹配到已有航段，将尝试自动点击执行并填写`);
        }} else {{
            console.log(`⚠️ 未匹配到已有航段，请手动点击该计划的“执行”按钮，脚本将自动填写表单`);
        }}
        const success = await processExistingPlanWithManual(null, plan);
        if (!success) {{
            console.error(`⚠️ 第 ${{processedCount}} 个计划未填完，先回到列表页再继续`);
            const back = await ensureListPage(90000);
            if (!back) {{
                console.error('❌ 回不到列表页，当日流程中止。请手动回到列表页，然后把脚本重新粘贴一次，从这条计划继续。');
                return;
            }}
            await sleep(1500);
        }}
        await sleep(1000);
    }}
    await ensureListPage();
    console.log('🎉 所有当日计划处理完毕！');
}}
"""
    return script


def generate_nextday_script(records, func_name="runNextDayPlans"):
    flight_records_json = json.dumps(records, ensure_ascii=False, indent=4)
    template = f"""
// ================= 自动生成的飞行计划脚本（次日/补备案专用） =================
// 待处理计划数: {len(records)}
// =========================================================

async function processNextDayRecord(record) {{
    console.log(`\\n开始处理计划（次日/补备案）：${{record.reg}} - ${{record.dep_city}} -> ${{record.arr_city}}`);
    console.log(`   ICAO: ${{record.dep_icao}} -> ${{record.arr_icao}}`);

    if (!(await ensureListPage())) {{
        console.error('无法返回列表页，终止流程');
        return false;
    }}

    const addBtn = await waitForElement('input.query.yuanjiao', 15000, false);
    if (!addBtn) {{
        console.error('未找到添加按钮，终止流程');
        return false;
    }}
    addBtn.click();
    console.log('已点击添加按钮，等待表单加载...');

    const aircraftSelectBtn = await waitForElement('//*[@id="ele7"]', 15000, true);
    if (!aircraftSelectBtn) {{
        console.error('未找到飞机机号选择按钮，终止流程');
        return false;
    }}
    console.log('表单加载完成，找到飞机机号选择按钮');

    aircraftSelectBtn.click();
    if (!(await selectAircraft(record.reg))) {{
        console.error('选择飞机失败，终止流程');
        return false;
    }}

    const doc = await getMainDoc();
    const specialSelect = doc.querySelector('#specialf');
    if (specialSelect) await setSelectValue(specialSelect, "否");
    else console.warn('未找到是否特殊任务飞行 select');

    const certSelect = doc.querySelector('#operationCertificate');
    if (certSelect) await setSelectValue(certSelect, "是");
    else console.warn('未找到是否有运行合格证 select');

    const operateSelect = doc.querySelector('#businessOperation');
    if (operateSelect) await setSelectValue(operateSelect, "否");
    else console.warn('未找到是否经营性作业 select');

    let purpose = "自用飞行";
    const purposeRaw = record.purpose || "";
    if (purposeRaw.includes("维修") || purposeRaw.includes("调机")) {{
        purpose = "调机";
    }}
    const purposeSelect = await waitForElement('//*[contains(text(), "非经营活动")]/following-sibling::*//select', 10000, true);
    if (purposeSelect) await setSelectValue(purposeSelect, purpose);
    else console.warn('未找到用途下拉框');

    const startDateInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[9]/div/input', 15000, true);
    const endDateInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[10]/div/input', 15000, true);
    if (startDateInput) setDateInput(startDateInput, record.start_date);
    else console.warn('未找到服务开始日期输入框');
    if (endDateInput) setDateInput(endDateInput, record.end_date);
    else console.warn('未找到服务结束日期输入框');

    const depInfo = getLocationInfo(record.dep_icao, record.dep_city);
    const arrInfo = getLocationInfo(record.arr_icao, record.arr_city);

    if (!(await withRetry(() => fillFirstSegmentSelects(record.dep_icao, record.dep_city), 2, 10000, "起飞机段"))) {{
        console.error(`❌ 起飞机段映射失败（${{record.dep_icao}} ${{record.dep_city}}），跳过该计划，不提交`);
        return false;
    }}
    if (!(await fillFirstSegmentTime(record))) {{
        console.error('❌ 第一个航段时间填写失败，终止该计划');
        return false;
    }}

    const addSegmentBtn = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[23]/div[1]/div[1]/div/div/button', 15000, true);
    if (addSegmentBtn) {{
        addSegmentBtn.click();
        console.log('已点击添加航段按钮，等待新航段加载...');
        await sleep(1000);
        if (!(await withRetry(() => fillSecondSegmentSelects(record.arr_icao, record.arr_city), 3, 8000, "降落航段"))) {{
            console.error(`❌ 降落航段映射失败（${{record.arr_icao}} ${{record.arr_city}}），跳过该计划，不提交`);
            return false;
        }}
        await fillSecondSegmentTime();
    }} else {{
        console.warn('未找到添加航段按钮');
    }}

    const detailAreaInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[25]/div/input', 15000, true);
    if (detailAreaInput) {{
        const detailText = `${{depInfo.airportName || record.dep_city}}-${{arrInfo.airportName || record.arr_city}}`;
        detailAreaInput.value = detailText;
        detailAreaInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        console.log(`已填入详细作业地区: ${{detailText}}`);
    }} else console.warn('未找到详细作业地区输入框');

    const customerInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[27]/div/input', 5000, true);
    if (customerInput) {{
        customerInput.value = "天成商务航空有限公司";
        customerInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        console.log('已填入服务客户名称');
    }} else console.warn('未找到服务客户名称输入框');

    const baseInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[28]/div/input', 5000, true);
    if (baseInput) {{
        const baseText = `${{depInfo.airportName || record.dep_city + "机场"}}-${{arrInfo.airportName || record.arr_city + "机场"}}`;
        baseInput.value = baseText;
        baseInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        console.log(`已填入作业基地名称: ${{baseText}}`);
    }} else console.warn('未找到作业基地名称输入框');

    const operatorInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[29]/div/input', 5000, true);
    if (operatorInput) {{
        operatorInput.value = "张永一";
        operatorInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        console.log('已填入作业负责人姓名');
    }} else console.warn('未找到作业负责人姓名输入框');

    const phoneInput = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[30]/div/input', 5000, true);
    if (phoneInput) {{
        phoneInput.value = "18566725728";
        phoneInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        console.log('已填入负责人联系电话');
    }} else console.warn('未找到负责人联系电话输入框');

    const contractSelect = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[36]/div/select', 5000, true);
    if (contractSelect) await setSelectValue(contractSelect, "已签订");
    else console.warn('未找到合同订立情况下拉框');

    const insuranceSelect = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[37]/div/select', 5000, true);
    if (insuranceSelect) await setSelectValue(insuranceSelect, "已参保");
    else console.warn('未找到保险情况下拉框');

    const submitBtn = await waitForElement('/html/body/div[1]/div/div[3]/div/div[2]/form/div[40]/ul/li[2]/input', 5000, true);
    if (submitBtn) {{
        submitBtn.click();
        console.log('✅ 已点击“提交”按钮，等待弹窗...');
        let confirmBtn = null;
        for (let attempt = 0; attempt < 8; attempt++) {{
            confirmBtn = await waitForDialogConfirmButton(2000);
            if (confirmBtn) break;
            console.log(`等待确定按钮... 第${{attempt+1}}次尝试`);
        }}
        if (confirmBtn) {{
            confirmBtn.click();
            console.log('✅ 已点击确定按钮，提交完成。');
            const backBtn = await waitForElement('input.query.yuanjiao', 5000, false);
            if (backBtn) {{
                console.log('✅ 已返回列表页');
            }} else {{
                console.warn('⚠️ 未检测到返回列表页，但继续执行下一个计划。');
            }}
        }} else {{
            console.warn('⚠️ 未找到确定按钮，可能出现了错误弹窗，等待30秒供人工处理...');
            await sleep(30000);
            console.log('⏳ 继续执行下一个计划（假设已人工处理完成）。');
        }}
        console.log(`处理完成：${{record.reg}}`);
    }} else {{
        console.warn('未找到提交按钮');
    }}
    await sleep(2000);
    return true;
}}

async function {func_name}() {{
    const flightRecords = {flight_records_json};
    console.log(`🚀 开始执行计划自动化流程（{func_name}），共 ${{flightRecords.length}} 条计划...`);
    logAirportMapping(flightRecords);
    const skipped = [];
    for (let i = 0; i < flightRecords.length; i++) {{
        let success = false;
        try {{
            success = await processNextDayRecord(flightRecords[i]);
        }} catch (e) {{
            console.error(`第 ${{i+1}} 条计划抛异常：${{e && e.message}}`);
            success = false;
        }}
        if (!success) {{
            skipped.push(`${{flightRecords[i].reg}} ${{flightRecords[i].dep_city}}→${{flightRecords[i].arr_city}}`);
            console.error(`⚠️ 第 ${{i+1}} 条计划未能填写，已跳过，继续下一条`);
            await sleep(3000);
        }}
    }}
    console.log(`✅ 完成 ${{flightRecords.length - skipped.length}} 条；⏭ 跳过 ${{skipped.length}} 条`, skipped);
}}
"""
    return template


def run_feature_d():
    # 精简上传文案
    uploaded_file = sess_file_uploader("上传：航段数据导出", type=["xlsx", "xls"], key="d_upload")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=1)
            df.columns = df.columns.str.strip()
            df = df.dropna(how='all')

            required_cols = ["飞机注册号", "出发日期", "到达日期", "用途", "出发城市", "到达城市",
                             "出发地", "到达地", "预计飞行时间", "实际到达"]
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                st.error(f"❌ 缺少必要列: {missing}")
                st.info(f"实际列名: {list(df.columns)}")
                return

            # 剔除不是 B 开头的外机计划
            df = df[df["飞机注册号"].astype(str).str.startswith("B", na=False)].copy()

            # 处理日期格式
            df['出发日期'] = pd.to_datetime(df['出发日期']).dt.date
            today = date.today()
            tomorrow = today + timedelta(days=1)

            # 数据拆分：今日、明日、未来
            df_today = df[df['出发日期'] == today].copy()
            df_tomorrow = df[df['出发日期'] == tomorrow].copy()
            df_future = df[df['出发日期'] > tomorrow].copy()

            def get_daily_rec(row):
                rec_dict = row.to_dict()
                for k, v in rec_dict.items():
                    if pd.isna(v):
                        rec_dict[k] = ""
                    elif isinstance(v, (date, datetime)):
                        rec_dict[k] = str(v)
                rec_dict["出发地"] = str(rec_dict.get("出发地", "") or "").strip().upper()
                rec_dict["到达地"] = str(rec_dict.get("到达地", "") or "").strip().upper()
                return rec_dict

            def get_future_rec(row):
                purpose_raw = row.get("用途", "")
                if "维修" in purpose_raw or "调机" in purpose_raw:
                    purpose = "调机"
                else:
                    purpose = "自用飞行"
                start_date = str(row["出发日期"])
                end_date = str(row["到达日期"])

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
                    "dep_icao": str(row.get("出发地") or "").strip().upper(),
                    "arr_icao": str(row.get("到达地") or "").strip().upper(),
                    "start_date": start_date,
                    "end_date": end_date,
                    "purpose": purpose,
                    "dep_city": dep_city,
                    "arr_city": arr_city,
                    "flight_hours": hours,
                    "flight_minutes": minutes
                }

            col1, col2 = st.columns(2)

            list_backup = []     # 补备案队列
            list_execute = []    # 执行队列
            list_future = []     # 未来备案队列

            with col1:
                st.markdown("### 📋 执行的计划（今日）")
                if len(df_today) > 0:
                    for idx, row in df_today.iterrows():
                        label = f"{row['飞机注册号']} {row['出发日期']} {row['出发城市']}->{row['到达城市']}"
                        is_landed = pd.notna(row["实际到达"]) and str(row["实际到达"]).strip() != ""

                        c1, c2 = st.columns([4, 1])
                        with c1:
                            # 未飞完的计划，直接禁用执行框，防止误勾
                            do_exec = st.checkbox(label, value=is_landed, key=f"exec_{idx}", disabled=not is_landed)
                        with c2:
                            # 补备案默认不勾选，任何时候都可以勾选
                            do_backup = st.checkbox("补备案", value=False, key=f"backup_{idx}", help="如果昨天漏备案，请勾选此项（脚本将先补备案，再执行）")

                        if do_exec:
                            list_execute.append(get_daily_rec(row))
                        if do_backup:
                            list_backup.append(get_future_rec(row))
                else:
                    st.info("暂无今日计划")

            with col2:
                st.markdown("### 📝 备案的计划（明日）")
                if len(df_tomorrow) > 0:
                    select_all_file = st.checkbox("全选明日计划", value=True, key="select_all_tomorrow")
                    for idx, row in df_tomorrow.iterrows():
                        label = f"{row['飞机注册号']} {row['出发日期']} {row['出发城市']}->{row['到达城市']}"
                        do_file = st.checkbox(label, value=select_all_file, key=f"file_tomorrow_{idx}")
                        if do_file:
                            list_future.append(get_future_rec(row))
                else:
                    st.info("暂无明日计划")

            # 显示未来计划（折叠区域）
            st.markdown("---")
            show_future = st.checkbox("📅 显示后天及未来的计划", value=False, key="show_future_plans")

            if show_future:
                if len(df_future) > 0:
                    # 按日期分组
                    future_dates = sorted(df_future['出发日期'].unique())
                    for d in future_dates:
                        df_day = df_future[df_future['出发日期'] == d]
                        with st.expander(f"📅 {d} (共 {len(df_day)} 条计划)", expanded=False):
                            for idx, row in df_day.iterrows():
                                label = f"{row['飞机注册号']} {row['出发城市']}->{row['到达城市']}"
                                do_file = st.checkbox(label, value=False, key=f"file_future_{idx}")
                                if do_file:
                                    list_future.append(get_future_rec(row))
                else:
                    st.info("暂无未来计划")

            # ===== 机场映射预览（新增，境内一律按 ICAO 判定）=====
            preview_rows = [(r.get("出发地"), r.get("出发城市")) for _, r in df.iterrows()] + \
                           [(r.get("到达地"), r.get("到达城市")) for _, r in df.iterrows()]
            ok_rows, bad_rows = airport_preview(preview_rows)
            with st.expander(f"🏷️ 机场映射核对（{len(ok_rows)} 个境内机场）", expanded=bool(bad_rows)):
                if ok_rows:
                    st.dataframe(pd.DataFrame(ok_rows), use_container_width=True, height=240)
                if bad_rows:
                    st.error(f"以下 {len(bad_rows)} 个机场未识别，对应计划运行时会被自动跳过：")
                    st.dataframe(pd.DataFrame(bad_rows), use_container_width=True)

            # 汇总提示
            st.success(f"✅ 最终脚本将按顺序执行：{len(list_backup)} 条补备案 -> {len(list_execute)} 条执行 -> {len(list_future)} 条备案")

            city_map = build_foreign_map(df)
            city_map_json = json.dumps(city_map, ensure_ascii=False, indent=4)

            base_script = generate_base_script(city_map_json)
            backup_script = generate_nextday_script(list_backup, func_name="runBackupPlans") if len(list_backup) > 0 else ""
            daily_script = generate_daily_script(list_execute) if len(list_execute) > 0 else ""
            future_script = generate_nextday_script(list_future, func_name="runNextDayPlans") if len(list_future) > 0 else ""

            final_script = base_script + "\n\n" + backup_script + "\n\n" + daily_script + "\n\n" + future_script + """
(async () => {
    console.log("========== 开始执行综合流程 ==========");

    // 第一步：先处理昨日漏备案的今日计划（必须先建记录）
    if (typeof runBackupPlans === 'function') {
        console.log(">>> 步骤 1/3: 开始补备案...");
        await runBackupPlans();
    } else {
        console.log(">>> 步骤 1/3: 无需补备案。");
    }

    // 第二步：再执行今日已飞完的计划
    if (typeof runDailyPlans === 'function') {
        console.log(">>> 步骤 2/3: 开始执行当日已飞完计划...");
        await runDailyPlans();
    } else {
        console.log(">>> 步骤 2/3: 无需执行当日计划。");
    }

    // 第三步：最后备案未来计划
    if (typeof runNextDayPlans === 'function') {
        console.log(">>> 步骤 3/3: 开始备案未来计划...");
        await runNextDayPlans();
    } else {
        console.log(">>> 步骤 3/3: 无需备案未来计划。");
    }

    console.log("========== 综合流程全部完成 ==========");
})();
"""
            st.subheader("📋 复制以下代码到浏览器控制台（F12）运行")
            st.caption(f"⏱ 本脚本生成于 {gen_time_text()}（北京时间），换行/勾选后时间会刷新——粘贴前先确认是最新时间。")

            # 新增手动刷新按钮
            if st.button("🔄 强制刷新脚本", help="如果下方代码没有更新，请点击此按钮"):
                st.rerun()

            st.code(final_script, language="javascript")

            st.info("💡 重要提示：\n1. 每次勾选计划后，下方脚本会自动刷新。\n2. **在粘贴到浏览器控制台之前，请先输入 `clear()` 并回车，清空旧代码**，然后再粘贴最新代码运行，以免旧脚本残留导致执行混乱。\n3. 境内/境外与省市一律按「出发地/到达地」的 ICAO 码判定；控制台会先打印映射表，映射不全的航段会自动跳过，不会提交。")

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
