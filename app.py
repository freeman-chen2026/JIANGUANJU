import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile
import os
import re

# ==============================
# 功能 A：每日飞行数据自动更新（模板 J3、K3、L3、N3、O3）
# ==============================
def parse_duration(dur_str) -> int:
    if pd.isna(dur_str):
        return 0
    s = str(dur_str).strip().replace('：', ':')
    match = re.search(r'(\d+):(\d{2})', s)
    if match:
        try:
            return int(match.group(1)) * 60 + int(match.group(2))
        except:
            return 0
    return 0

def format_duration(total_minutes: int) -> str:
    return f"{total_minutes // 60}:{total_minutes % 60:02d}"

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

def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col):
    wb = load_workbook(excel1_path)
    ws = wb.active

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday.month}月{yesterday.day}日）*"

    valid_mask = excel2_df[flight_col].notna()
    valid_df = excel2_df[valid_mask].copy()

    total_minutes = 0
    for val in valid_df[flight_col]:
        total_minutes += parse_duration(val)

    ws.cell(row=3, column=10).value = format_duration(total_minutes)
    ws.cell(row=3, column=11).value = len(valid_df)

    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    new_total = old_minutes + total_minutes
    ws.cell(row=3, column=12).value = format_duration(new_total)

    segments = []
    for _, row in valid_df.iterrows():
        dep = str(row[dep_col]).strip() if pd.notna(row[dep_col]) else ''
        arr = str(row[arr_col]).strip() if pd.notna(row[arr_col]) else ''
        if dep and arr:
            segments.append(f"{dep}-{arr}")
        elif dep:
            segments.append(dep)
        elif arr:
            segments.append(arr)
    ws.cell(row=3, column=14).value = '、'.join(segments)

    reg_series = valid_df[reg_col].dropna()
    ws.cell(row=3, column=15).value = len(reg_series.astype(str).unique())

    stats = {
        '昨日飞行时间': format_duration(total_minutes),
        '架次': len(valid_df),
        '注册号数量': len(reg_series.astype(str).unique()),
        '截止昨日总飞行时间': format_duration(old_minutes),
        '截止今日总飞行时间': format_duration(new_total),
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats

def run_feature_a():
    excel1_file = st.file_uploader("📂 上传昨日飞行数据（operation-每日飞行数据）", type=["xlsx", "xlsm"], key="a_excel1")
    excel2_file = st.file_uploader("📂 上传：航段数据导出", type=["xlsx", "xlsm"], key="a_excel2")

    if excel1_file and excel2_file:
        with st.spinner("正在自动处理..."):
            try:
                keywords = ["客户", "航班号", "出发城市", "到达城市", "实际飞行时间", "飞机注册号"]
                excel2_df = read_excel_with_auto_header(excel2_file, keywords)
                if excel2_df.empty or len(excel2_df.columns) == 0:
                    st.error("航段数据没有有效列，请检查文件格式。")
                    return

                flight_col = auto_match_column(excel2_df, ["实际飞行时间", "飞行时间", "航段时间"])
                dep_col = auto_match_column(excel2_df, ["出发城市", "起飞机场", "出发地"])
                arr_col = auto_match_column(excel2_df, ["到达城市", "目的地机场", "到达地"])
                reg_col = auto_match_column(excel2_df, ["飞机注册号", "注册号", "机号"])

                missing = []
                if not flight_col: missing.append("飞行时间")
                if not dep_col: missing.append("出发城市")
                if not arr_col: missing.append("到达城市")
                if not reg_col: missing.append("飞机注册号")
                if missing:
                    st.error(f"未能自动匹配以下列：{', '.join(missing)}，请检查文件列名是否包含关键词。")
                    return

                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                output_path, stats = update_excel1(
                    excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col
                )

                st.success("✅ 处理完成")
                col1, col2, col3 = st.columns(3)
                col1.metric("昨日总飞行时间", stats['昨日飞行时间'])
                col2.metric("架次", stats['架次'])
                col3.metric("使用航空器数量", stats['注册号数量'])

                col4, col5 = st.columns(2)
                col4.metric("截止昨日总飞行时间", stats['截止昨日总飞行时间'])
                col5.metric("截止今日总飞行时间", stats['截止今日总飞行时间'])

                yesterday = datetime.now().date() - timedelta(days=1)
                file_name = f"中南-深圳局-天成商务航空有限公司-{yesterday.month}月{yesterday.day}日飞行数据.xlsx"

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
# 功能 B：模板生成备案表（原代码）
# ==============================
# ---------- 注册号 -> ICAO 机型映射 ----------
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

# ---------- 辅助函数 ----------
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

def parse_duration_to_minutes(duration_str):
    if pd.isna(duration_str) or duration_str == "" or duration_str is None:
        return None
    s = str(duration_str).strip()
    try:
        if s.count(":") == 1:
            parts = s.split(":")
            if len(parts) == 2:
                h = int(parts[0])
                m = int(parts[1])
                return h * 60 + m
        elif s.count(":") == 2:
            parts = s.split(":")
            h = int(parts[0])
            m = int(parts[1])
            return h * 60 + m
        else:
            return int(float(s))
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
    # ---------- 侧边栏：自定义固定值 ----------
    st.sidebar.header("✏️ 自定义固定填入值（A、B列）")
    default_supervision = st.sidebar.text_input("所属监管局（A列）", value="深圳局", key="b_supervision")
    default_operator = st.sidebar.text_input("运行人标准名称（B列）", value="天成商务航空有限公司", key="b_operator")

    st.sidebar.header("✈️ 注册号 → ICAO 机型映射")
    user_mapping_text = st.sidebar.text_area(
        "自定义映射（覆盖内置）",
        value="\n".join([f"{k} {v}" for k, v in DEFAULT_ICAO_MAP.items()]),
        height=150,
        key="b_mapping"
    )
    def parse_mapping(text):
        map_dict = {}
        for line in text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) >= 2:
                reg, icao = parts[0], parts[1]
                map_dict[reg.upper()] = icao.upper()
        return map_dict
    icao_map = parse_mapping(user_mapping_text)
    for k, v in DEFAULT_ICAO_MAP.items():
        if k not in icao_map:
            icao_map[k] = v

    # ---------- 初始化 session_state ----------
    if "template_wb" not in st.session_state:
        st.session_state.template_wb = None
    if "header_row" not in st.session_state:
        st.session_state.header_row = None
    if "data_start_row" not in st.session_state:
        st.session_state.data_start_row = None
    if "group_rows" not in st.session_state:
        st.session_state.group_rows = 4  # 固定每组4行

    # ---------- 模板管理 ----------
    st.subheader("📂 模板管理")
    if st.session_state.template_wb is None:
        st.info("首次使用请上传模板文件。")
        template_file = st.file_uploader("上传：桌面-Jetops申请一览-每日通航运行情况跟踪表（天成商务航空有限公司）20260708", type=["xlsx"], key="b_template_upload")
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

    # ---------- 数据上传 ----------
    st.subheader("📊 数据上传")
    data_file = st.file_uploader("上传：航段数据导出", type=["xlsx"], key="b_data_upload")

    if data_file and st.session_state.template_wb is not None:
        try:
            df_raw = parse_uploaded_file(data_file)
            st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")
            if len(df_raw) > 20:
                st.warning(f"数据条数（{len(df_raw)}）超过模板预设的20行，多余数据将被忽略。")

            # ---- 统计信息（替换原预览表格） ----
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

            # ---- 生成汇报文案（新规则） ----
            parts = []
            if landed_count > 0:
                parts.append(f"{landed_count}班已落地")
            if unlanded_count > 0:
                parts.append(f"{unlanded_count}班未落地")
            if not_executed_count > 0:
                parts.append(f"{not_executed_count}班未起飞")

            if not parts:  # 理论上不会，但预防
                report = "今天无飞行计划"
            else:
                if landed_count > 0 and unlanded_count == 0 and not_executed_count == 0:
                    report = "今天飞完了"
                else:
                    report = "今天" + "、".join(parts)

            # 飞机列表
            if reg_list:
                plane_text = f"今日有飞行计划的飞机：{'、'.join(reg_list)}"
            else:
                plane_text = ""

            # 显示统计信息
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

            # 汇报文案区域
            full_report = report
            if plane_text:
                full_report += "\n" + plane_text
            st.text_area("📋 汇报文案（可复制）", full_report, height=120)

            # ---- 继续原有数据填充逻辑 ----
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
# 主界面：使用选项卡切换两个功能
# ==============================
st.set_page_config(page_title="飞行数据工具组合", layout="wide")
st.title("🛩️ 飞行数据工具组合")

tab1, tab2 = st.tabs(["📋 每日飞行数据-10：00发", "📄 每日通航运行情况跟踪表-16：30发"])

with tab1:
    run_feature_a()

with tab2:
    run_feature_b()

st.caption("💡 功能1自动更新模板的汇总数据，功能2按航段逐行填写备案表。")
