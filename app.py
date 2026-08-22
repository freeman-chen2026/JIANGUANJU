import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile
import os
import re

# ---------- 辅助函数 ----------
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

def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col):
    wb = load_workbook(excel1_path)
    ws = wb.active

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday.month}月{yesterday.day}日）*"

    valid_mask = excel2_df[flight_col].notna()
    valid_df = excel2_df[valid_mask].copy()

    total_minutes = 0
    time_parsed = []
    for idx, val in valid_df[flight_col].items():
        mins = parse_duration(val)
        time_parsed.append((idx, val, mins))
        total_minutes += mins

    ws.cell(row=3, column=10).value = format_duration(total_minutes)
    ws.cell(row=3, column=11).value = len(valid_df)

    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    ws.cell(row=3, column=12).value = format_duration(old_minutes + total_minutes)

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
        '昨日总分钟': total_minutes,
        '架次': len(valid_df),
        '旧累计分钟': old_minutes,
        '新累计分钟': old_minutes + total_minutes,
        '航段数': len(segments),
        '注册号数量': len(reg_series.astype(str).unique()),
        '注册号列表': list(reg_series.astype(str).unique()),
        '时间解析详情': time_parsed,
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新（航段版）", layout="wide")
st.title("🛩️ 航段数据 → 模板更新工具（自动表头检测）")
st.markdown("上传 **昨日飞行数据（operation-每日飞行数据）** 和 **航段数据导出**，程序自动检测表头并计算昨日日期。")

col1, col2 = st.columns(2)
with col1:
    excel1_file = st.file_uploader("📂 上传昨日飞行数据（operation-每日飞行数据）", type=["xlsx", "xlsm"])
with col2:
    excel2_file = st.file_uploader("📂 上传航段数据导出", type=["xlsx", "xlsm"])

if excel1_file and excel2_file:
    # 读取 Excel 2，自动检测表头
    try:
        keywords = ["客户", "航班号", "出发城市", "到达城市", "实际飞行时间", "飞机注册号"]
        excel2_df = read_excel_with_auto_header(excel2_file, keywords)
    except Exception as e:
        st.error(f"读取 Excel 2 失败：{e}")
        st.stop()

    if excel2_df.empty or len(excel2_df.columns) == 0:
        st.error("Excel 2 没有有效数据或列，请检查文件。")
        st.stop()

    valid_cols = excel2_df.columns.tolist()
    st.subheader("⚙️ 请指定 Excel 2 中各列的含义（务必仔细核对）")
    
    def suggest(cols, candidates):
        for cand in candidates:
            for c in cols:
                if cand in c:
                    return c
        return cols[0]

    default_flight = suggest(valid_cols, ["实际飞行时间", "飞行时间", "航段时间"])
    default_dep = suggest(valid_cols, ["出发城市", "起飞机场", "出发地"])
    default_arr = suggest(valid_cols, ["到达城市", "目的地机场", "到达地"])
    default_reg = suggest(valid_cols, ["飞机注册号", "注册号", "机号"])

    flight_col = st.selectbox("🕒 飞行时间列", valid_cols, index=valid_cols.index(default_flight))
    dep_col = st.selectbox("🏙️ 出发城市列", valid_cols, index=valid_cols.index(default_dep))
    arr_col = st.selectbox("🏙️ 到达城市列", valid_cols, index=valid_cols.index(default_arr))
    reg_col = st.selectbox("✈️ 飞机注册号列", valid_cols, index=valid_cols.index(default_reg))

    st.subheader("📊 数据预览（所选列的前5行）")
    preview_df = excel2_df[[flight_col, dep_col, arr_col, reg_col]].head(5)
    st.dataframe(preview_df)

    valid_count = excel2_df[flight_col].notna().sum()
    st.info(f"📌 飞行时间非空的航段数：{valid_count} 条")

    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理..."):
            try:
                # 计算昨日日期用于文件名
                today = datetime.now().date()
                yesterday = today - timedelta(days=1)
                month = yesterday.month
                day = yesterday.day

                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                output_path, stats = update_excel1(
                    excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col
                )

                st.success("✅ 处理完成！")
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("昨日飞行时间 (分钟)", stats['昨日总分钟'])
                    st.metric("格式化 (J3)", format_duration(stats['昨日总分钟']))
                with c2:
                    st.metric("架次 (K3)", stats['架次'])
                    st.metric("去重注册号 (O3)", stats['注册号数量'])
                with c3:
                    st.metric("旧累计分钟", stats['旧累计分钟'])
                    st.metric("新累计 (L3)", format_duration(stats['新累计分钟']))

                st.write("注册号列表：", ", ".join(stats['注册号列表']))

                with st.expander("🔍 时间解析详情"):
                    df_time = pd.DataFrame(stats['时间解析详情'], columns=['行索引', '原始值', '解析分钟数'])
                    st.dataframe(df_time)

                with open(output_path, 'rb') as f:
                    file_data = f.read()

                # 生成动态文件名
                file_name = f"中南-深圳局-天成商务航空有限公司-{month}月{day}日飞行数据.xlsx"
                st.download_button(
                    label="📥 下载 Excel 1（已更新）",
                    data=file_data,
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

st.markdown("---")
with st.expander("📖 使用说明"):
    st.markdown("""
    **操作步骤**
    1. 上传 **昨日飞行数据（operation-每日飞行数据）**（即模板文件）和 **航段数据导出**。
    2. 程序自动检测航段数据的表头（寻找“客户”、“航班号”等关键词），清理无效列。
    3. 在下拉列表中选择对应的列：飞行时间、出发城市、到达城市、飞机注册号。
    4. 点击“开始处理”，查看统计信息并下载更新后的文件。
    5. 下载的文件名自动包含昨日的月份和日期，例如“中南-深圳局-天成商务航空有限公司-8月21日飞行数据.xlsx”。
    """)
