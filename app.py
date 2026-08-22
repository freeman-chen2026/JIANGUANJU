import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile
import os
import re

# ---------- 辅助函数 ----------
def parse_duration(dur_str) -> int:
    """将 'HH:MM' 或 'HH：MM' 格式的时间字符串转换为总分钟数"""
    if pd.isna(dur_str):
        return 0
    s = str(dur_str).strip()
    # 统一为英文冒号
    s = s.replace('：', ':')
    # 提取数字部分
    match = re.search(r'(\d+):(\d{2})', s)
    if match:
        try:
            hours = int(match.group(1))
            minutes = int(match.group(2))
            return hours * 60 + minutes
        except:
            return 0
    return 0

def format_duration(total_minutes: int) -> str:
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}"

def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col, today_date):
    """
    根据 excel2 的数据更新 excel1 模板
    返回 (输出文件路径, 统计信息字典)
    """
    wb = load_workbook(excel1_path)
    ws = wb.active

    # 1. 更新 J2：昨日日期
    yesterday = today_date - timedelta(days=1)
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday_str}）*"

    # 2. 筛选有效航段（飞行时间非空）
    valid_mask = excel2_df[flight_col].notna()
    valid_df = excel2_df[valid_mask].copy()

    # 计算飞行时间总和
    total_minutes = 0
    time_parsed = []  # 用于调试
    for idx, val in valid_df[flight_col].items():
        mins = parse_duration(val)
        time_parsed.append((idx, val, mins))
        total_minutes += mins

    new_j3 = format_duration(total_minutes)
    ws.cell(row=3, column=10).value = new_j3

    # 3. 架次（K3）
    num_flights = len(valid_df)
    ws.cell(row=3, column=11).value = num_flights

    # 4. 累计飞行时间（L3）
    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    new_total_minutes = old_minutes + total_minutes
    ws.cell(row=3, column=12).value = format_duration(new_total_minutes)

    # 5. 航段信息（N3）
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

    # 6. 飞机注册号数量（O3）
    reg_series = valid_df[reg_col].dropna()
    unique_regs = reg_series.astype(str).unique()
    ws.cell(row=3, column=15).value = len(unique_regs)

    # 统计信息
    stats = {
        '昨日总分钟': total_minutes,
        '架次': num_flights,
        '旧累计分钟': old_minutes,
        '新累计分钟': new_total_minutes,
        '航段数': len(segments),
        '去重注册号数量': len(unique_regs),
        '注册号列表': list(unique_regs),
        '时间解析详情': time_parsed,  # 用于调试
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新（航段版）", layout="wide")
st.title("🛩️ 航段数据 → 模板更新工具（最终版）")
st.markdown("上传 **Excel 1（模板）** 和 **Excel 2（航段数据）**，**请务必手动选择正确的列**，然后处理。")

# 日期选择
today = st.date_input(
    "📅 选择今天的日期（用于计算“昨日”）",
    datetime.now().date(),
    help="程序将自动计算昨日日期并填入模板"
)

# 文件上传
col1, col2 = st.columns(2)
with col1:
    excel1_file = st.file_uploader("📂 上传 Excel 1（模板）", type=["xlsx", "xlsm"])
with col2:
    excel2_file = st.file_uploader("📂 上传 Excel 2（航段数据）", type=["xlsx", "xlsm"])

if excel1_file and excel2_file:
    # 读取 Excel 2
    try:
        excel2_df = pd.read_excel(excel2_file, header=0)
    except Exception as e:
        st.error(f"读取 Excel 2 失败：{e}")
        st.stop()

    st.subheader("⚙️ 请指定 Excel 2 中各列的含义（务必仔细核对）")
    cols = excel2_df.columns.tolist()

    # 不使用自动建议，让用户手动选择（避免误判）
    flight_col = st.selectbox("🕒 飞行时间列（格式如 09:01，例如“实际飞行时间”）", cols)
    dep_col = st.selectbox("🏙️ 出发城市列（例如“出发城市”）", cols)
    arr_col = st.selectbox("🏙️ 到达城市列（例如“到达城市”）", cols)
    reg_col = st.selectbox("✈️ 飞机注册号列（例如“飞机注册号”）", cols)

    # 预览所选列的内容
    st.subheader("📊 数据预览（所选列的前5行）")
    preview_df = excel2_df[[flight_col, dep_col, arr_col, reg_col]].head(5)
    st.dataframe(preview_df)

    # 显示有效航段数量（飞行时间非空）
    valid_count = excel2_df[flight_col].notna().sum()
    st.info(f"📌 飞行时间非空的航段数：{valid_count} 条")

    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                # 保存模板
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                # 执行更新
                output_path, stats = update_excel1(
                    excel1_path,
                    excel2_df,
                    flight_col,
                    dep_col,
                    arr_col,
                    reg_col,
                    today
                )

                # 显示详细统计
                st.success("✅ 处理完成！以下是本次更新的统计信息：")

                col_stats1, col_stats2, col_stats3 = st.columns(3)
                with col_stats1:
                    st.metric("昨日飞行时间（总分钟）", stats['昨日总分钟'])
                    st.metric("格式化后 (J3)", format_duration(stats['昨日总分钟']))
                with col_stats2:
                    st.metric("架次 (K3)", stats['架次'])
                    st.metric("去重注册号数量 (O3)", stats['去重注册号数量'])
                with col_stats3:
                    st.metric("旧累计分钟", stats['旧累计分钟'])
                    st.metric("新累计分钟 (L3)", stats['新累计分钟'])

                # 显示注册号列表
                st.write("去重注册号列表：", ", ".join(stats['注册号列表']))

                # 显示时间解析详情（调试用）
                with st.expander("🔍 查看每个航段的时间解析详情（用于排查）"):
                    df_time = pd.DataFrame(stats['时间解析详情'], columns=['行索引', '原始值', '解析分钟数'])
                    st.dataframe(df_time)

                # 下载
                with open(output_path, 'rb') as f:
                    bytes_data = f.read()
                st.download_button(
                    label="📥 下载 Excel 1（已更新）",
                    data=bytes_data,
                    file_name="updated_excel1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # 清理
                os.unlink(excel1_path)
                os.unlink(output_path)

            except Exception as e:
                st.error(f"处理出错：{e}")
                if 'excel1_path' in locals() and os.path.exists(excel1_path):
                    os.unlink(excel1_path)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.unlink(output_path)

# ---------- 使用说明 ----------
st.markdown("---")
with st.expander("📖 使用说明（必读）"):
    st.markdown("""
    **为什么之前的更新不正确？**
    - 很可能是您选择的列不对。例如，将“飞行时间”选成了其他列，或者将“出发城市”选成了“客户”。
    - 本版本**强制您手动选择每一列**，并且会预览所选列的数据，确保万无一失。

    **操作步骤：**
    1. 上传两个 Excel 文件。
    2. 从下拉框中为 **飞行时间、出发城市、到达城市、飞机注册号** 分别选择正确的列名。
    3. 观察下方的数据预览，确认所选列的内容正确（如飞行时间应为 `09:01` 格式）。
    4. 点击“开始处理”，查看统计结果，下载更新后的文件。

    **常见问题排查：**
    - 如果 J3 仍为 0:00，请检查“飞行时间列”是否选对，且数据是否包含 `HH:MM` 格式。
    - 如果 N3 仍显示人名，说明您把“出发城市”或“到达城市”错选为“客户”列了。
    - 如果 O3 数量不对，请检查“飞机注册号列”是否选对，并查看“去重注册号列表”是否合理。
    - 处理前，请确保 Excel 2 中所有已执飞航段都有飞行时间（非空），否则会被忽略。
    """)
