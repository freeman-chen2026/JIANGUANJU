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
        '时间解析详情': time_parsed,
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新（航段版）", layout="wide")
st.title("🛩️ 航段数据 → 模板更新工具（最终稳健版）")
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

    # 过滤掉无效列（如 Unnamed 开头的列）
    all_cols = excel2_df.columns.tolist()
    valid_cols = [c for c in all_cols if not c.startswith('Unnamed') and not c.startswith('unnamed')]
    if not valid_cols:
        st.warning("未检测到有效列名，请确保 Excel 2 的第一行为表头。")
        valid_cols = all_cols  # 如果全部无效，则使用所有列

    st.subheader("⚙️ 请指定 Excel 2 中各列的含义（务必仔细核对）")
    # 根据数据内容建议默认选项（但不强制）
    def suggest(cols, keywords):
        for kw in keywords:
            for c in cols:
                if kw in c:
                    return c
        return cols[0] if cols else None

    default_flight = suggest(valid_cols, ["实际飞行时间", "飞行时间", "航段时间"])
    default_dep = suggest(valid_cols, ["出发城市", "起飞机场", "出发地"])
    default_arr = suggest(valid_cols, ["到达城市", "目的地机场", "到达地"])
    default_reg = suggest(valid_cols, ["飞机注册号", "注册号", "机号"])

    flight_col = st.selectbox("🕒 飞行时间列（格式如 09:01）", valid_cols, index=valid_cols.index(default_flight) if default_flight in valid_cols else 0)
    dep_col = st.selectbox("🏙️ 出发城市列", valid_cols, index=valid_cols.index(default_dep) if default_dep in valid_cols else 0)
    arr_col = st.selectbox("🏙️ 到达城市列", valid_cols, index=valid_cols.index(default_arr) if default_arr in valid_cols else 0)
    reg_col = st.selectbox("✈️ 飞机注册号列（用于去重统计）", valid_cols, index=valid_cols.index(default_reg) if default_reg in valid_cols else 0)

    # 预览所选列的内容（加入空值检查）
    st.subheader("📊 数据预览（所选列的前5行）")
    try:
        preview_df = excel2_df[[flight_col, dep_col, arr_col, reg_col]].head(5)
        st.dataframe(preview_df)
    except Exception as e:
        st.error(f"预览失败，可能是所选列包含重复列名或不存在的列。请检查列选择。错误：{e}")

    # 显示有效航段数量
    if flight_col in excel2_df.columns:
        valid_count = excel2_df[flight_col].notna().sum()
        st.info(f"📌 飞行时间非空的航段数：{valid_count} 条")
    else:
        st.error("所选飞行时间列不存在，请重新选择。")

    if st.button("🚀 开始处理", type="primary"):
        if flight_col not in excel2_df.columns:
            st.error("飞行时间列无效，请重新选择。")
            st.stop()

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

                # 显示时间解析详情（调试）
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
    - 很可能您选择了错误的列（例如将 `飞行时间` 选成了 `客户` 或 `Unnamed` 列）。
    - 本版本自动过滤无效列（如 `Unnamed`），并让您手动选择正确的列。

    **操作步骤：**
    1. 上传两个 Excel 文件。
    2. 从下拉框中为 **飞行时间、出发城市、到达城市、飞机注册号** 分别选择正确的列名。
    3. 观察下方的数据预览，确认所选列的内容正确（如飞行时间应为 `09:01` 格式）。
    4. 点击“开始处理”，查看统计结果，下载更新后的文件。

    **常见问题排查：**
    - 如果 J3 仍为 0:00，请检查“飞行时间列”是否选对，且数据是否包含 `HH:MM` 格式。展开“时间解析详情”可看到每个航段的解析结果。
    - 如果 N3 显示错误内容，说明您把“出发城市”或“到达城市”错选了其他列。
    - 如果 O3 数量不对，请检查“飞机注册号列”是否选对，并查看“去重注册号列表”。
    - 如果预览报错，请确保所选列在数据中存在，且列名不重复。
    """)
