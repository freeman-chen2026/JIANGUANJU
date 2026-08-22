import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile
import os

# ---------- 辅助函数 ----------
def parse_duration(dur_str) -> int:
    """将 'HH:MM' 或 'HH：MM' 格式的时间字符串转换为总分钟数"""
    if pd.isna(dur_str):
        return 0
    s = str(dur_str).strip()
    s = s.replace('：', ':')  # 中文冒号转英文
    parts = s.split(':')
    if len(parts) == 2:
        try:
            return int(parts[0]) * 60 + int(parts[1])
        except:
            return 0
    return 0

def format_duration(total_minutes: int) -> str:
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}"

def suggest_column(df, candidates):
    """从 DataFrame 的列中查找第一个匹配的列名（不区分大小写和空格）"""
    cols_lower = {col.strip().lower(): col for col in df.columns}
    for cand in candidates:
        cand_lower = cand.strip().lower()
        if cand_lower in cols_lower:
            return cols_lower[cand_lower]
    return df.columns[0]

def update_excel1(excel1_path, excel2_df, flight_col, dep_city_col, arr_city_col, reg_col, today_date):
    """
    根据 excel2 的数据更新 excel1 模板，返回 (输出文件路径, 统计信息字典)
    """
    wb = load_workbook(excel1_path)
    ws = wb.active

    # 1. 更新 J2：昨日日期
    yesterday = today_date - timedelta(days=1)
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday_str}）*"

    # 2. 计算昨日飞行时间总和（J3），仅对非空时间行求和
    valid_times = excel2_df[flight_col].dropna()
    total_minutes = 0
    for val in valid_times:
        total_minutes += parse_duration(val)
    new_j3 = format_duration(total_minutes)
    ws.cell(row=3, column=10).value = new_j3

    # 3. 架次（K3）：有效航段数（有时间值的行数）
    num_flights = len(valid_times)
    ws.cell(row=3, column=11).value = num_flights

    # 4. 累计飞行时间（L3）：旧累计 + 昨日新增
    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    new_total_minutes = old_minutes + total_minutes
    ws.cell(row=3, column=12).value = format_duration(new_total_minutes)

    # 5. 航段信息（N3）：出发城市-到达城市，顿号分隔，只取有时间值的行（保持与架次一致）
    segments = []
    for idx, row in excel2_df.iterrows():
        # 检查该行是否有有效飞行时间（非空）
        if pd.isna(row[flight_col]):
            continue
        dep = str(row[dep_city_col]).strip() if pd.notna(row[dep_city_col]) else ''
        arr = str(row[arr_city_col]).strip() if pd.notna(row[arr_city_col]) else ''
        if dep and arr:
            segments.append(f"{dep}-{arr}")
        elif dep:
            segments.append(dep)
        elif arr:
            segments.append(arr)
    ws.cell(row=3, column=14).value = '、'.join(segments)

    # 6. 飞机注册号数量（O3）：去重后的注册号个数（只统计有时间值的航段）
    # 提取注册号，去除空值，去重
    reg_series = excel2_df.loc[valid_times.index, reg_col]  # 只取有时间值的行
    unique_regs = reg_series.dropna().astype(str).unique()
    ws.cell(row=3, column=15).value = len(unique_regs)

    # 收集统计信息用于预览
    stats = {
        '昨日总分钟': total_minutes,
        '架次': num_flights,
        '旧累计分钟': old_minutes,
        '新累计分钟': new_total_minutes,
        '航段数': len(segments),
        '注册号去重数量': len(unique_regs),
        '注册号列表': list(unique_regs),
    }

    # 保存到临时文件
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name, stats


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新（航段版）", layout="centered")
st.title("🛩️ 航段数据 → 模板更新工具")
st.markdown("上传 **Excel 1（模板）** 和 **Excel 2（航段导出）**，程序将按业务规则自动更新模板。")

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
    # 读取 Excel 2 获取列名
    try:
        excel2_df = pd.read_excel(excel2_file, header=0)
    except Exception as e:
        st.error(f"读取 Excel 2 失败：{e}")
        st.stop()

    st.subheader("⚙️ 请指定 Excel 2 中各列的含义")
    cols = excel2_df.columns.tolist()

    # 自动建议列名
    default_flight = suggest_column(excel2_df, ["实际飞行时间", "飞行时间", "航段时间"])
    default_dep = suggest_column(excel2_df, ["出发城市", "起飞机场", "出发地"])
    default_arr = suggest_column(excel2_df, ["到达城市", "目的地机场", "到达地"])
    default_reg = suggest_column(excel2_df, ["飞机注册号", "注册号", "机号"])

    flight_col = st.selectbox("🕒 飞行时间列（格式如 09:01）", cols, index=cols.index(default_flight) if default_flight in cols else 0)
    dep_col = st.selectbox("🏙️ 出发城市列", cols, index=cols.index(default_dep) if default_dep in cols else 0)
    arr_col = st.selectbox("🏙️ 到达城市列", cols, index=cols.index(default_arr) if default_arr in cols else 0)
    reg_col = st.selectbox("✈️ 飞机注册号列（用于去重统计）", cols, index=cols.index(default_reg) if default_reg in cols else 0)

    # 显示数据预览
    with st.expander("📊 数据预览（前5行）"):
        st.dataframe(excel2_df.head())

    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                # 保存模板到临时文件
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                # 执行更新，获取统计信息
                output_path, stats = update_excel1(
                    excel1_path,
                    excel2_df,
                    flight_col,
                    dep_col,
                    arr_col,
                    reg_col,
                    today
                )

                # 显示处理统计
                st.success("✅ 处理完成！以下是本次更新的统计信息：")
                col_stats1, col_stats2 = st.columns(2)
                with col_stats1:
                    st.metric("昨日飞行时间（总分钟）", stats['昨日总分钟'])
                    st.metric("架次（航段数）", stats['架次'])
                    st.metric("旧累计分钟", stats['旧累计分钟'])
                with col_stats2:
                    st.metric("新累计分钟", stats['新累计分钟'])
                    st.metric("航段数（拼接）", stats['航段数'])
                    st.metric("去重注册号数量", stats['注册号去重数量'])
                    st.write("去重注册号列表：", ", ".join(stats['注册号列表']))

                # 提供下载
                with open(output_path, 'rb') as f:
                    bytes_data = f.read()
                st.download_button(
                    label="📥 下载 Excel 1（已更新）",
                    data=bytes_data,
                    file_name="updated_excel1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # 清理临时文件
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
with st.expander("📖 使用说明与注意事项"):
    st.markdown("""
    **1. 上传文件**
    - **Excel 1**：您的模板文件（如 `中南-深圳局-天成商务航空有限公司-8月21日飞行数据.xlsx`）。
    - **Excel 2**：航段数据导出文件（如 `航段数据导出 (28).xlsx`）。

    **2. 选择列**
    - 程序会自动匹配常见列名，您也可以手动调整。
    - 请确保飞行时间列是 `HH:MM` 格式（支持中文冒号）。

    **3. 处理逻辑（关键）**
    - 只统计飞行时间非空的航段（架次、时间、航段拼接均基于这些有效航段）。
    - 注册号去重统计仅针对这些有效航段，且去除空值。
    - 累计时间 = 模板中的原累计时间 + 昨日新增总时间。

    **4. 预览与验证**
    - 处理后会显示统计信息，便于您核对是否正确。
    - 如果发现注册号数量不符，请检查注册号列是否有空值或重复。

    **5. 下载**
    - 点击下载按钮获取更新后的 Excel 1。
    """)
