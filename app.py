import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile
import os

# ---------- 辅助函数 ----------
def parse_duration(dur_str: str) -> int:
    """将 'HH:MM' 格式的时间字符串转换为总分钟数"""
    try:
        parts = str(dur_str).split(':')
        if len(parts) == 2:
            hours = int(parts[0])
            minutes = int(parts[1])
            return hours * 60 + minutes
        return 0
    except:
        return 0

def format_duration(total_minutes: int) -> str:
    """将总分钟数格式化为 'HH:MM'（小时不限制位数）"""
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}"

def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col, today_date):
    """
    根据 excel2 的数据更新 excel1 模板
    返回修改后的临时文件路径
    """
    # 加载模板
    wb = load_workbook(excel1_path)
    ws = wb.active

    # 1. 更新 J2：昨日日期
    yesterday = today_date - timedelta(days=1)
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"
    new_j2 = f"*昨日总飞行时间\n（昨日指{yesterday_str}）*"
    ws.cell(row=2, column=10).value = new_j2

    # 2. 计算昨日飞行时间总和（J3）
    total_minutes = 0
    for val in excel2_df[flight_col]:
        if pd.notna(val):
            total_minutes += parse_duration(val)
    new_j3 = format_duration(total_minutes)
    ws.cell(row=3, column=10).value = new_j3

    # 3. 架次（K3）
    num_flights = len(excel2_df)
    ws.cell(row=3, column=11).value = num_flights

    # 4. 累计飞行时间（L3）：旧累计 + 昨日新增
    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    new_total_minutes = old_minutes + total_minutes
    ws.cell(row=3, column=12).value = format_duration(new_total_minutes)

    # 5. 航段信息（N3）：出发-到达，顿号分隔
    segments = []
    for _, row in excel2_df.iterrows():
        dep = str(row[dep_col]).strip() if pd.notna(row[dep_col]) else ''
        arr = str(row[arr_col]).strip() if pd.notna(row[arr_col]) else ''
        if dep and arr:
            segments.append(f"{dep}-{arr}")
        elif dep:          # 只有出发
            segments.append(dep)
        elif arr:          # 只有到达
            segments.append(arr)
    ws.cell(row=3, column=14).value = '、'.join(segments)

    # 6. 飞机注册号数量（O3）：直接取行数（每行一个注册号）
    ws.cell(row=3, column=15).value = num_flights

    # 保存到临时文件
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新", layout="centered")
st.title("🛫 Excel 模板更新工具")
st.markdown("上传 **Excel 1（模板）** 和 **Excel 2（飞行计划）**，程序将根据 Excel 2 自动修改模板中的指定单元格。")

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
    excel2_file = st.file_uploader("📂 上传 Excel 2（飞行计划）", type=["xlsx", "xlsm"])

if excel1_file and excel2_file:
    # 读取 Excel 2 获取列名
    try:
        excel2_df = pd.read_excel(excel2_file, header=0)
    except Exception as e:
        st.error(f"读取 Excel 2 失败：{e}")
        st.stop()

    st.subheader("⚙️ 请指定 Excel 2 中各列的含义")
    cols = excel2_df.columns.tolist()
    flight_col = st.selectbox("🕒 飞行时间列（格式如 34:53）", cols)
    dep_col = st.selectbox("🏙️ 出发城市列", cols)
    arr_col = st.selectbox("🏙️ 到达城市列", cols)
    reg_col = st.selectbox("✈️ 飞机注册号列", cols)  # 实际只用于计数，但保留选择

    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                # 保存模板到临时文件
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                # 执行更新
                output_path = update_excel1(
                    excel1_path,
                    excel2_df,
                    flight_col,
                    dep_col,
                    arr_col,
                    reg_col,
                    today
                )

                # 提供下载
                with open(output_path, 'rb') as f:
                    bytes_data = f.read()
                st.success("✅ 处理完成！点击下方按钮下载更新后的 Excel 1")
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
                # 清理可能残留的临时文件
                if 'excel1_path' in locals() and os.path.exists(excel1_path):
                    os.unlink(excel1_path)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.unlink(output_path)

# ---------- 使用说明 ----------
st.markdown("---")
with st.expander("📖 使用说明"):
    st.markdown("""
    **1. 上传文件**
    - Excel 1：您日常填写的模板文件（.xlsx 或 .xlsm）。
    - Excel 2：昨日的飞行计划数据，需包含飞行时间、出发/到达城市、飞机注册号等列。

    **2. 指定列**
    - 程序会读取 Excel 2 的第一行作为列名，您需从下拉框中选择对应的列。
    - 若列名不符合，可在上传前修改 Excel 2 的表头。

    **3. 自动更新内容**
    - **J2**：昨日日期（根据您选择的“今天”计算）。
    - **J3**：昨日飞行时间总和（HH:MM）。
    - **K3**：飞行架次（Excel 2 的行数）。
    - **L3**：累计飞行时间 = 原累计 + 昨日新增。
    - **N3**：航段信息，格式为“出发-到达”，顿号分隔。
    - **O3**：飞机注册号数量（Excel 2 的行数）。

    **4. 下载**
    - 处理完成后，点击下载按钮即可获取更新后的 Excel 1。
    """)
