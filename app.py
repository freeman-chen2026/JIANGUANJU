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
            return int(parts[0]) * 60 + int(parts[1])
        return 0
    except:
        return 0

def format_duration(total_minutes: int) -> str:
    """将总分钟数格式化为 'HH:MM'（小时不限制位数）"""
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}"

def suggest_column(df, possible_names):
    """从 DataFrame 的列中查找第一个匹配的列名（不区分大小写）"""
    for name in possible_names:
        for col in df.columns:
            if col.strip().lower() == name.lower():
                return col
    return df.columns[0]  # 默认返回第一列

def update_excel1(excel1_path, excel2_df, flight_time_col, dep_city_col, arr_city_col, reg_col, today_date):
    """
    根据 excel2 的数据更新 excel1 模板
    """
    wb = load_workbook(excel1_path)
    ws = wb.active

    # 1. 更新 J2：昨日日期
    yesterday = today_date - timedelta(days=1)
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday_str}）*"

    # 2. 计算昨日飞行时间总和（J3）
    total_minutes = 0
    for val in excel2_df[flight_time_col]:
        if pd.notna(val):
            total_minutes += parse_duration(val)
    new_j3 = format_duration(total_minutes)
    ws.cell(row=3, column=10).value = new_j3

    # 3. 架次（K3）：总行数
    num_flights = len(excel2_df)
    ws.cell(row=3, column=11).value = num_flights

    # 4. 累计飞行时间（L3）：旧累计 + 昨日新增
    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    new_total_minutes = old_minutes + total_minutes
    ws.cell(row=3, column=12).value = format_duration(new_total_minutes)

    # 5. 航段信息（N3）：出发城市-到达城市，顿号分隔
    segments = []
    for _, row in excel2_df.iterrows():
        dep = str(row[dep_city_col]).strip() if pd.notna(row[dep_city_col]) else ''
        arr = str(row[arr_city_col]).strip() if pd.notna(row[arr_city_col]) else ''
        if dep and arr:
            segments.append(f"{dep}-{arr}")
        elif dep:
            segments.append(dep)   # 只有出发
        elif arr:
            segments.append(arr)   # 只有到达
    ws.cell(row=3, column=14).value = '、'.join(segments)

    # 6. 飞机注册号数量（O3）：去重后的注册号个数
    unique_regs = excel2_df[reg_col].dropna().unique()
    ws.cell(row=3, column=15).value = len(unique_regs)

    # 保存到临时文件
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return tmp.name


# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="Excel 模板自动更新（航段版）", layout="centered")
st.title("🛩️ 航段数据 -> 模板更新工具")
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

    flight_col = st.selectbox("🕒 飞行时间列（格式如 09:01）", cols, index=cols.index(default_flight))
    dep_col = st.selectbox("🏙️ 出发城市列", cols, index=cols.index(default_dep))
    arr_col = st.selectbox("🏙️ 到达城市列", cols, index=cols.index(default_arr))
    reg_col = st.selectbox("✈️ 飞机注册号列（用于去重统计）", cols, index=cols.index(default_reg))

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
                if 'excel1_path' in locals() and os.path.exists(excel1_path):
                    os.unlink(excel1_path)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.unlink(output_path)

# ---------- 使用说明 ----------
st.markdown("---")
with st.expander("📖 使用说明（重要）"):
    st.markdown("""
    **1. 上传文件**
    - **Excel 1**：您日常填写的模板文件（.xlsx 或 .xlsm），必须包含第2、3行，且 J~O 列为待更新区域。
    - **Excel 2**：航段数据导出文件（如 `航段数据导出 (28).xlsx`），需包含以下列：
      - 飞行时间（如 `实际飞行时间`）
      - 出发城市
      - 到达城市
      - 飞机注册号

    **2. 选择列**
    - 程序会自动识别常见列名，您只需检查确认，如有偏差可手动调整。

    **3. 自动更新内容**
    - **J2**：昨日日期（根据您选择的“今天”计算）。
    - **J3**：昨日所有航段的 **实际飞行时间** 总和（HH:MM）。
    - **K3**：航段总条数（架次）。
    - **L3**：累计飞行时间 = 模板中原累计值 + 昨日新增总和。
    - **N3**：所有航段的 **出发城市-到达城市**，用中文顿号分隔（例如：`泉州晋江-美国安克雷奇 史蒂文斯、大连周水子-上海虹桥`）。
    - **O3**：**去重后的飞机注册号数量**（即使用了多少架不同的飞机）。

    **4. 下载**
    - 处理完成后，点击下载按钮即可获取更新后的 Excel 1。
    """)
