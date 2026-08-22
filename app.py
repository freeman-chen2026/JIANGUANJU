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

def update_excel1(excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col):
    """
    根据 excel2 的数据更新 excel1 模板
    """
    wb = load_workbook(excel1_path)
    ws = wb.active

    # 1. 更新 J2：昨日日期（自动计算）
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    yesterday_str = f"{yesterday.month}月{yesterday.day}日"
    ws.cell(row=2, column=10).value = f"*昨日总飞行时间\n（昨日指{yesterday_str}）*"

    # 2. 筛选有效航段（飞行时间非空）
    valid_mask = excel2_df[flight_col].notna()
    valid_df = excel2_df[valid_mask].copy()

    # 计算飞行时间总和
    total_minutes = 0
    time_parsed = []
    for idx, val in valid_df[flight_col].items():
        mins = parse_duration(val)
        time_parsed.append((idx, val, mins))
        total_minutes += mins

    ws.cell(row=3, column=10).value = format_duration(total_minutes)   # J3
    ws.cell(row=3, column=11).value = len(valid_df)                   # K3

    # 累计飞行时间 L3
    old_l3 = ws.cell(row=3, column=12).value
    old_minutes = parse_duration(old_l3) if old_l3 is not None else 0
    ws.cell(row=3, column=12).value = format_duration(old_minutes + total_minutes)

    # 航段信息 N3
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

    # 注册号数量 O3
    reg_series = valid_df[reg_col].dropna()
    ws.cell(row=3, column=15).value = len(reg_series.astype(str).unique())

    # 统计信息（用于显示）
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
st.title("🛩️ 航段数据 → 模板更新工具")
st.markdown("上传 **Excel 1（模板）** 和 **Excel 2（航段数据）**，程序自动计算昨日日期，请手动选择正确的列。")

# 文件上传
col1, col2 = st.columns(2)
with col1:
    excel1_file = st.file_uploader("📂 上传 Excel 1（模板）", type=["xlsx", "xlsm"])
with col2:
    excel2_file = st.file_uploader("📂 上传 Excel 2（航段数据）", type=["xlsx", "xlsm"])

if excel1_file and excel2_file:
    # 读取 Excel 2，并自动清理无效列
    try:
        raw_df = pd.read_excel(excel2_file, header=0)
    except Exception as e:
        st.error(f"读取 Excel 2 失败：{e}")
        st.stop()

    # 删除所有列名包含 'Unnamed' 的列（不区分大小写）
    cols_to_drop = [c for c in raw_df.columns if 'unnamed' in c.lower()]
    if cols_to_drop:
        raw_df.drop(columns=cols_to_drop, inplace=True)
        st.info(f"已自动删除无效列：{', '.join(cols_to_drop)}")

    # 如果还有重复列名，保留第一个
    raw_df = raw_df.loc[:, ~raw_df.columns.duplicated()]

    # 若列数为0，报错
    if raw_df.empty or len(raw_df.columns) == 0:
        st.error("Excel 2 没有有效的列，请检查文件是否正确（需包含表头）。")
        st.stop()

    excel2_df = raw_df
    valid_cols = excel2_df.columns.tolist()

    st.subheader("⚙️ 请指定 Excel 2 中各列的含义（务必仔细核对）")
    # 智能建议默认列名
    def suggest(cols, keywords):
        for kw in keywords:
            for c in cols:
                if kw in c:
                    return c
        return cols[0]

    default_flight = suggest(valid_cols, ["实际飞行时间", "飞行时间", "航段时间"])
    default_dep = suggest(valid_cols, ["出发城市", "起飞机场", "出发地"])
    default_arr = suggest(valid_cols, ["到达城市", "目的地机场", "到达地"])
    default_reg = suggest(valid_cols, ["飞机注册号", "注册号", "机号"])

    flight_col = st.selectbox("🕒 飞行时间列（格式如 09:01）", valid_cols, index=valid_cols.index(default_flight))
    dep_col = st.selectbox("🏙️ 出发城市列", valid_cols, index=valid_cols.index(default_dep))
    arr_col = st.selectbox("🏙️ 到达城市列", valid_cols, index=valid_cols.index(default_arr))
    reg_col = st.selectbox("✈️ 飞机注册号列（用于去重统计）", valid_cols, index=valid_cols.index(default_reg))

    # 预览数据
    st.subheader("📊 数据预览（所选列的前5行）")
    try:
        preview_df = excel2_df[[flight_col, dep_col, arr_col, reg_col]].head(5)
        st.dataframe(preview_df)
    except Exception as e:
        st.error(f"预览失败，请检查所选列是否正确。错误：{e}")

    # 显示有效航段数
    if flight_col in excel2_df.columns:
        valid_count = excel2_df[flight_col].notna().sum()
        st.info(f"📌 飞行时间非空的航段数：{valid_count} 条")
    else:
        st.error("所选飞行时间列无效，请重新选择。")

    if st.button("🚀 开始处理", type="primary"):
        if flight_col not in excel2_df.columns:
            st.error("飞行时间列无效，请重新选择。")
            st.stop()

        with st.spinner("正在处理，请稍候..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(excel1_file.getvalue())
                    excel1_path = tmp1.name

                output_path, stats = update_excel1(
                    excel1_path,
                    excel2_df,
                    flight_col,
                    dep_col,
                    arr_col,
                    reg_col
                )

                st.success("✅ 处理完成！以下是本次更新的统计信息：")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("昨日飞行时间 (总分钟)", stats['昨日总分钟'])
                    st.metric("格式化后 (J3)", format_duration(stats['昨日总分钟']))
                with col2:
                    st.metric("架次 (K3)", stats['架次'])
                    st.metric("去重注册号数量 (O3)", stats['注册号数量'])
                with col3:
                    st.metric("旧累计分钟", stats['旧累计分钟'])
                    st.metric("新累计分钟 (L3)", stats['新累计分钟'])

                st.write("去重注册号列表：", ", ".join(stats['注册号列表']))

                with st.expander("🔍 查看每个航段的时间解析详情（用于排查）"):
                    df_time = pd.DataFrame(stats['时间解析详情'], columns=['行索引', '原始值', '解析分钟数'])
                    st.dataframe(df_time)

                with open(output_path, 'rb') as f:
                    bytes_data = f.read()
                st.download_button(
                    label="📥 下载 Excel 1（已更新）",
                    data=bytes_data,
                    file_name="updated_excel1.xlsx",
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

# ---------- 使用说明 ----------
st.markdown("---")
with st.expander("📖 使用说明"):
    st.markdown("""
    **操作步骤**
    1. 上传模板（Excel 1）和航段数据（Excel 2）。
    2. 从下拉列表中为 **飞行时间、出发城市、到达城市、飞机注册号** 分别选择正确的列。
    3. 检查下方的数据预览，确认选择无误。
    4. 点击“开始处理”，查看统计结果，并下载更新后的模板。

    **常见问题**
    - 如果 J3 仍为 0:00，请展开“时间解析详情”，检查每个航段的原始值是否被正确解析为分钟数。
    - 如果 N3 显示的内容不对，说明您把出发/到达城市选成了其他列（如客户名）。
    - 程序会自动删除 `Unnamed` 之类的无效列，所以下拉框不会出现它们。
    """)
