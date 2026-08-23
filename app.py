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

def auto_match_column(df, candidates):
    """根据候选关键词自动匹配列名"""
    cols_lower = {col.strip().lower(): col for col in df.columns}
    for cand in candidates:
        cand_lower = cand.strip().lower()
        if cand_lower in cols_lower:
            return cols_lower[cand_lower]
    # 模糊匹配
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

# ---------- Streamlit 界面 ----------
st.set_page_config(page_title="飞行数据自动更新", layout="centered")
st.title("🛩️ 飞行数据自动更新")

excel1_file = st.file_uploader("📂 上传昨日飞行数据（operation-每日飞行数据）", type=["xlsx", "xlsm"])
excel2_file = st.file_uploader("📂 上传航段数据导出", type=["xlsx", "xlsm"])

# 当两个文件都上传时自动处理
if excel1_file and excel2_file:
    with st.spinner("正在自动处理..."):
        try:
            # 读取航段数据，自动检测表头
            keywords = ["客户", "航班号", "出发城市", "到达城市", "实际飞行时间", "飞机注册号"]
            excel2_df = read_excel_with_auto_header(excel2_file, keywords)
            if excel2_df.empty or len(excel2_df.columns) == 0:
                st.error("航段数据没有有效列，请检查文件格式。")
                st.stop()

            # 自动匹配列（若匹配失败则提示）
            flight_col = auto_match_column(excel2_df, ["实际飞行时间", "飞行时间", "航段时间"])
            dep_col = auto_match_column(excel2_df, ["出发城市", "起飞机场", "出发地"])
            arr_col = auto_match_column(excel2_df, ["到达城市", "目的地机场", "到达地"])
            reg_col = auto_match_column(excel2_df, ["飞机注册号", "注册号", "机号"])

            # 检查是否所有列都已匹配
            missing = []
            if not flight_col:
                missing.append("飞行时间")
            if not dep_col:
                missing.append("出发城市")
            if not arr_col:
                missing.append("到达城市")
            if not reg_col:
                missing.append("飞机注册号")
            if missing:
                st.error(f"未能自动匹配以下列：{', '.join(missing)}，请检查文件列名是否包含关键词。")
                st.stop()

            # 保存模板到临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                tmp1.write(excel1_file.getvalue())
                excel1_path = tmp1.name

            # 执行更新
            output_path, stats = update_excel1(
                excel1_path, excel2_df, flight_col, dep_col, arr_col, reg_col
            )

            # 显示核心指标
            st.success("✅ 处理完成")
            col1, col2, col3 = st.columns(3)
            col1.metric("昨日总飞行时间", stats['昨日飞行时间'])
            col2.metric("架次", stats['架次'])
            col3.metric("使用航空器数量", stats['注册号数量'])

            col4, col5 = st.columns(2)
            col4.metric("截止昨日总飞行时间", stats['截止昨日总飞行时间'])
            col5.metric("截止今日总飞行时间", stats['截止今日总飞行时间'])

            # 生成下载文件名
            yesterday = datetime.now().date() - timedelta(days=1)
            file_name = f"中南-深圳局-天成商务航空有限公司-{yesterday.month}月{yesterday.day}日飞行数据.xlsx"

            with open(output_path, 'rb') as f:
                st.download_button(
                    label="📥 下载更新后的文件",
                    data=f.read(),
                    file_name=file_name,
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

# 底部小提示
st.caption("💡 上传两个文件后自动处理，无需点击任何按钮。")
