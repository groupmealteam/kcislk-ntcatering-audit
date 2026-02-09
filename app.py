import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 標題定義：嚴格遵守 Alison 規範 ---
ST_TITLE = "🛡️ 團膳區(新北食品) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"

# 樣式：黑底白字 (專門對付漏填)
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}

def alison_smart_audit_v2(file):
    fname = file.name
    # 模式鎖定 BUG 修正
    if any(kw in fname for kw in ["小學", "幼兒園"]):
        mode = "教育學部"
        # 營養分析固定欄位 (J-P 欄)
        nutri_indices = [9, 10, 11, 12, 13, 14, 15]
    elif "美食街" in fname:
        mode = "美食街"
        nutri_indices = [3, 4, 5, 6, 7]
    else:
        return None, "BLOCK", None # 非指定關鍵字直接阻斷

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 清理偽空值
        df_audit = df.astype(str).replace(['nan', 'NaN', 'None', '0.0', '0'], '')
        
        for r_idx in range(len(df_audit)):
            label = str(df_audit.iloc[r_idx, 0]).strip()
            
            # 偵測日期列
            if "/" in label and "(" in label:
                # 聰明檢查：如果當天有主食，營養分析就不能空
                has_lunch = df_audit.iloc[r_idx, 1] != ""
                if has_lunch:
                    for n_idx in nutri_indices:
                        val = df_audit.iloc[r_idx, n_idx].strip() if n_idx < len(df_audit.columns) else ""
                        # 只要是空值或非數字，直接噴黑 (解決 4/29 數據偏移導致的真空問題)
                        is_numeric = val.replace('.', '', 1).isdigit()
                        if val == "" or not is_numeric:
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌數據缺失"
                            logs.append({"分頁": sn, "日期": label, "缺失": "營養分析空白或格式錯誤"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- Streamlit 啟動區 ---
st.set_page_config(page_title="Alison 稽核系統", layout="wide")
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

up_file = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if up_file:
    logs, m_detected, out_data = alison_smart_audit_v2(up_file)
    if m_detected == "BLOCK":
        st.error("❌ 檔名未包含指定關鍵字，系統拒絕審核。")
    else:
        if logs:
            st.warning(f"🚩 偵測到 {len(logs)} 處缺失（包含營養分析空白）。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件檔", out_data, f"退件_{up_file.name}")
        else:
            st.success("✅ 數據審核通過！")
