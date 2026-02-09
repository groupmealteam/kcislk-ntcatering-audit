import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 標題完全遵照 Alison 規範，一字不差
ST_TITLE = "🛡️ 團膳稽核系統 - 小學部 / 幼兒園 (細項模式)"
ST_AUTHOR = "製作者：Alison"

# 樣式：黑底白字
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}

def alison_audit_core(file):
    fname = file.name
    
    # 嚴格過濾：沒關鍵字不准審 (修正妳說的 BUG)
    if any(kw in fname for kw in ["小學", "幼兒園"]):
        mode = "教育學部"
        nutri_cols = [9, 10, 11, 12, 13, 14, 15] # J-P 欄
    elif "美食街" in fname:
        mode = "美食街"
        nutri_cols = [3, 4, 5, 6, 7]
    else:
        return None, "INVALID_FILE", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        df_audit = df.astype(str).replace(['nan', 'NaN', 'None', '0.0', '0'], '')
        
        for r_idx in range(len(df_audit)):
            label = df_audit.iloc[r_idx, 0]
            
            # 只有日期行才審核營養分析 (解決裝瞎問題)
            if "/" in label and "(" in label:
                # 檢查營養成分分析是否為空 (妳最在意的點)
                for c_idx in nutri_cols:
                    if c_idx < len(df_audit.columns):
                        val = df_audit.iloc[r_idx, c_idx].strip()
                        if val == "":
                            cell = ws.cell(row=r_idx+1, column=c_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌數據缺失"
                            logs.append({"分頁": sn, "日期": label, "缺失": f"第{c_idx+1}欄營養數據空白"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# Streamlit UI
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

up = st.file_uploader("📂 請上傳菜單檔案", type=["xlsx"])
if up:
    logs, mode, data = alison_audit_core(up)
    if mode == "INVALID_FILE":
        st.error("❌ 檔名不符規範，系統拒絕審核。")
    else:
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項缺失，包含營養分析空白。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件檔", data, f"退件_{up.name}")
        else:
            st.success("✅ 數據完整！")
