import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 標題嚴格遵守 Alison 規範
ST_TITLE = "🛡️ 團膳區(新北食品) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"

# 樣式：黑底白字 (專治完全不填的廠商)
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}

def alison_pro_audit(file):
    fname = file.name
    # 模式判斷：包含小學、幼兒園、美食街、素食、輕食
    if any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode = "新北食品-教育學部"
        nutri_indices = [9, 10, 11, 12, 13, 14, 15] # J-P 欄
    elif "美食街" in fname or "素食" in fname:
        mode = "新北食品-美食街/素食"
        nutri_indices = [3, 4, 5, 6, 7]
    else:
        return None, "BLOCK", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 【核心修正】：只把真正的 NaN 轉為空字串，絕對保留 '0'
        df_audit = df.astype(str).replace(['nan', 'NaN', 'None'], '')
        
        for r_idx in range(len(df_audit)):
            label = str(df_audit.iloc[r_idx, 0]).strip()
            
            # 日期行稽核 (例如: 3/27 (五))
            if "/" in label and "(" in label:
                # 橫向檢查：只要主食(第1欄)有寫字，營養(J-P欄)就得有字
                if df_audit.iloc[r_idx, 1] != "":
                    for n_idx in nutri_indices:
                        val = df_audit.iloc[r_idx, n_idx].strip()
                        
                        # 【聰明判定】：只有「完全沒填」才噴黑，寫 0 或 0.1 都是合格的！
                        if val == "":
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌漏填數據"
                            logs.append({"日期": label, "缺失": f"欄位{n_idx+1}空白缺失"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# Streamlit UI 
st.set_page_config(page_title="新北食品稽核系統", layout="wide")
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

up = st.file_uploader("📂 請上傳菜單檔案", type=["xlsx"])
if up:
    logs, m, data = alison_pro_audit(up)
    if m == "BLOCK":
        st.error("❌ 檔名不符！請確認包含關鍵字（如：小學、幼兒園、美食街）。")
    else:
        st.success(f"已識別：{m}")
        if logs:
            st.error(f"🚩 發現 {len(logs)} 處『完全漏填』的缺失（已噴黑標註）。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載 Alison 標註退件檔", data, f"退件_{up.name}")
        else:
            st.success("🎉 數據非常完整，包含 0 的部分皆已確認。")
