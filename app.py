import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 標題與規範
ST_TITLE = "🛡️ 團膳區(新北食品) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"

# 樣式：黑底白字 (專門退件完全不填的廠商)
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}

def alison_pro_audit(file):
    fname = file.name
    # 1. 檔名第一道防線
    if any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode = "新北食品-教育學部"
        nutri_indices = [9, 10, 11, 12, 13, 14, 15] # J-P 欄
    elif any(kw in fname for kw in ["美食街", "素食"]):
        mode = "新北食品-美食街/素食"
        nutri_indices = [3, 4, 5, 6, 7]
    else:
        return None, "BLOCK", None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        real_content_flag = False  # --- 【新增】防詐騙旗標 ---

        for sn, df in sheets_df.items():
            ws = wb[sn]
            # 數據清洗：確保 '0' 不會被轉成空字串
            df_audit = df.astype(str).replace(['nan', 'NaN', 'None', 'NoneType'], '')
            
            for r_idx in range(len(df_audit)):
                # 抓取第一欄標籤 (日期標籤)
                label = str(df_audit.iloc[r_idx, 0]).strip()
                
                # 判定這行是否為有效日期行 (例如: 3/27 (五))
                if "/" in label and "(" in label:
                    real_content_flag = True  # 只要抓到一行日期，就代表這檔案是真的菜單
                    
                    # 檢查主食欄(第1欄)是否有內容
                    main_food = str(df_audit.iloc[r_idx, 1]).strip()
                    if main_food != "":
                        for n_idx in nutri_indices:
                            if n_idx >= len(df_audit.columns): continue
                            
                            val = str(df_audit.iloc[r_idx, n_idx]).strip()
                            
                            # 【核心邏輯】：只有「真空」才算錯，'0' 是數據，不准報錯
                            if val == "":
                                cell = ws.cell(row=r_idx+1, column=n_idx+1)
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填數據"
                                logs.append({"分頁": sn, "日期": label, "缺失": f"營養欄位{n_idx+1}真空漏填"})

        # --- 【核心防詐修正】：如果檔名對但內容找不到任何日期標籤 ---
        if not real_content_flag:
            return None, "INVALID_CONTENT", None

        output = BytesIO()
        wb.save(output)
        return logs, mode, output.getvalue()
    except Exception as e:
        return None, f"ERROR: {str(e)}", None

# --- Streamlit UI 介面 ---
st.set_page_config(page_title="新北食品稽核系統", layout="wide")
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

up = st.file_uploader("📂 請上傳菜單檔案 (xlsx)", type=["xlsx"])

if up:
    with st.spinner("Alison 正在嚴格校對中..."):
        logs, m, data = alison_pro_audit(up)
        
        if m == "BLOCK":
            st.error("❌ 檔名識別錯誤！請確認包含『小學』、『幼兒園』、『美食街』或『素食』。")
        elif m == "INVALID_CONTENT":
            st.error("❌ 內容識別失敗！雖然檔名正確，但內容偵測不到菜單格式（日期標籤），請確認檔案內容。")
        elif "ERROR" in m:
            st.error(f"❌ 程式崩潰：{m}")
        else:
            st.success(f"✅ 已識別模式：{m}")
            if logs:
                st.warning(f"🚩 偵測到 {len(logs)} 處『真空空白』缺失（已噴黑標註）。")
                st.table(pd.DataFrame(logs))
                st.download_button("📥 下載 Alison 標註退件檔", data, f"退件_{up.name}")
            else:
                st.success("🎉 數據稽核完美！(包含 0 值數據已確認無誤)")
