import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 視覺規範
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}

def alison_pro_audit(file):
    fname = file.name
    if any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode = "新北食品-教育學部"
        nutri_indices = [9, 10, 11, 12, 13, 14, 15] 
    elif any(kw in fname for kw in ["美食街", "素食"]):
        mode = "新北食品-美食街/素食"
        nutri_indices = [3, 4, 5, 6, 7]
    else:
        return None, "BLOCK", None, 0

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        total_data_points = 0 # --- 讓妳看見檢核確實度的計數器 ---

        for sn, df in sheets_df.items():
            ws = wb[sn]
            # 保留原始 0 的數據清洗
            df_audit = df.astype(str).replace(['nan', 'NaN', 'None'], '')
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 0]).strip()
                
                # --- 核心修正：放寬日期識別，確保不會漏掃 ---
                # 只要包含 "/" 或 "202" 或 "月" 且長度適中，就視為日期行
                if ("/" in label or "202" in label) and len(label) < 15:
                    
                    # 檢查該列指定的營養欄位
                    for n_idx in nutri_indices:
                        if n_idx >= len(df_audit.columns): continue
                        
                        val = str(df_audit.iloc[r_idx, n_idx]).strip()
                        total_data_points += 1 # 確實掃描到一個數據點
                        
                        # 只有真正「什麼都沒寫」才噴黑
                        if val == "":
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌漏填數據"
                            logs.append({"分頁": sn, "日期": label, "缺失": f"欄位 {n_idx+1} 真空"})

        if total_data_points == 0:
            return None, "INVALID_CONTENT", None, 0

        output = BytesIO()
        wb.save(output)
        return logs, mode, output.getvalue(), total_data_points
    except Exception as e:
        return None, f"ERROR: {str(e)}", None, 0

# --- Streamlit UI ---
st.set_page_config(page_title="新北食品稽核系統", layout="wide")
st.title("🛡️ 團膳區(新北食品) 菜單自主稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單檔案", type=["xlsx"])
if up:
    logs, m, data, count = alison_pro_audit(up)
    
    if m == "BLOCK":
        st.error("❌ 檔名不符關鍵字。")
    elif m == "INVALID_CONTENT":
        st.error("❌ 內容格式不符！程式在檔案中找不到任何有效的日期與數據標籤。")
    else:
        st.info(f"📊 檢核確實度報告：本次共深入掃描了 **{count}** 個營養數據欄位。")
        if logs:
            st.error(f"🚩 偵測到 {len(logs)} 處『真空漏填』缺失。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件檔", data, f"退件_{up.name}")
        else:
            st.success("🎉 數據稽核確實！所有偵測到的欄位皆包含有效數值（含 0 值）。")
