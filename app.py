import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 樣式設定：缺失處一律黑底白字
STYLE_ERR = {
    "fill": PatternFill("solid", fgColor="000000"), 
    "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)
}

def alison_smart_audit(file):
    fname = file.name
    
    # --- 第一階段：嚴格身分判讀 (解決 BUG：沒關鍵字不准審) ---
    mode = None
    if any(kw in fname for kw in ["小學", "幼兒園"]):
        mode, label_idx, data_indices, nutri_indices = "教育學部", 0, [1, 2, 3, 4, 5, 6, 7], [9, 10, 11, 12, 13, 14, 15]
    elif "美食街" in fname:
        mode, label_idx, data_indices, nutri_indices = "美食街", 2, [3, 4, 5, 6, 7], [3, 4, 5, 6, 7]
    elif "輕食" in fname:
        mode, label_idx, data_indices, nutri_indices = "輕食專區", 0, [1, 2], [5, 6, 7, 8, 9, 10, 11]
    
    if mode is None:
        return None, "BLOCK", None

    # 讀取 Excel
    wb = load_workbook(file)
    # 使用 pandas 輔助掃描數據邏輯
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 清理空格
        df_audit = df.astype(str).applymap(lambda x: "" if str(x).strip().lower() in ['nan', 'none', '0', '0.0', ''] else str(x).strip())
        
        for r_idx in range(len(df_audit)):
            cell_first = df_audit.iloc[r_idx, 0]
            
            # 判斷是否為日期行 (審核起點)
            if "/" in cell_first and "(" in cell_first:
                
                # --- A. 營養成分分析全檢 (解決妳抓到的空白問題) ---
                has_lunch = df_audit.iloc[r_idx, 1] != "" # 主食有填就要審
                if has_lunch:
                    for n_idx in nutri_indices:
                        if n_idx >= df_audit.shape[1]: continue
                        val = df_audit.iloc[r_idx, n_idx]
                        # 檢查是否為純數字
                        is_numeric = val.replace('.','',1).isdigit()
                        if val == "" or not is_numeric:
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌數據缺失"
                            logs.append({"分頁": sn, "日期": cell_first, "缺失": f"營養數據異常(欄{n_idx+1})"})

                # --- B. 垂直菜名黑洞檢查 (4/29 專用) ---
                for c_idx in data_indices:
                    if c_idx >= df_audit.shape[1]: continue
                    content = df_audit.iloc[r_idx, c_idx]
                    if content == "":
                        try:
                            detail = df_audit.iloc[r_idx+1, c_idx]
                            if detail != "":
                                cell = ws.cell(row=r_idx+1, column=c_idx+1)
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填菜名"
                                logs.append({"分頁": sn, "日期": cell_first, "缺失": "有食材無菜名"})
                        except: pass

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- 3. Streamlit 介面啟動邏輯 (這段沒寫就會打不開) ---
st.set_page_config(page_title="Alison 團膳稽核系統", layout="wide")
st.title("🛡️ 團膳稽核系統 - Alison 專業嚴選版")
st.caption("製作者：Alison")

uploaded_file = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])

if uploaded_file:
    results, detected_mode, excel_data = alison_smart_audit(uploaded_file)
    
    if detected_mode == "BLOCK":
        st.error(f"❌ 無法審核：檔名『{uploaded_file.name}』不符規範，請確認是否包含「小學/美食街/輕食」關鍵字。")
    else:
        st.success(f"✅ 已啟動【{detected_mode}】稽核模式")
        if results:
            st.warning(f"🚩 發現 {len(results)} 項不完整缺失（已噴黑標註）。")
            st.table(pd.DataFrame(results))
            st.download_button(
                label="📥 下載 Alison 標註退件檔",
                data=excel_data,
                file_name=f"退件_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("🎉 完美！這份檔案營養數據與菜名均完整無缺。")
