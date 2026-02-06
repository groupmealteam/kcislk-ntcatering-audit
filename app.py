import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (保持 Alison 的原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (專殺 4/28-4/29 的空白) / 黃底紅字 (殺規格)
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 修正核心：將所有空值強迫轉為 "MISSING"，讓它變為實體內容
        df_audit = df.fillna("MISSING")
        
        # 尋找日期列 (定錨點)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心邏輯 A：強制空白偵測 (專抓 4/28, 4/29) ---
                target_tags = ["熱量", "主菜", "副菜", "湯品"]
                if any(tag in label for tag in target_tags):
                    if content == "MISSING":
                        # 特別針對 4/29：菜名是空的，但下一格食材不是空的，這必殺
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "MISSING" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "缺失": "不完整", "原因": f"❌ {label} 沒填寫！"})
                        except: pass

                # --- 核心邏輯 B：原有規格審核 ---
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格缺失", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳 Excel，看我這次還敢不敢裝瞎", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 成功抓到 {len(results)} 項缺失！包含紅框空白。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載標註檔案", data, f"退件_{up.name}")
