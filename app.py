import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (保持 Alison 原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 定義：黑底白字 30 級 (專殺空白) / 黃底紅字 (殺規格)
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制將所有空值變為 "VOID"，讓程式看得到黑洞
        df_audit = df.fillna("VOID")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列往下每一格都要過濾
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心邏輯：強制偵測不完整 ---
                # 只要標籤包含這些關鍵字，右邊如果是 VOID，直接噴黑
                critical_labels = ["熱量", "主菜", "副菜", "湯品"]
                if any(tag in label for tag in critical_labels):
                    if content == "VOID":
                        # 特別針對 4/29：如果菜名空，但下一格食材有字，這必殺
                        try:
                            next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_val != "VOID" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "缺失": "不完整", "原因": f"❌ {label} 沒寫內容！"})
                        except: pass

                # --- 核心邏輯：規格審核 ---
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

up = st.file_uploader("📂 上傳 Excel 測試 4/28-4/29 紅框缺失", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 成功抓到 {len(results)} 項缺失！包含紅框空白與規格不符。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
