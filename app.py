import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (維持 Alison 原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (專抓紅框空白) / 黃底紅字 (殺規格)
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    # 關鍵修正：將所有 Sheet 讀入後，強制將 NaN 填補為 "MISSING" 字串
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        df_audit = df.fillna("MISSING")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 突破 BUG：強制查核模式 ---
                # 只要標籤包含這些關鍵字，內容絕對不能是 MISSING
                mandatory_tags = ["熱量", "主菜", "副菜", "套餐", "主食"]
                if any(tag in label for tag in mandatory_tags):
                    if content == "MISSING" or content == "":
                        # 檢查 4/29 特殊漏填：菜名空，但下一行(食材)有字，必抓
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "MISSING" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "原因": f"❌ {label} 欄位未填！"})
                        except: pass

                # --- 規格審核：原有的穩定功能 ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單檔案 (最後測試：4/28-4/29 黑洞)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 成功抓到 {len(results)} 項缺失（含紅框處空白）。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
