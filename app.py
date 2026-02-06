import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題依照 Alison 要求固定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 視覺規範：黑底白字代表「重大缺失/漏填」 ---
STYLE = {
    "BLACK_ALERT": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "RED_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF")},
    "YELLOW_CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制將所有 NaN 轉為字串 "EMPTY_CELL"
        df_audit = df.fillna("EMPTY_CELL")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D 到 H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]

            # --- 核心稽核 A：針對 4/28, 4/29 熱量空白 ---
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                
                if "熱量" in label:
                    if content in ["EMPTY_CELL", "", "0", "nan"]:
                        cell = ws.cell(row=r_idx+1, column=col+1)
                        cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                        logs.append({"日期": date_val, "項目": "數據缺失", "原因": "⚠️ 熱量未填！違反審閱原則"})

                # --- 核心稽核 B：針對 4/29 副菜「有明細無菜名」 ---
                # 判斷：標籤是主菜/副菜，若內容為空，但其下方一格(食材明細)有內容
                target_tags = ["主菜", "副菜", "青菜", "湯品"]
                if any(t == label for t in target_tags):
                    detail_content = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "EMPTY_CELL" and detail_content != "EMPTY_CELL":
                        cell = ws.cell(row=r_idx+1, column=col+1)
                        cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                        logs.append({"日期": date_val, "項目": "結構缺失", "原因": f"❌ {label} 漏填菜名(只有明細)！"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳 4/28-4/30 測試檔案", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項違規。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔 (檢視黑洞處)", data, f"退件_{up.name}")
