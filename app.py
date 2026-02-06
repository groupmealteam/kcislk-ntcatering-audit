import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (嚴格遵照 Alison 要求，標題與佈局不准動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義 (黑底白字：針對刪除熱量、少菜、漏填)
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

STYLE = {
    "CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)},
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF")},
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制標記空值為 "MISSING"，防止程式裝瞎
        df_audit = df.fillna("MISSING")
        
        # 定位日期列 (C欄「日期」關鍵字)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # 檢查週一到週五
            date_val = str(df_audit.iloc[d_row, col]).strip()

            # --- 針對妳指出的 4/28-4/29 紅框現場進行精準獵殺 ---
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 1. 抓包：熱量空白 (4/28, 4/29 現場)
                if "熱量" in label and content in ["MISSING", "", "0", "nan"]:
                    cell.fill, cell.font = STYLE["CRITICAL"]["fill"], STYLE["CRITICAL"]["font"]
                    logs.append({"日期": date_val, "項目": "數據缺失", "原因": "⚠️ 熱量欄位被挖空！"})

                # 2. 抓包：副菜有明細無菜名 (4/29 現場)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    # 檢查：菜名格是空的，但下面那一格「食材明細」卻有字
                    detail_content = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "MISSING" and detail_content != "MISSING":
                        cell.fill, cell.font = STYLE["CRITICAL"]["fill"], STYLE["CRITICAL"]["font"]
                        logs.append({"日期": date_val, "項目": "結構缺失", "原因": f"❌ {label} 漏填菜名(只有明細)"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 (標題不准改) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])

if up:
    with st.spinner("稽核中..."):
        results, processed_data = audit_process(up)
        
        if results:
            st.error(f"🚩 抓到了！共發現 {len(results)} 項重大缺失（含紅框處）。")
            st.table(pd.DataFrame(results))
            st.download_button(
                label="📥 下載退件標註檔",
                data=processed_data,
                file_name=f"退件建議_{up.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("🎉 通過稽核。")
