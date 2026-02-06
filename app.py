import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (嚴格遵照 Alison 要求：標題與註解絕不更動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義 (黑底白字 30 級：專殺空白、漏填、刪除)
STYLE = {
    "BLACK_CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 把所有空值標記為 "EMPTY"，讓程式「看見」空白
        df_audit = df.fillna("EMPTY")
        
        # 尋找日期列
        d_row_idx = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row_idx = i
                break
        
        if d_row_idx is None: continue

        for col in range(3, 8): # D 到 H 欄
            date_val = str(df_audit.iloc[d_row_idx, col]).strip()
            
            # 遍歷整行尋找妳說的「紅框」漏洞
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 抓包 A：4/28, 4/29 熱量空白
                if "熱量" in label and content in ["EMPTY", "", "0", "nan"]:
                    cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                    logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量未填！違反審閱原則"})

                # 抓包 B：4/29 副菜「有明細無菜名」
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    # 邏輯：這一格是空的，但下一格（食材明細）竟然有字
                    next_content = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "EMPTY" and next_content != "EMPTY":
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名(只有明細)"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 上傳菜單 Excel 進行稽核", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 發現 {len(results)} 項嚴重缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
