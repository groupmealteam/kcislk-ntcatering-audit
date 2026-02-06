import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (嚴格遵循 Alison 指示，標題與註解絕不更動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：黑底白字 30 級 (專殺 4/28-4/29 的空白漏洞)
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
        # 關鍵修正：將所有 NaN 強制轉為字串 "VOID"，不准程式裝瞎跳過
        df_audit = df.fillna("VOID")
        
        # 搜尋日期 Row (C 欄位)
        d_row_idx = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row_idx = i
                break
        if d_row_idx is None: continue

        # 掃描週一到週五 (D-H 欄)
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row_idx, col]).split("\n")[0]
            
            # 遍歷該欄位的所有儲存格進行「標籤 vs 內容」強制檢查
            for r_idx in range(d_row_idx + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 偵測 A：熱量黑洞 (4/28, 4/29 晚餐熱量空白)
                if "熱量" in label:
                    if content in ["VOID", "", "nan", "0"]:
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量未填！(違反審閱原則)"})

                # 偵測 B：幽靈副菜 (4/29 菜名空白但有食材)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    # 邏輯：這一格是 VOID，但下一格(食材明細)卻不是 VOID
                    next_row_content = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "VOID" and next_row_content != "VOID":
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名！"})

                # 偵測 C：合約文字遊戲 (白帶魚 150g / 獅子頭 60gX2)
                check_specs = {"白帶魚": "150g", "獅子頭": "60gX2"}
                for item, spec in check_specs.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item}需標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 (標題不准改) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項重大缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
