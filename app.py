import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (嚴格遵循 Alison 指示，標題與註解絕不更動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式規範：黑底白字 30 級 (專殺 4/28-4/29 的空白漏洞)
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
        # 強制將所有 NaN 轉為 "EMPTY"，不准程式裝瞎跳過
        df_audit = df.fillna("EMPTY")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D 到 H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 偵測 A：熱量黑洞 (針對 4/28, 4/29 晚餐熱量空白)
                if "熱量" in label and content in ["EMPTY", "", "0", "nan"]:
                    cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                    logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量欄位被挖空！"})

                # 偵測 B：幽靈副菜 (針對 4/29 有明細無菜名)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    if content == "EMPTY":
                        # 檢查：如果這一格是空白，但下一行(食材明細)卻有文字
                        try:
                            next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_val != "EMPTY":
                                cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                                logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名(只有明細)"})
                        except: pass

                # 偵測 C：合約文字遊戲 (針對白帶魚、獅子頭)
                check_specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, spec in check_specs.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 需標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 (Alison 規範格式) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案進行最後審核", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項缺失（含紅框處）。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
