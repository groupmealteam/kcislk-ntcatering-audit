import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (嚴格遵循 Alison 要求，標題與註解絕不更動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (專殺 4/28-4/29 的空白漏洞)
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
        # 核心修正：將所有 NaN 強制標註為 "EMPTY"，不准程式裝瞎
        df_audit = df.fillna("EMPTY")
        
        # 尋找日期列 (定錨)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        # 掃描週一到週五 (D-H 欄)
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列往下，地毯式掃描標籤
            for r_idx in range(d_row + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 抓包 A：熱量黑洞 (針對 4/28, 4/29 晚餐熱量)
                if "熱量" in label:
                    if content in ["EMPTY", "", "0", "nan"]:
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量未填！"})

                # 抓包 B：幽靈菜單 (針對 4/29 副菜：菜名空白但下格明細有字)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    try:
                        detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                        if content == "EMPTY" and detail_val != "EMPTY":
                            cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                            logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名！"})
                    except: pass

                # 抓包 C：規格缺失 (白帶魚/獅子頭/漢堡排)
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 (Alison 規範格式) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳 Excel 檔案進行最後審核", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項重大缺失（含 4/28-4/29 空格）。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
