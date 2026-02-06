import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題與註解完全鎖死，不准更動)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
STYLE = {
    "BLACK_ALERT": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 修正 BUG 1：強制把所有 NaN 轉為字串 "MISSING"，讓它無所遁形
        df_audit = df.fillna("MISSING")
        
        # 尋找日期列 (定錨)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D 到 H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 修正 BUG 2：改用標籤掃描制，而不是位置對齊制
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 抓包：4/28, 4/29 熱量空白
                if "熱量" in label and content in ["MISSING", "", "0"]:
                    cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                    logs.append({"日期": date_val, "項目": "數據缺失", "原因": "⚠️ 熱量欄位被挖空！"})

                # 抓包：4/29 副菜有明細無菜名
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "MISSING" and next_val != "MISSING":
                        cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                        logs.append({"日期": date_val, "項目": "結構缺失", "原因": f"❌ {label} 只有明細，菜名空白！"})

                # 抓包：規格缺失 (白帶魚/獅子頭)
                specs = {"白帶魚": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "項目": "規格不符", "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 上傳 0428-0430 檔案測試", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項重大缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載標註檔", data, f"退件建議_{up.name}")
