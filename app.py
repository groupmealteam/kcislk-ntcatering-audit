import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (完全不動 Alison 的原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：黑底白字 30 級
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 把所有空值填成一個字串，這樣程式才「看得到」它
        df_audit = df.fillna("MISSING_DATA")
        
        # 定位日期 (C 欄)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 1. 抓熱量空值 (針對 4/28, 4/29 晚餐)
                if "熱量" in label and content == "MISSING_DATA":
                    cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                    logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量未填！"})

                # 2. 抓副菜空值 (針對 4/29 有明細無菜名)
                if label in ["主菜", "副菜", "青菜", "湯品"] and content == "MISSING_DATA":
                    # 往下看一格，如果有食材明細，這格就必須噴黑
                    try:
                        next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                        if next_val != "MISSING_DATA":
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名！"})
                    except: pass

                # 3. 抓規格 (白帶魚 150g / 獅子頭 60gX2)
                if "白帶魚" in content and "150g" not in content:
                    cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                    logs.append({"日期": date_val, "缺失": "規格缺失", "原因": "白帶魚需標註 150g"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載結果", data, f"退件_{up.name}")
