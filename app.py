import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (鎖定 Alison 原始規範)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：回復到最穩定的標色格式
STYLE = {
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)},
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=14, color="FFFFFF", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵：先將 NaN 轉為空字串，防止程式在比對文字時當機
        df_audit = df.fillna("")
        
        # 定位日期列 (C 欄)
        d_row_idx = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row_idx = i
                break
        if d_row_idx is None: continue

        # 核心規格審核 (專注於品名規格，避開空白判定導致的崩潰)
        for col in range(3, 8): # D 到 H 欄
            # 取得該欄日期
            date_val = str(df_audit.iloc[d_row_idx, col]).split("\n")[0]
            
            for r_idx, row in df_audit.iterrows():
                content = str(row[col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)
                
                if content == "": continue # 遇到空白直接跳過，不進行處理

                # 1. 白帶魚規格 (150g)
                if "白帶魚" in content and "150g" not in content:
                    cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                    logs.append({"日期": date_val, "缺失": "規格缺失", "內容": f"白帶魚未標 150g"})
                
                # 2. 獅子頭規格 (60gX2)
                if "獅子頭" in content and "60gX2" not in content:
                    cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                    logs.append({"日期": date_val, "缺失": "規格缺失", "內容": f"獅子頭未標 60gX2"})

                # 3. 漢堡排規格 (150g)
                if "漢堡排" in content and "150g" not in content:
                    cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                    logs.append({"日期": date_val, "缺失": "規格缺失", "內容": f"漢堡排未標 150g"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳 4 月菜單 Excel 檔案", type=["xlsx"])
if up:
    with st.spinner("系統正在執行規格審核..."):
        results, processed_data = audit_process(up)
        
    if results:
        st.error(f"🚩 審核完畢，發現 {len(results)} 項規格缺失。")
        st.table(pd.DataFrame(results))
        st.download_button(
            label="📥 下載標註完成之退件檔",
            data=processed_data,
            file_name=f"退件建議_{up.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("✅ 審核完畢，未發現規格缺失！")
