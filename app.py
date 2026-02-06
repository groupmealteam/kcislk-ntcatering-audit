import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 註解：製作者 Alison (針對 4/28-4/29 缺失校正) ---
STYLE = {
    "BLACK_CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)}, # 黑底白字：針對妳抓到的「空白」
    "RED_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF")},             # 紅底白字
    "YELLOW_CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)} # 黃底紅字
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        df_audit = df.fillna("MISSING") 
        
        # 定位日期 Row (C 欄)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D 到 H 欄
            date_val = str(df_audit.iloc[d_row, col]).strip()
            
            # --- 核心糾錯 1：針對 4/28, 4/29 熱量空白 ---
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                
                if "熱量" in label:
                    if content in ["MISSING", "", "0", "nan"]:
                        ws.cell(row=r_idx+1, column=col+1).fill = STYLE["BLACK_CRITICAL"]["fill"]
                        ws.cell(row=r_idx+1, column=col+1).font = STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "項目": "數據缺失", "原因": "⚠️ 熱量欄位不可空白！"})

                # --- 核心糾錯 2：針對 4/29 副菜「有明細無菜名」 ---
                # 邏輯：檢查主菜/副菜標籤格，若為空但其下方一格(食材明細)有字，即為嚴重缺失
                target_tags = ["主菜", "副菜", "青菜", "湯品"]
                if any(t == label for t in target_tags):
                    detail_content = str(df_audit.iloc[r_idx+1, col]).strip()
                    if content == "MISSING" and detail_content != "MISSING":
                        ws.cell(row=r_idx+1, column=col+1).fill = STYLE["BLACK_CRITICAL"]["fill"]
                        ws.cell(row=r_idx+1, column=col+1).font = STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "項目": "結構缺失", "原因": f"❌ {label} 有明細卻無菜名！"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison (已鎖定 4/29 幽靈菜名漏洞)")

up = st.file_uploader("📂 請上傳那份 4/28-4/30 的 Excel 檔案", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項重大違規。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔 (去看 4/29 的黑洞)", data, f"退件建議_{up.name}")
