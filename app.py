import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 樣式設定：黑底白字 30 級 (專殺空白) / 黃底紅字 (殺規格) ---
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 核心突破：強迫讀取所有內容為字串，並把 NaN 填補為 "MISSING"
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0'], 'MISSING')
        
        # 定位日期 Row (定錨點)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 暖禾式對位檢查：標籤驅動 ---
                # 只要標籤包含關鍵字，右邊如果是 MISSING，就直接噴黑
                critical_labels = ["熱量", "主菜", "副菜", "主食", "套餐"]
                if any(tag in label for tag in critical_labels):
                    if content == "MISSING" or content == "":
                        # 特別針對 4/29：檢查下一行是否有食材明細
                        is_blank_fail = False
                        if "熱量" in label:
                            is_blank_fail = True
                        else:
                            try:
                                next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                                if next_val != "MISSING": is_blank_fail = True
                            except: pass
                        
                        if is_blank_fail:
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"日期": date_val, "原因": f"❌ {label} 漏填內容！"})

                # --- 規格稽核 (模糊匹配) ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 上傳 Excel (暖禾邏輯加強版)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項缺失（含 4/28-4/29 黑洞）。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完整，這次廠商沒逃掉！")
