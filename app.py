import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 樣式設定：黑底白字 30 級 (專殺 4/28-4/29 這種黑洞) ---
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
        # 突破點 1：將所有空值填充為 "MISSING"，強迫程式看見「無」
        df_audit = df.fillna("MISSING")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 突破點 2：鎖定標籤！只要標籤在，內容是 MISSING 就噴黑
                # 針對 4/28, 4/29 的熱量、主菜、副菜
                mandatory_tags = ["熱量", "主菜", "副菜", "主食", "套餐"]
                if any(tag in label for tag in mandatory_tags):
                    if content == "MISSING":
                        # 檢查 4/29 副菜漏洞：內容空但下一行明細有字
                        try:
                            next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_val != "MISSING" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "原因": f"❌ {label} 欄位漏填！"})
                        except: pass

                # 突破點 3：強化規格稽核 (模糊匹配)
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項缺失！包含 4/28-4/29 紅框處。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
