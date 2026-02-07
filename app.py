import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 樣式設定：黑底白字 30 級 (專殺 4/28-4/29 這種空洞) ---
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 突破點 1：強迫程式看見「無」，將 NaN 變成 "EMPTY_VOID"
        df_audit = df.fillna("EMPTY_VOID")
        
        # 定位日期標籤 (定錨點)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # 掃描週一到週五 (D-H 欄)
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 突破點 2：強制標籤連動檢查 (針對熱量、菜名)
                target_tags = ["熱量", "主菜", "副菜", "主食", "套餐"]
                if any(tag in label for tag in target_tags):
                    # 如果內容是空的 (或我們先前標記的 EMPTY_VOID)
                    if content in ["EMPTY_VOID", "", "nan", "0"]:
                        # 檢查 4/29 漏洞：菜名空，但下一行有食材明細
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            # 只要是熱量欄位，或是「漏填菜名但有食材」的情況，直接噴黑
                            if "熱量" in label or detail_val != "EMPTY_VOID":
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "缺失": "內容不全", "原因": f"❌ {label} 沒填寫！"})
                        except: pass

                # 突破點 3：強化規格稽核 (白帶魚、漢堡排等)
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單檔案 (最後測試：4/28-4/29 黑洞)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項嚴重缺失，已自動噴黑/噴黃標註。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
    else:
        st.success("✅ 結構完整，且未發現規格缺失！")
