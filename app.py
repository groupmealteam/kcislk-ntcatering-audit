import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：黑底白字 30 級 (專殺 4/28-4/29 的空白)
STYLE = {
    "BLACK_ERROR": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_SPEC": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵修正 1：強迫程式看見空白，將所有 NaN 填補為 "VOID_CELL"
        df_audit = df.fillna("VOID_CELL")
        
        # 定位日期 Row (C 欄位)
        d_row = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row = i
                break
        if d_row is None: continue

        # 掃描 D-H 欄
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(d_row + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 偵測 A：標籤存在但內容空白 (紅框缺失) ---
                # 只要左邊標籤有這些字，右邊如果是 VOID_CELL 或是空的，就噴黑漆
                mandatory_tags = ["熱量", "套餐", "主菜", "副菜", "湯品"]
                
                if any(t in label for t in mandatory_tags):
                    # 邏輯：如果是熱量標籤，右邊絕對不能空
                    if "熱量" in label and content in ["VOID_CELL", "", "0"]:
                        cell.fill, cell.font = STYLE["BLACK_ERROR"]["fill"], STYLE["BLACK_ERROR"]["font"]
                        logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量未填！"})
                    
                    # 邏輯：如果是菜名標籤(主/副菜)，右邊空但「下一行」有食材，這必抓！
                    elif content == "VOID_CELL":
                        try:
                            next_row_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_row_val != "VOID_CELL":
                                cell.fill, cell.font = STYLE["BLACK_ERROR"]["fill"], STYLE["BLACK_ERROR"]["font"]
                                logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名！"})
                        except: pass

                # --- 偵測 B：原本的規格審核 ---
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳 Excel 進行「空白偵測」壓力測試", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 發現 {len(results)} 項缺失（包含 4/28-4/29 的空白黑洞）")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
