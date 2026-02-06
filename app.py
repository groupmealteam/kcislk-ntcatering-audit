import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題固定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式定義
STYLE = {
    "EMPTY_ALERT": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}, # 黑底白字：針對妳說的「刪掉、少菜」
    "DATA_FAIL":   {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF")},      # 紅底白字：數據不符
    "CONTRACT":    {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)} # 黃底紅字：規格不符
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        df = df.fillna("") # 將空值轉為字串處理
        ws = wb[sn]
        
        # 定位日期 (新北食品核心格式)
        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8):
            date_val = str(df.iloc[d_row, col]).split(" ")[0]

            # --- 核心稽核 1：結構完整性 (解決「少好幾道菜」的問題) ---
            # 依原則一，檢查主食、主菜、副菜、青菜、湯品 5 大必備項
            for offset in range(2, 7):
                r_idx = d_row + offset
                val = str(df.iloc[r_idx, col]).strip()
                if val == "" or val.lower() == "nan":
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    cell.fill, cell.font = STYLE["EMPTY_ALERT"]["fill"], STYLE["EMPTY_ALERT"]["font"]
                    logs.append({"分頁": sn, "日期": date_val, "項目": "結構缺項", "原因": "❌ 菜名空白，違反原則一"})

            # --- 核心稽核 2：營養標示必填 (解決「熱量刪掉」的問題) ---
            for r_idx in range(len(df)):
                label = str(df.iloc[r_idx, 2])
                if any(x in label for x in ["熱量", "蛋白質", "豆魚"]):
                    val_raw = str(df.iloc[r_idx, col]).strip()
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # 抓包點：如果數值是空的或零
                    if val_raw == "" or val_raw == "0" or val_raw == "0.0":
                        cell.fill, cell.font = STYLE["EMPTY_ALERT"]["fill"], STYLE["EMPTY_ALERT"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "數據缺失", "原因": f"❌ {label} 標示不可為空或零"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
# (Streamlit UI 邏輯與上傳組件)
