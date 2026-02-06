import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題依照要求固定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式定義：這次強化了「異常缺失」的視覺
STYLE = {
    "CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}, # 黑底白字：針對刪除熱量、少菜
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF")},      # 紅底白字：數據違規
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)} # 黃底紅字：合約規格
}

# 規格對標 (依據增補協議書)
MUST_CHECK = {"獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", "白蝦": "X3", "砂鍋魚丁": "250g"}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        df = df.fillna("MISSING_DATA") # 強制把空值標註出來，不讓它逃過稽核
        ws = wb[sn]
        
        # 識別學部熱量標準 (依修訂2)
        std = None
        if "幼兒園" in sn: std = {"熱量": (350, 480), "蛋白質": 2.0}
        elif "小學" in sn: std = {"熱量": (650, 780), "蛋白質": 3.0}
        elif "美食街" in sn: std = {"熱量": (750, 850), "蛋白質": 4.0}
        if not std: continue

        # 定位日期 (C 欄「日期Date」)
        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8):
            date_val = str(df.iloc[d_row, col]).split(" ")[0]

            # --- 1. 結構完整性抓包 (原則一：少菜必噴黑底) ---
            # 強制掃描主食到湯品共 5 行
            for offset in range(2, 7):
                r_idx = d_row + offset
                val = str(df.iloc[r_idx, col]).strip()
                if val == "MISSING_DATA" or val == "":
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    cell.fill, cell.font = STYLE["CRITICAL"]["fill"], STYLE["CRITICAL"]["font"]
                    logs.append({"日期": date_val, "項目": "結構缺失", "原因": "⚠️ 菜名空白！違反原則一"})

            # --- 2. 營養數據抓包 (針對妳說的「熱量刪掉」) ---
            for r_idx in range(len(df)):
                label = str(df.iloc[r_idx, 2])
                if any(x in label for x in ["熱量", "蛋白質", "豆魚"]):
                    val_raw = str(df.iloc[r_idx, col]).strip()
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    if val_raw == "MISSING_DATA" or val_raw == "0" or val_raw == "0.0":
                        cell.fill, cell.font = STYLE["CRITICAL"]["fill"], STYLE["CRITICAL"]["font"]
                        logs.append({"日期": date_val, "項目": "數據缺失", "原因": f"❌ {label} 標示不可缺失！"})
                    else:
                        # 既有的數值判斷邏輯...
                        pass

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
# ... (UI 略)
