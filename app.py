import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題與註解嚴格遵守要求)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
FONT_NAME = "微軟正黑體"
STYLE = {
    "CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF", bold=True)}, # 黑底白字：重大缺失
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF")},       # 紅底白字：數據違規
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=30, color="FF0000", bold=True)} # 黃底紅字：規格不符
}

# 規格鎖死 (依據 SE1140803 增補協議書)
CONTRACT_MAP = {"獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", "白蝦": "X3"}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 識別學部
        std = None
        if "幼兒園" in sn: std = {"熱量": (350, 480), "蛋白質": 2.0}
        elif "小學" in sn: std = {"熱量": (650, 780), "蛋白質": 3.0}
        elif "美食街" in sn: std = {"熱量": (750, 850), "蛋白質": 4.0}
        
        if not std: continue

        # 定位日期 (新北食品固定 C 欄)
        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8):
            date_val = str(df.iloc[d_row, col]).split(" ")[0]
            
            # --- 抓包點 1：菜單結構完整性 (原則一) ---
            # 檢查主食、主菜、副菜、青菜、湯品 5 格是否為空
            for r_offset in range(2, 7):
                r_idx = d_row + r_offset
                val = str(df.iloc[r_idx, col]).strip()
                if val in ["", "nan", "None"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = STYLE["CRITICAL"]["fill"]
                    logs.append({"日期": date_val, "項目": "結構缺失", "原因": "⚠️ 菜名空白 (違反原則一)"})

            # --- 抓包點 2：營養標示完整性 ---
            for r_idx in range(len(df)):
                label = str(df.iloc[r_idx, 2])
                if "熱量" in label or "蛋白質" in label or "豆魚" in label:
                    val_raw = str(df.iloc[r_idx, col]).strip()
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # 抓包：如果妳把熱量刪掉
                    if val_raw in ["", "nan", "0", "0.0"]:
                        cell.fill, cell.font = STYLE["CRITICAL"]["fill"], STYLE["CRITICAL"]["font"]
                        logs.append({"日期": date_val, "項目": "數據缺失", "原因": f"❌ {label} 被刪除或為0"})
                    else:
                        # 既有數據稽核邏輯... (省略)
                        pass

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
# UI 邏輯...
