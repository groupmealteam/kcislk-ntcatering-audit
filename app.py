import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁基本設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
FONT_NAME = "微軟正黑體"
STYLE = {
    "EMPTY": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF", bold=True)}, # 黑底白字：漏填地雷
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF")}, # 紅底：數據不符
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=30, color="FF0000", bold=True)} # 黃底：合約規格
}

# 根據《增補協議書》
CONTRACT_CHECK = {"獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", "白蝦": "X3"}
# 根據《審閱原則_修訂2》
STD_MAP = {
    "幼兒園": {"熱量": (350, 480), "蛋白質": 2.0},
    "小學":   {"熱量": (650, 780), "蛋白質": 3.0},
    "美食街": {"熱量": (750, 850), "蛋白質": 4.0}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        df = df.fillna("") # 先補空字串方便處理
        ws = wb[sn]
        current_std = next((STD_MAP[k] for k in STD_MAP if k in sn), None)
        if not current_std: continue

        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8):
            date_val = str(df.iloc[d_row, col]).split(" ")[0]
            
            # --- 1. 結構完整性檢查 (原則一：不得缺項) ---
            # 檢查主菜、副菜區 (假設 row 3-10 是菜名區)
            empty_count = 0
            for r_idx in range(d_row + 2, d_row + 8):
                txt = str(df.iloc[r_idx, col]).strip()
                if txt == "" or "None" in txt:
                    ws.cell(row=r_idx+1, column=col+1).fill = STYLE["EMPTY"]["fill"]
                    empty_count += 1
            if empty_count > 0:
                logs.append({"分頁": sn, "日期": date_val, "項目": "結構缺項", "原因": f"偵測到 {empty_count} 處菜名空白，違反原則一"})

            # --- 2. 營養標示檢查 (絕對不能刪掉！) ---
            for r_idx in range(len(df)):
                label = str(df.iloc[r_idx, 2])
                if any(x in label for x in ["熱量", "蛋白質", "豆魚"]):
                    val_raw = str(df.iloc[r_idx, col]).strip()
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # 抓包點：如果是空的
                    if val_raw == "" or val_raw == "0" or "None" in val_raw:
                        cell.fill, cell.font = STYLE["EMPTY"]["fill"], STYLE["EMPTY"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "數據缺失", "原因": f"重大缺失：{label} 標示不可為空"})
                    else:
                        num = float(re.findall(r"\d+\.?\d*", val_raw)[0]) if re.findall(r"\d+\.?\d*", val_raw) else 0.0
                        # 檢查數值是否符合法規 (略)
    
    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
# (介面略)
