import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 定義視覺與合約紅線 ---
CONTRACT_SPECS = {"獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", "烤肉串": "80gX2", "白蝦": "X3"}
STD_MAP = {
    "幼兒園": {"熱量": (350, 480), "全榖": 2.0, "蛋白質": 2.0, "蔬菜": 1.0},
    "小學":   {"熱量": (650, 780), "全榖": 3.0, "蛋白質": 3.0, "蔬菜": 1.5},
    "美食街": {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}
}

# 樣式定義 (30級字)
STYLE = {
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF")},
    "CHEF_WARN": {"fill": PatternFill("solid", fgColor="FFCC00"), "font": Font(name="微軟正黑體", size=30, color="000000")}, # 大廚警告：口感或色澤
    "SPICY": {"fill": PatternFill("solid", fgColor="C6EFCE"), "font": Font(name="微軟正黑體", size=30)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        df = df.fillna("")
        ws = wb[sn]
        current_std = next((STD_MAP[k] for k in STD_MAP if k in sn), None)
        if not current_std: continue

        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8):
            day_name = str(df.iloc[d_row+1, col])
            menu_items = [str(df.iloc[r, col]) for r in range(d_row + 2, d_row + 15)]
            combined_text = "".join(menu_items)

            # --- 大廚審美 A: 烹調避讓 (原則六) ---
            if menu_items.count("◎") >= 2:
                logs.append({"分頁": sn, "項目": "大廚品味", "原因": "重複炸物(◎)：口感過於油膩"})
            if menu_items.count("燴") + menu_items.count("羹") >= 2:
                logs.append({"分頁": sn, "項目": "大廚品味", "原因": "重複勾芡：缺乏層次感"})

            # --- 審核官 B: 合約規格 ---
            for item, spec in CONTRACT_SPECS.items():
                if item in combined_text and spec not in combined_text:
                    logs.append({"分頁": sn, "項目": "合約規格", "原因": f"{item}規格應為{spec}"})

            # --- 營養師 C: 數據紅線 ---
            # (此處執行數值比對邏輯，若不符則標註 DATA_FAIL)

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 專業審閱系統")
st.caption("製作者：Alison | 整合『營養數據』與『大廚審美』")
# ... (Streamlit UI 程式碼)
