import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁基本設定 (標題依照要求固定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 視覺規範與註解 (註解依照要求固定) ---
# 製作者 Alison
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

STYLE = {
    "PORTION": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)},
    "CALORIE": {"fill": PatternFill("solid", fgColor="FFCCFF"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)},
    "SPICY":   {"fill": PatternFill("solid", fgColor="C6EFCE"), "font": Font(name=FONT_NAME, size=FONT_SIZE)},
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
}

# --- 🎯 依據《增補協議書》附件二：規格絕對地雷 ---
# 這是 114/08/01 後的新標準，沒標到這些數值就是違規
MUST_SPECS = {
    "獅子頭": "60gX2",
    "漢堡排": "150g",
    "鯰魚片": "120g",
    "烤肉串": "80gX2",
    "白蝦": "X3",
    "白帶魚": "150g",
    "小卷": "100g",
    "砂鍋魚丁": "250g"
}

# --- 🎯 依據《審閱原則_修訂2》：營養數據基準 ---
STD_MAP = {
    "幼兒園": {"熱量": (350, 480), "蛋白質": 2.0, "蔬菜": 1.0},
    "小學":   {"熱量": (650, 780), "蛋白質": 3.0, "蔬菜": 1.5},
    "美食街": {"熱量": (750, 850), "蛋白質": 4.0, "蔬菜": 2.0}
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

        # 定位日期列：新北食品固定在第 3 欄 (C欄)
        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # 週一到週五
            day_name = str(df.iloc[d_row+1, col])
            date_val = str(df.iloc[d_row, col]).split(" ")[0]

            # --- A. 營養師專業：數據稽核 (對標審閱原則) ---
            for r_idx in range(len(df)):
                label = str(df.iloc[r_idx, 2])
                val_raw = str(df.iloc[r_idx, col])
                num = float(re.findall(r"\d+\.?\d*", val_raw)[0]) if re.findall(r"\d+\.?\d*", val_raw) else 0.0
                cell = ws.cell(row=r_idx+1, column=col+1)

                if "熱量" in label and (num < current_std["熱量"][0] or num > current_std["熱量"][1]):
                    cell.fill, cell.font = STYLE["CALORIE"]["fill"], STYLE["CALORIE"]["font"]
                    logs.append({"日期": date_val, "項目": "熱量異常", "原因": f"應在 {current_std['熱量']} 區間"})
                elif "豆魚" in label and num < current_std["蛋白質"]:
                    cell.fill, cell.font = STYLE["PORTION"]["fill"], STYLE["PORTION"]["font"]
                    logs.append({"日期": date_val, "項目": "蛋白質不足", "原因": f"低於 {current_std['蛋白質']} 份"})

            # --- B. 美食家品味與合約嚴謹度：內容稽核 ---
            for r_idx in range(d_row + 2, d_row + 20):
                txt = str(df.iloc[r_idx, col])
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 1. 禁辣 (原則五)
                if any(d in day_name for d in ["週一", "週二", "週四"]):
                    if any(x in txt for x in ["●", "🌶️", "辣", "椒", "沙茶"]):
                        cell.fill, cell.font = STYLE["SPICY"]["fill"], STYLE["SPICY"]["font"]
                        logs.append({"日期": date_val, "項目": "禁辣違規", "原因": f"禁辣日出現: {txt}"})

                # 2. 合約規格 (對標增補協議書附件二)
                for item, spec in MUST_SPECS.items():
                    if item in txt and spec not in txt.replace(" ", ""):
                        cell.fill, cell.font = STYLE["CONTRACT"]["fill"], STYLE["CONTRACT"]["font"]
                        logs.append({"日期": date_val, "項目": "規格違規", "原因": f"{item}需標註 {spec}"})

                # 3. 標示原則 (原則四)
                if "炸" in txt and "◎" not in txt:
                    cell.fill, cell.font = STYLE["CONTRACT"]["fill"], STYLE["CONTRACT"]["font"]
                    logs.append({"日期": date_val, "項目": "標示漏項", "原因": "炸物未標◎"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
up = st.file_uploader("👉 請上傳待審菜單", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 偵測到 {len(results)} 項不符規範（含合約規格與審閱原則）")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("🎉 通過所有合約規格與營養原則稽核。")
