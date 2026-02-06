import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 專業稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 視覺規範：30級字 + 微軟正黑體
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式設定
STYLE = {
    "PORTION": {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)},
    "CALORIE": {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)},
    "SPICY":   {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)},
    "CONTRACT_FAIL": {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
}

# --- 🎯 精確對標：學制營養基準 (依據原則定稿版) ---
STD_MAP = {
    "幼兒園": {"熱量": (350, 480), "全榖": 2.0, "蛋白質": 2.0, "蔬菜": 1.0},
    "小學":   {"熱量": (650, 780), "全榖": 3.0, "蛋白質": 3.0, "蔬菜": 1.5},
    "美食街": {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0},
    "素食":   {"熱量": (700, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}
}

# --- 🎯 規格對標：增補協議書重量限制 ---
CONTRACT_SPECS = {
    "獅子頭": "60gX2",
    "鯰魚片": "120g",
    "漢堡排": "150g",
    "烤肉串": "80g",
    "小卷": "100g"
}

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return 0.0
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def audit_process(file):
    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            df = df.fillna("")
            ws = wb[sn]
            current_std = next((STD_MAP[k] for k in STD_MAP if k in sn), None)
            if not current_std: continue

            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                if col >= len(df.columns): break
                date_val = str(df.iloc[d_row, col]).split(" ")[0]
                if "202" not in date_val: continue
                day_name = str(df.iloc[d_row+1, col])

                # --- 1. 原則四 & 五：標示與禁辣檢查 ---
                for r_idx in range(d_row + 2, d_row + 15):
                    if r_idx >= len(df): break
                    txt = str(df.iloc[r_idx, col])
                    cell = ws.cell(row=r_idx+1, column=col+1)

                    # 禁辣 (原則五)
                    if any(d in day_name for d in ["週一", "週二", "週四"]):
                        if any(x in txt for x in ["●", "🌶️", "辣", "椒"]):
                            cell.fill, cell.font = STYLE["SPICY"]["fill"], STYLE["SPICY"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "禁辣日違規", "原因": f"標示辣味({txt})"})

                    # 規格檢核 (增補協議書)
                    for item, spec in CONTRACT_SPECS.items():
                        if item in txt and spec not in txt.replace(" ", ""):
                            cell.fill, cell.font = STYLE["CONTRACT_FAIL"]["fill"], STYLE["CONTRACT_FAIL"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "規格不符", "原因": f"{item}需標註{spec}"})

                # --- 2. 營養法規紅線 ---
                for r_idx in range(d_row + 1, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    val = to_num(df.iloc[r_idx, col])
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    if "熱量" in label:
                        if val < current_std["熱量"][0] or val > current_std["熱量"][1]:
                            cell.fill, cell.font = STYLE["CALORIE"]["fill"], STYLE["CALORIE"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "熱量", "原因": f"不符{current_std['熱量']}"})
                    elif any(k in label for k in ["全榖", "豆魚", "蔬菜"]):
                        key = "全榖" if "全榖" in label else "蛋白質" if "豆魚" in label else "蔬菜"
                        if val < current_std[key]:
                            cell.fill, cell.font = STYLE["PORTION"]["fill"], STYLE["PORTION"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "份數不足", "原因": f"{key}低於{current_std[key]}"})

        wb.save(output)
        return logs, output.getvalue()
    except Exception as e:
        return [{"分頁": "系統錯誤", "原因": str(e)}], None

# --- UI ---
st.title("🛡️ 團膳區(新北食品) 專業稽核系統")
st.caption("製作者：Alison")
st.markdown("---")
st.info("📌 **本系統已載入 114/8/1 生效之增補協議規格 (如：獅子頭 60gX2, 漢堡排 150g 等)**")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    with st.spinner("正在對標合約條款..."):
        results, data = audit_process(up)
        if results:
            st.error(f"🚩 偵測到 {len(results)} 項不符規範")
            st.table(pd.DataFrame(results))
            st.download_button("📥 下載標註檔", data, f"稽核結果_{up.name}")
        else:
            st.success("🎉 通過所有合約與原則稽核！")
