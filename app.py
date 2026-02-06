import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(新北食品) 菜單稽核系統", layout="wide")

# --- 定義視覺規範 (30級字 + 微軟正黑體) ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式設定
PORTION_STYLE = {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}
CALORIE_STYLE = {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)}
REPEAT_STYLE  = {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
SPICY_STYLE   = {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)}

# --- 🎯 學部營養基準 (依據新北市學校午餐食物內容基準) ---
# 國小取高年級基準，高中取男/女平均基準
SCHOOL_LEVELS = {
    "國小部": {"熱量": (650, 750), "全榖": 3.5, "蛋白質": 3.5, "蔬菜": 1.5},
    "國中部": {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0},
    "高中部": {"熱量": (750, 950), "全榖": 4.5, "蛋白質": 4.5, "蔬菜": 2.0}
}

# 食材肉種識別字典 (僅用於當天 A/B 餐避讓)
MEAT_DICT = {"豬": ["豬", "肉絲", "肉片", "排骨", "焢肉", "培根", "火腿", "里肌"], "雞": ["雞", "翅", "鳳", "咔啦", "柳", "腿"], "牛": ["牛"], "魚": ["魚", "吻仔", "海鮮", "蝦"], "蛋": ["蛋"], "豆": ["豆", "腐", "干", "素肉"]}

def get_meat(text):
    if not text or any(x in text for x in ["水果", "Fruit", "甜湯", "湯品"]): return None
    for key, words in MEAT_DICT.items():
        if any(w in text for w in words): return key
    return text[:2] if len(text) >= 2 else None

def to_num(val):
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def audit_process(file, level):
    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()
        std = SCHOOL_LEVELS[level] # 獲取對應學部標準

        for sn, df in sheets_df.items():
            df = df.fillna("")
            ws = wb[sn]
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8): 
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]
                if "202" not in date_str: continue
                day_name = str(df.iloc[d_row+1, col]) if (d_row+1) < len(df) else ""

                # --- 1. 營養標示審核 (對標選定學部) ---
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    val = to_num(df.iloc[r_idx, col])
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    if "熱量" in label and (val < std["熱量"][0] or val > std["熱量"][1]):
                        cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal (基準:{std['熱量'][0]}-{std['熱量'][1]})"})
                    elif "全榖" in label and val < std["全榖"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "全榖", "原因": f"紅底：不足{std['全榖']}份"})
                    elif "豆魚" in label and val < std["蛋白質"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "蛋白質", "原因": f"紅底：不足{std['蛋白質']}份"})
                    elif "蔬菜" in label and val < std["蔬菜"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "蔬菜", "原因": f"紅底：不足{std['蔬菜']}份"})

                # --- 2. 食材重複性審核 (A/B 餐道避讓) ---
                main_A_idx = d_row + 3
                meat_A = get_meat(str(df.iloc[main_A_idx, col]))
                label_B = next((i for i in range(d_row+5, len(df)) if "輕食B餐" in str(df.iloc[i, 2])), None)
                main_B_idx = label_B + 1 if label_B else None
                meat_B = get_meat(str(df.iloc[main_B_idx, col])) if main_B_idx else None

                if meat_A and meat_B and meat_A == meat_B:
                    for r in [main_A_idx, main_B_idx]:
                        if r and r < len(df):
                            c = ws.cell(row=r+1, column=col+1)
                            c.fill, c.font = REPEAT_STYLE["fill"], REPEAT_STYLE["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "餐道衝突", "原因": f"黃底紅字：A/B餐主食重複({meat_A})"})

                # --- 3. 禁辣原則審核 ---
                if any(day in day_name for day in ["週一", "週二", "週四"]):
                    for r_idx in range(d_row + 2, d_row + 12):
                        if r_idx >= len(df) or "水果" in str(df.iloc[r_idx, 2]): continue
                        txt = str(df.iloc[r_idx, col])
                        if "●" in txt or "🌶️" in txt:
                            cell = ws.cell(row=r_idx+1, column=col+1)
                            cell.fill, cell.font = SPICY_STYLE["fill"], SPICY_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": "禁辣日", "原因": "淺綠底：標註辣味標示"})

        wb.save(output)
        return logs, output.getvalue()
    except Exception as e:
        return [f"發生錯誤：{str(e)}"], None

# --- 介面呈現 ---
st.title("🛡️ 輕食區菜單自主稽核系統 (新北食品優化版)")
st.caption("製作者：Alison")

# 增加學部標準選擇器
target_level = st.radio(
    "請選擇審核學部：", 
    ["國小部", "國中部", "高中部"], 
    index=1, 
    horizontal=True,
    help="系統將根據選擇自動調整熱量、全榖、蛋白質與蔬菜的檢核基準。"
)

st.markdown("---")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    with st.spinner(f"正在以 {target_level} 基準進行稽核..."):
        results, data = audit_process(up, target_level)
        if data:
            if results:
                st.error(f"🚩 偵測到 {len(results)} 項異常 (學部基準：{target_level})")
                st.download_button("📥 下載退件標註檔", data, f"稽核結果_{up.name}")
                st.table(pd.DataFrame(results))
            else:
                st.success(f"🎉 完美！通過 {target_level} 所有稽核項目。")
        else:
            st.error(results[0])
