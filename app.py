import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(新北食品) 全方位稽核系統", layout="wide")

# --- 定義視覺規範 (30級字 + 微軟正黑體) ---
# 註解：製作者 Alison
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式設定
PORTION_STYLE = {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}
CALORIE_STYLE = {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)}
REPEAT_STYLE  = {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
SPICY_STYLE   = {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)}

# --- 🎯 多學制營養基準字典 (自動依分頁關鍵字識別) ---
STD_MAP = {
    "幼兒園": {"熱量": (350, 450), "全榖": 2.0, "蛋白質": 2.0, "蔬菜": 1.0},
    "小學":   {"熱量": (650, 750), "全榖": 3.5, "蛋白質": 3.5, "蔬菜": 1.5},
    "美食街": {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0},
    "素食":   {"熱量": (700, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}
}

MEAT_DICT = {"豬": ["豬", "肉絲", "肉片", "排骨", "焢肉", "培根", "火腿"], "雞": ["雞", "翅", "鳳", "咔啦", "柳"], "牛": ["牛"], "魚": ["魚", "海鮮", "蝦"], "蛋": ["蛋"], "豆": ["豆", "腐", "干", "素肉"]}

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

def audit_process(file):
    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            df = df.fillna("")
            ws = wb[sn]
            
            # 自動識別分頁標準
            current_std = STD_MAP["美食街"]
            for key in STD_MAP.keys():
                if key in sn:
                    current_std = STD_MAP[key]
                    break

            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8): 
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]
                if "202" not in date_str: continue
                day_name = str(df.iloc[d_row+1, col]) if (d_row+1) < len(df) else ""

                # 1. 營養標示審核
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    val = to_num(df.iloc[r_idx, col])
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    if "熱量" in label and (val < current_std["熱量"][0] or val > current_std["熱量"][1]):
                        cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal"})
                    elif "全榖" in label and val < current_std["全榖"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "全榖", "原因": f"不足{current_std['全榖']}份"})
                    elif "豆魚" in label and val < current_std["蛋白質"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "蛋白質", "原因": f"不足{current_std['蛋白質']}份"})
                    elif "蔬菜" in label and val < current_std["蔬菜"]:
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "蔬菜", "原因": f"不足{current_std['蔬菜']}份"})

                # 2. 食材重複審核 (僅限當天 A/B 避讓，已刪除跨日重複)
                main_A_idx = d_row + 3
                meat_A = get_meat(str(df.iloc[main_A_idx, col]))
                label_B = next((i for i in range(d_row+5, len(df)) if "輕食B餐" in str(df.iloc[i, 2])), None)
                main_B_idx = label_B + 1 if label_B else None
                meat_B = get_meat(str(df.iloc[main_B_idx, col])) if main_B_idx else None

                if meat_A and meat_B and meat_A == meat_B:
                    for r in [main_A_idx, main_B_idx]:
                        if r and r < len(df):
                            ws.cell(row=r+1, column=col+1).fill = REPEAT_STYLE["fill"]
                            ws.cell(row=r+1, column=col+1).font = REPEAT_STYLE["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "餐道重複", "原因": f"黃底紅字：A/B餐皆為{meat_A}"})

                # 3. 禁辣原則審核
                if any(day in day_name for day in ["週一", "週二", "週四"]):
                    for r_idx in range(d_row + 2, d_row + 15):
                        if r_idx >= len(df) or "水果" in str(df.iloc[r_idx, 2]): continue
                        txt = str(df.iloc[r_idx, col])
                        if "●" in txt or "🌶️" in txt:
                            cell = ws.cell(row=r_idx+1, column=col+1)
                            cell.fill, cell.font = SPICY_STYLE["fill"], SPICY_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": "禁辣日", "原因": "淺綠底標辣"})

        wb.save(output)
        return logs, output.getvalue()
    except Exception as e:
        return [f"發生錯誤：{str(e)}"], None

# --- 介面呈現 ---
st.title("🛡️ 輕食區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.info("💡 系統會自動根據分頁名稱檢核：幼兒園、小學、素食、美食街基準。")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    with st.spinner("稽核中..."):
        results, data = audit_process(up)
        if data:
            if results:
                st.error(f"🚩 發現 {len(results)} 項異常")
                st.download_button("📥 下載退件標註檔", data, f"稽核結果_{up.name}")
                st.table(pd.DataFrame(results))
            else:
                st.success("🎉 完美！所有分頁皆通過合約稽核。")
        else:
            st.error(results[0])
