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
FONT_SIZE = 30

# 樣式表：對應四大違規等級
STYLE = {
    "MISSING":   {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}, # 黑底白字：重大缺失/漏填
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF")},      # 紅底白字：數據違規
    "CONTRACT":  {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}, # 黃底紅字：合約規格
    "SPICY":     {"fill": PatternFill("solid", fgColor="C6EFCE"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000")}       # 綠底黑字：禁辣日違規
}

# 根據《SE1140803 增補協議書》附件二：規格紅線
CONTRACT_SPECS = {
    "獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", 
    "白蝦": "X3", "烤肉串": "80gX2", "白帶魚": "150g", "小卷": "100g", "砂鍋魚丁": "250g"
}

# 根據《審閱原則_修訂2》：營養基準紅線
NUTRITION_STD = {
    "幼兒園": {"熱量": (350, 480), "蛋白質": 2.0},
    "小學":   {"熱量": (650, 780), "蛋白質": 3.0},
    "美食街": {"熱量": (750, 850), "蛋白質": 4.0},
    "素食":   {"熱量": (700, 850), "蛋白質": 4.0}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        # 初始化：將所有 NaN 轉為 "EMPTY" 以便精準抓包
        df_audit = df.fillna("EMPTY")
        ws = wb[sn]
        
        # 匹配學部標準
        std_key = next((k for k in NUTRITION_STD if k in sn), None)
        if not std_key: continue
        std = NUTRITION_STD[std_key]

        # 定位「日期Date」所在列 (新北食品標準 C欄)
        d_row = next((i for i, r in df_audit.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # 週一到週五
            date_val = str(df_audit.iloc[d_row, col]).split(" ")[0]
            day_name = str(df_audit.iloc[d_row+1, col])

            # --- A. 結構完整性抓包 (原則一) ---
            # 檢查主食、主菜、副菜、青菜、湯品 5 項
            for offset in range(2, 7):
                r_idx = d_row + offset
                val = str(df_audit.iloc[r_idx, col]).strip()
                if val in ["EMPTY", "", "nan"]:
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    cell.fill, cell.font = STYLE["MISSING"]["fill"], STYLE["MISSING"]["font"]
                    logs.append({"分頁": sn, "日期": date_val, "項目": "結構缺項", "原因": "❌ 菜名空白！(原則一)"})

            # --- B. 內容與規格稽核 (原則四、五 & 增補協議) ---
            for r_idx in range(d_row + 2, d_row + 20):
                txt = str(df_audit.iloc[r_idx, col])
                cell = ws.cell(row=r_idx+1, column=col+1)
                
                # 1. 禁辣日檢查 (週一、二、四)
                if any(d in day_name for d in ["週一", "週二", "週四"]):
                    if any(x in txt for x in ["🌶️", "●", "辣", "椒", "麻", "沙茶"]):
                        cell.fill, cell.font = STYLE["SPICY"]["fill"], STYLE["SPICY"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "禁辣違規", "原因": f"違反原則五: {txt}"})

                # 2. 合約規格對標 (新規格地雷)
                for item, spec in CONTRACT_SPECS.items():
                    if item in txt and spec not in txt.replace(" ", ""):
                        cell.fill, cell.font = STYLE["CONTRACT"]["fill"], STYLE["CONTRACT"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "規格違規", "原因": f"{item}需標註 {spec}"})

            # --- C. 營養數據完整性與紅線 (修訂2) ---
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2])
                val_raw = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)
                
                if any(x in label for x in ["熱量", "蛋白質", "豆魚"]):
                    # 抓包點：刪除數據
                    if val_raw in ["EMPTY", "0", "0.0"]:
                        cell.fill, cell.font = STYLE["MISSING"]["fill"], STYLE["MISSING"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "數據缺失", "原因": f"❌ {label} 標示不可為空"})
                    else:
                        num = float(re.findall(r"\d+\.?\d*", val_raw)[0]) if re.findall(r"\d+\.?\d*", val_raw) else -1
                        # 紅線判定
                        if "熱量" in label and (num < std["熱量"][0] or num > std["熱量"][1]):
                            cell.fill, cell.font = STYLE["DATA_FAIL"]["fill"], STYLE["DATA_FAIL"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "熱量超標", "原因": f"應在 {std['熱量']}"})
                        elif ("蛋白質" in label or "豆魚" in label) and num < std["蛋白質"]:
                            cell.fill, cell.font = STYLE["DATA_FAIL"]["fill"], STYLE["DATA_FAIL"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "項目": "份數不足", "原因": f"低於 {std['蛋白質']} 份"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- Streamlit 介面渲染 ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 上傳待審菜單 Excel", type=["xlsx"])

if up:
    with st.spinner("正在對標 114 學年合約與審閱原則..."):
        results, processed_data = audit_process(up)
        
        if results:
            st.error(f"🚩 稽核完畢：發現 {len(results)} 項不符規範（含重大缺失）")
            st.table(pd.DataFrame(results))
            st.download_button(
                label="📥 下載退件標註檔案",
                data=processed_data,
                file_name=f"退件建議_{up.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("🎉 通過稽核！菜單結構完整且符合合約規格。")
