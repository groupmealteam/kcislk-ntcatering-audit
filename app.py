import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 標題嚴格鎖定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# 樣式定義：黑底白字 30 級 / 黃底紅字 20 級
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

# 2. 選單：讓妳決定現在要審哪一種，不准混在一起
mode = st.sidebar.radio("📋 選擇審核目標：", ["美食街 (標籤在C欄)", "小學部/幼兒園 (標籤在A欄)"])

def audit_process(file, mode):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 核心修正：將所有 NaN 或 0 或 None 全部轉為 "MISSING" 標籤
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' '], 'MISSING')
        
        # 定位標籤欄：美食街看第 2 欄(C)，小學/幼兒園看第 0 欄(A)
        label_col = 2 if "美食街" in mode else 0
        data_cols = range(3, 8) if "美食街" in mode else range(1, 6)

        for r_idx, row in df_audit.iterrows():
            label = str(row[label_col]).strip()
            
            # --- 精準捕捉標籤 ---
            target_tags = ["熱量", "主菜", "副菜", "套餐", "主食"]
            if any(t in label for t in target_tags):
                
                for c_idx in data_cols:
                    content = str(df_audit.iloc[r_idx, c_idx]).strip()
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)
                    
                    # 偵測 A：熱量黑洞 (針對 4/28, 4/29)
                    if "熱量" in label and content == "MISSING":
                        cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                        logs.append({"日期": f"第{c_idx-2}天", "項目": label, "原因": "⚠️ 熱量漏填"})

                    # 偵測 B：菜名消失但食材有字 (針對 4/29 副菜)
                    # 邏輯：如果這格空，但同一欄的「下一列」不是 MISSING，代表漏了菜名
                    elif content == "MISSING":
                        try:
                            next_row_val = str(df_audit.iloc[r_idx+1, c_idx]).strip()
                            if next_row_val != "MISSING":
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": f"第{c_idx-2}天", "項目": label, "原因": "❌ 漏填菜名 (但有填食材)"})
                        except: pass

                    # 偵測 C：規格缺失
                    specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                    for item, spec in specs.items():
                        if item in content and spec not in content.replace(" ", ""):
                            cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                            logs.append({"日期": f"第{c_idx-2}天", "項目": label, "原因": f"{item} 未標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("團膳區(新北食品) 全方位稽核系統")
st.markdown(f"--- 模式：**{mode}** ---")

up = st.file_uploader("📂 上傳有缺失的新北菜單 (xlsx)", type=["xlsx"])
if up:
    results, data = audit_process(up, mode)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項缺失！請看下方表格與標註檔。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完整，未發現缺失。")
