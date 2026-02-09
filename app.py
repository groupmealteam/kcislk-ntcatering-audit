import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題嚴格鎖定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# 樣式定義：黑底白字 30 級 / 黃底紅字 20 級
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

# 2. 審核模式切換 (側邊欄)
mode = st.sidebar.selectbox("請選擇審核部別：", ["美食街", "小學部/幼兒園"])

def audit_process(file, mode):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵 BUG 修正：強制將所有空值(NaN)轉為字串 "EMPTY"
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0'], 'EMPTY')
        
        # 根據模式決定標籤在哪一欄 (美食街在 C 欄[index 2], 小學部在 A 欄[index 0])
        label_col = 2 if mode == "美食街" else 0
        data_cols = range(3, 8) if mode == "美食街" else range(1, 6)

        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[label_col])), None)
        if d_row is None: continue

        for col in data_cols:
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, label_col]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心 BUG 解決邏輯：針對 4/28-4/30 缺失 ---
                critical_tags = ["熱量", "主食", "主菜", "副菜", "套餐"]
                if any(tag in label for tag in critical_tags):
                    # 如果該格是 EMPTY 或是只有空白字元
                    if content == "EMPTY" or content == "":
                        # 針對 4/29 副菜漏填：如果這格空，但下一列(食材)有字，代表漏填菜名
                        is_missing = False
                        if "熱量" in label:
                            is_missing = True
                        else:
                            try:
                                next_row_val = str(df_audit.iloc[r_idx+1, col]).strip()
                                if next_row_val != "EMPTY": is_missing = True
                            except: pass
                        
                        if is_missing:
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"日期": date_val, "缺失": f"{label} 欄位空白"})

                # --- 規格稽核 ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "缺失": f"{item} 規格錯誤"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("團膳區(新北食品) 全方位稽核系統")
st.markdown(f"**目前模式：{mode}**")

up = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up, mode)
    if results:
        st.error(f"🚩 發現 {len(results)} 項缺失")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 未發現缺失")
