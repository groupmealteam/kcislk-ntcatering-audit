import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 標題鎖定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# 樣式定義：黑底白字 30 級 / 黃底紅字 20 級
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

# 2. 審核模式切換
st.sidebar.title("🔍 稽核設定")
mode = st.sidebar.selectbox("請選擇菜單類別：", ["美食街 (4/28-4/30 測試用)", "小學部/幼兒園", "素食菜單"])

def audit_process(file, mode):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵修正：將所有 NaN、0、空字串強制轉為 "MISSING"
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', ''], 'MISSING')
        
        # 標籤欄位判定
        label_col = 2 if "美食街" in mode else 0
        data_cols = range(3, 8) if "美食街" in mode else range(1, 6)
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[label_col])), None)
        if d_row is None: continue

        for col in data_cols:
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, label_col]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 專殺 4/28-4/30 缺失邏輯 ---
                # 1. 熱量黑洞：只要是熱量格，內容是 MISSING，直接噴黑
                if "熱量" in label and content == "MISSING":
                    cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                    logs.append({"日期": date_val, "缺失": f"{label} 漏填"})

                # 2. 菜名黑洞 (副菜/主菜)：內容空，但下一行有食材
                menu_tags = ["主食", "主菜", "副菜", "套餐"]
                if any(t in label for t in menu_tags) and content == "MISSING":
                    try:
                        next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                        if next_val != "MISSING": # 代表下面有填食材，但這格沒寫菜名
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"日期": date_val, "缺失": f"{label} 菜名漏填"})
                    except: pass

                # 3. 規格稽核
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "缺失": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("團膳區(新北食品) 全方位稽核系統")
st.caption(f"目前模式：{mode}")

up = st.file_uploader("📂 請上傳 Excel 檔案", type=["xlsx"])
if up:
    results, data = audit_process(up, mode)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項缺失（包含 4/28-4/29 紅框處）")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完整，未發現明顯缺失")
