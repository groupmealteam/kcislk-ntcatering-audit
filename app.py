import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# 樣式：黑底白字 30 級 / 黃底紅字 20 級
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

# 2. 正確模式選擇 (美食街、小學部、幼兒園、素食)
mode = st.sidebar.selectbox("📋 選擇部別：", ["美食街", "小學部", "幼兒園", "素食菜單"])

def audit_process(file, mode):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制轉字串並標記空白
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', ''], 'MISSING')
        
        # 決定標籤欄位：美食街固定看 C 欄 (Index 2)；其餘看 A 欄 (Index 0)
        label_col = 2 if mode == "美食街" else 0
        data_cols = range(3, 8) if mode == "美食街" else range(1, 6)
        
        # 找到「日期」定錨
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[label_col])), None)
        if d_row is None: continue

        for col in data_cols:
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, label_col]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心缺失偵測 ---
                critical_tags = ["熱量", "主食", "主菜", "副菜", "套餐"]
                if any(t in label for t in critical_tags):
                    
                    # A. 針對熱量：只要是 MISSING 就噴黑 (解決 4/28, 4/29 熱量紅框)
                    if "熱量" in label and content == "MISSING":
                        cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                        logs.append({"日期": date_val, "項目": label, "缺失": "⚠️ 熱量數據缺失"})

                    # B. 針對菜名：如果這格空，但下一格有食材 (解決 4/29 副菜紅框)
                    elif content == "MISSING":
                        try:
                            # 往下看一格是不是有寫食材 (包含 + 號或複數食材)
                            next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_val != "MISSING":
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "項目": label, "缺失": "❌ 菜名漏填 (下方有食材)"})
                        except: pass

                # --- 規格稽核 ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, spec in specs.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "項目": label, "缺失": f"{item} 未標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.markdown(f"**部別：{mode}**")

up = st.file_uploader(f"📂 請上傳【{mode}】菜單 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up, mode)
    if results:
        st.error(f"🚩 發現 {len(results)} 項缺失")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 檢查完畢，未發現缺失")
