import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (維持 Alison 的原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 核心修正：強迫程式看見空白，將 NaN 填補為字串 "MISSING"
        df_audit = df.fillna("MISSING")
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 偵測 A：強制空白查核 (專殺紅框缺失) ---
                # 只要左邊標籤有這些字，內容就絕對不能是 MISSING
                critical_tags = ["熱量", "主菜", "副菜", "套餐", "主食"]
                if any(tag in label for tag in critical_tags):
                    if content in ["MISSING", "", "nan", "0"]:
                        # 4/29 專用：若菜名空，但下面食材明細有字，必殺！
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "MISSING" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"日期": date_val, "缺失": "內容不全", "原因": f"❌ {label} 欄位未填！"})
                        except: pass

                # --- 偵測 B：原本穩定的規格稽核 ---
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格缺失", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單檔案 (最後測試：4/28-4/29 空白黑洞)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 發現 {len(results)} 項嚴重缺失，已完成標色。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔案", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完整，未發現明顯缺失。")
