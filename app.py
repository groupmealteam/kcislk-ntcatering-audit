import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題與註解完全鎖死，這是我最後的誠意)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (專殺 4/28-4/29 的空白漏洞)
STYLE = {
    "BLACK_CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=30, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 修正核心 BUG：不准跳過 NaN，全部強制轉為 "VOID"
        df_audit = df.fillna("VOID")
        
        # 尋找日期 Row (C 欄位定錨)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        # 掃描週一到週五 (D-H 欄)
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列往下掃描所有標籤，只要標籤在，內容就必須在
            for r_idx in range(d_row + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 偵測 A：熱量黑洞 (4/28, 4/29 晚餐熱量空白)
                if "熱量" in label:
                    if content in ["VOID", "", "nan", "0"]:
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量欄位被挖空！(重大違規)"})

                # 偵測 B：幽靈菜名 (4/29 菜名空白但有明細)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    try:
                        detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                        # 如果菜名是空的 (VOID)，但下一格食材明細不是空的，就是抓包！
                        if content == "VOID" and detail_val != "VOID":
                            cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                            logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名(只有明細)"})
                    except: pass

                # 偵測 C：規格缺失 (白帶魚/獅子頭)
                specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案 (這版抓不到我就報廢)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項嚴重缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
