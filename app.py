import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題與註解鎖死)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (針對 4/28-4/29 這種黑洞)
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
        # 修正：不准跳過 NaN，全部強制變為字串，讓程式「看見」空洞
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0'], '')
        
        # 定位「日期」列 (定錨)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        # 掃描 D 到 H 欄 (週一至週五)
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列往下，每一格都必須接受「標籤審核」
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 偵測 A：熱量與菜名缺失 (針對 4/28-4/29 紅框)
                critical_tags = ["熱量", "主菜", "副菜", "套餐"]
                if any(tag in label for tag in critical_tags):
                    # 如果這格是空的
                    if content == "":
                        # 特別針對 4/29 副菜：如果這格空，但下一格「食材明細」有字，必抓！
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "" or "熱量" in label:
                                cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                                logs.append({"日期": date_val, "缺失": "嚴重缺失", "原因": f"❌ {label} 欄位空白！"})
                        except: pass

                # 偵測 B：原本的規格審核 (白帶魚 150g)
                check_specs = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, spec in check_specs.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格缺失", "原因": f"{item} 需標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面 (Alison 原始設定) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 上傳 Excel 進行最後審核", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項嚴重缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
    else:
        st.success("✅ 恭喜！結構完整，規格正確。")
