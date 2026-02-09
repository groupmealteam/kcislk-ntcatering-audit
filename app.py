import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (維持 Alison 的原始風格)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式設定：黑底白字 30 級 (專殺 4/28-4/29 的空白) / 黃底紅字 (殺規格)
STYLE = {
    "BLACK_CRITICAL": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_SPEC": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    # 關鍵修正：強迫讀取所有內容為字串，並把 NaN 填補為特定的字串 "VOID_ERROR"
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        df_audit = df.fillna("VOID_ERROR")
        
        # 尋找日期列
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # 掃描週一到週五
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心邏輯：強制偵測空白 (針對 4/28, 4/29 紅框) ---
                # 只要左邊標籤有這些關鍵字，右邊如果是 VOID_ERROR 或空白，一律噴黑漆
                mandatory_labels = ["熱量", "主菜", "副菜", "套餐", "主食"]
                
                if any(tag in label for tag in mandatory_labels):
                    if content in ["VOID_ERROR", "", "nan", "0"]:
                        # 4/29 特殊漏填：菜名空，但下一行(食材)有字，必抓
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            # 針對熱量或是有明細無菜名的情況
                            if "熱量" in label or detail_val != "VOID_ERROR":
                                cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                                logs.append({"日期": date_val, "缺失": "不完整", "原因": f"❌ {label} 欄位漏填！"})
                        except: pass

                # --- 核心邏輯：規格稽核 (原本穩定的功能) ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格缺失", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單 Excel (測試 4/28-4/29 空白黑洞)", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項嚴重缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構與規格完美，這份菜單沒問題！")
