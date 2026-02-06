import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
STYLE = {
    "BLACK_ALERT": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=24, color="FFFFFF", bold=True)},
    "YELLOW_SPEC": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵修正 1：先將 NaN 全部轉為特定字串，讓它「變為可見」
        df_audit = df.fillna("!!!MISSING!!!")
        
        # 定位日期標籤所在的 Row (通常在 C 欄)
        d_row = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row = i
                break
        if d_row is None: continue

        # 掃描週一到週五 (D-H 欄)
        for col in range(3, 8):
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列開始往下掃
            for r_idx in range(d_row + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 關鍵修正 2：標籤強制偵測 (針對熱量、菜名)
                # 如果左邊標籤有「熱量」、「套餐」等字眼，但內容是 MISSING，就噴黑漆
                mandatory_tags = ["熱量", "套餐", "主食", "主菜", "副菜"]
                if any(tag in label for tag in mandatory_tags):
                    # 偵測是否為空值
                    if content in ["!!!MISSING!!!", "", "0", "nan"]:
                        # 4/29 特殊邏輯：如果這格是空的，但下一格(明細)卻有字，這必抓
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "!!!MISSING!!!":
                                cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                                logs.append({"日期": date_val, "類別": "漏填缺失", "原因": f"❌ {label} 沒寫菜名但有食材"})
                        except: pass
                        
                        # 熱量強制檢查
                        if "熱量" in label:
                            cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                            logs.append({"日期": date_val, "類別": "漏填缺失", "原因": "⚠️ 熱量數據空白"})

                # 關鍵修正 3：原有規格稽核 (確保原本功能不壞掉)
                check_list = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for fish, spec in check_list.items():
                    if fish in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                        logs.append({"日期": date_val, "類別": "規格不符", "原因": f"{fish} 未標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 上傳 0428-0430 檔案測試最後一哩路", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項不完整或規格缺失。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件建議_{up.name}")
    else:
        st.success("✅ 結構與規格完美無缺！")
