import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 嚴格鎖定標題
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    fname = file.name
    # 第一關：檔名判讀 (自動切換美食街或小學部邏輯)
    if "美食街" in fname:
        target_mode = "美食街"
        label_col = 2  # C欄為標籤
        data_cols = [3, 4, 5, 6, 7] # D-H 欄為數據
    else:
        target_mode = "小學/幼兒園"
        label_col = 0  # A欄為標籤
        data_cols = [1, 2, 3, 4, 5] # B-F 欄為數據

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵：保留原始結構，不隨便填充，才能抓到「空值」
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' '], '')

        for r_idx in range(len(df_audit)):
            label = str(df_audit.iloc[r_idx, label_col]).strip()
            
            # 鎖定關鍵標籤 (熱量、主菜、副菜...)
            critical_tags = ["熱量", "主食", "主菜", "副菜", "套餐"]
            if any(t in label for t in critical_tags):
                for c_idx in data_cols:
                    try:
                        content = df_audit.iloc[r_idx, c_idx].strip()
                        cell = ws.cell(row=r_idx+1, column=c_idx+1)
                        
                        # --- 判讀核心 A：熱量缺失 ---
                        if "熱量" in label and content == "":
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"分頁": sn, "項目": label, "原因": "❌ 熱量格完全空白"})

                        # --- 判讀核心 B：菜名缺失聯動 (專殺 4/29 副菜漏洞) ---
                        elif content == "":
                            # 檢查下一行(食材明細)是否有內容
                            next_row_val = str(df_audit.iloc[r_idx+1, c_idx]).strip()
                            if next_row_val != "":
                                cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                                logs.append({"分頁": sn, "項目": label, "原因": "⚠️ 漏填菜名 (下方食材有內容)"})
                    except: pass

            # --- 判讀核心 C：規格嚴審 ---
            specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
            for c_idx in data_cols:
                content = str(df_audit.iloc[r_idx, c_idx])
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell = ws.cell(row=r_idx+1, column=c_idx+1)
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"分頁": sn, "項目": "規格錯誤", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue(), target_mode

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
up = st.file_uploader("📂 上傳 Excel 檔案（系統將自動判讀檔名與內容條件）", type=["xlsx"])

if up:
    logs, data, detected_mode = audit_process(up)
    st.info(f"📁 檔名判讀結果：**{detected_mode} 模式**")
    
    if logs:
        st.error(f"🚩 抓到 {len(logs)} 項缺失（包含 4/28-4/29 紅框位置）")
        st.table(pd.DataFrame(logs))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 檢查完畢，未發現缺失。")
