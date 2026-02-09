import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 頁面設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# 樣式設定
STYLE_MISSING = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=24, color="FFFFFF", bold=True)}
STYLE_SPEC = {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}

def audit_process(file):
    fname = file.name
    # --- 第一階段：檔名優先判讀 ---
    if "美食街" in fname:
        mode = "美食街"
        label_col = 2  # C欄
        data_cols = [3, 4, 5, 6, 7] # D-H 欄
    elif "小學" in fname or "幼兒園" in fname:
        mode = "教育學部"
        label_col = 0  # A欄
        data_cols = [1, 2, 3, 4, 5] # B-F 欄
    else:
        # 檔名不對，直接判定為無法判讀
        return None, "INVALID_FILENAME", None

    # --- 第二階段：內容完整性判讀 ---
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 將空值統一標記，避免判讀引擎裝瞎
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', ''], 'MISSING')
        
        for r_idx, row in df_audit.iterrows():
            label = str(row[label_col]).strip()
            
            # 偵測關鍵標籤
            target_tags = ["熱量", "主食", "主菜", "副菜", "套餐"]
            if any(t in label for t in target_tags):
                for c_idx in data_cols:
                    content = df_audit.iloc[r_idx, c_idx].strip()
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)
                    
                    # 缺失判定 A：熱量格完全沒填
                    if "熱量" in label and content == "MISSING":
                        cell.fill, cell.font = STYLE_MISSING["fill"], STYLE_MISSING["font"]
                        logs.append({"分頁": sn, "缺失項目": label, "原因": "❌ 熱量數據完全漏填"})
                    
                    # 缺失判定 B：菜名格空，但下一行(食材)有寫字 (針對 4/29 副菜)
                    elif content == "MISSING":
                        try:
                            next_val = str(df_audit.iloc[r_idx+1, c_idx]).strip()
                            if next_val != "MISSING":
                                cell.fill, cell.font = STYLE_MISSING["fill"], STYLE_MISSING["font"]
                                logs.append({"分頁": sn, "缺失項目": label, "原因": "⚠️ 漏填菜名(下方食材有內容)"})
                        except: pass

                    # 規格判定 C：白帶魚(150g), 漢堡排(150g), 獅子頭(60gX2)
                    specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                    for item, spec in specs.items():
                        if item in content and spec not in content.replace(" ", ""):
                            cell.fill, cell.font = STYLE_SPEC["fill"], STYLE_SPEC["font"]
                            logs.append({"分頁": sn, "缺失項目": "規格錯誤", "原因": f"{item} 未標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# 網頁介面
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])

if up:
    logs, result_mode, data = audit_process(up)
    
    if result_mode == "INVALID_FILENAME":
        st.error(f"❌ **第一階段判讀失敗**：檔名『{up.name}』不符合規範。")
        st.warning("請確認檔名是否包含「美食街」或「小學/幼兒園」關鍵字。")
    else:
        st.success(f"✅ **第一階段通過**：偵測到『{result_mode}』模式。進入第二階段內容稽核...")
        
        if logs:
            st.error(f"🚩 **第二階段結果**：發現 {len(logs)} 項內容不完整或錯誤！")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載缺失標註檔", data, f"退件_{up.name}")
        else:
            st.success("🎉 **第二階段通過**：內容完整，規格全數正確！")
