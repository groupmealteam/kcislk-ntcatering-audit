import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 樣式：黑底白字(缺失)、黃底紅字(規格不符)
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=20, color="FFFFFF", bold=True)}
STYLE_SPEC = {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}

def final_audit(file):
    fname = file.name
    # 第一階段：檔名與座標鎖定
    if "美食街" in fname:
        mode, l_col, d_cols = "美食街", 2, [3, 4, 5, 6, 7]
    elif any(k in fname for k in ["小學", "幼兒園"]):
        mode, l_col, d_cols = "教育學部", 0, [1, 2, 3, 4, 5]
    else:
        return None, "INVALID_NAME", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    # 第二階段：內容深度審核
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 預處理：把所有 '0', 'nan', '空格' 統一化，讓黑洞現形
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', '　'], '')
        
        for r_idx, row in df_audit.iterrows():
            label = str(row[l_col]).strip()
            
            # 1. 關鍵標籤缺失抓取
            targets = ["熱量", "主食", "主菜", "副菜", "套餐"]
            if any(t in label for t in targets):
                for c_idx in d_cols:
                    content = str(df_audit.iloc[r_idx, c_idx]).strip()
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)

                    # A. 熱量完全漏填
                    if "熱量" in label and content == "":
                        cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                        logs.append({"分頁": sn, "缺失": f"{label} 漏填"})

                    # B. 菜名漏填聯動 (專抓 4/29 副菜漏洞)
                    elif content == "" and any(x in label for x in ["主菜", "副菜"]):
                        try:
                            # 往下看一列，如果明細有東西，這格必噴黑
                            detail = str(df_audit.iloc[r_idx+1, c_idx]).strip()
                            if detail != "":
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                logs.append({"分頁": sn, "缺失": f"{label} 漏填菜名(明細有字)"})
                        except: pass

            # 2. 規格強硬審核 (白帶魚/漢堡排)
            specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
            for c_idx in d_cols:
                item_content = str(df_audit.iloc[r_idx, c_idx])
                for fish, weight in specs.items():
                    if fish in item_content and weight not in item_content.replace(" ", ""):
                        cell = ws.cell(row=r_idx+1, column=c_idx+1)
                        cell.fill, cell.font = STYLE_SPEC["fill"], STYLE_SPEC["font"]
                        logs.append({"分頁": sn, "缺失": f"{fish} 規格未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- 簡潔介面 ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
up = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if up:
    logs, m, data = final_audit(up)
    if m == "INVALID_NAME":
        st.error("❌ 檔名無法辨識，請包含『美食街』或『小學』關鍵字。")
    else:
        st.info(f"📁 判定模式：{m}")
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項不完整或規格錯誤！")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
        else:
            st.success("✅ 內容完整，未發現缺失。")
