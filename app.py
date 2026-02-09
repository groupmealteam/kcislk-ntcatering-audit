import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 樣式固定：黑底白字(缺失)、黃底紅字(規格不符)
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=20, color="FFFFFF", bold=True)}
STYLE_SPEC = {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}

def final_audit_process(file):
    fname = file.name
    # 第一階段：檔名優先判讀 (標題與座標鎖定)
    if "美食街" in fname:
        mode, label_idx, data_indices = "美食街", 2, [3, 4, 5, 6, 7]
    elif any(kw in fname for kw in ["小學", "幼兒園"]):
        mode, label_idx, data_indices = "教育學部", 0, [1, 2, 3, 4, 5]
    else:
        return None, "INVALID_FILENAME", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 預處理：將所有偽裝空值、合併格空值徹底轉換為純空字串
        df_audit = df.astype(str).applymap(lambda x: "" if str(x).strip().lower() in ['nan', 'none', '0', '0.0', ''] else str(x).strip())
        
        max_rows, max_cols = df_audit.shape

        for r_idx in range(max_rows):
            if label_idx >= max_cols: continue
            # 清理標籤內的干擾符號
            label = df_audit.iloc[r_idx, label_idx].replace('\n', '').strip()
            
            # 鎖定合約規範標籤
            target_tags = ["熱量", "主食", "主菜", "副菜", "套餐", "湯品"]
            if any(t in label for t in target_tags):
                for c_idx in data_indices:
                    if c_idx >= max_cols: continue
                    
                    # 跳過「放假判讀」：如果整天(該欄)都是空的，則不視為缺失
                    col_data = df_audit.iloc[:, c_idx].str.cat()
                    if len(col_data) == 0: continue 
                    
                    # 跳過「週一早餐」：若是週一且為早餐標籤，則忽略
                    # (註：此處需搭配分頁日期判斷，簡易邏輯為略過特定標籤組合)
                    if "早餐" in label and "一" in sn: continue

                    content = df_audit.iloc[r_idx, c_idx]
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)

                    # 核心判斷：抓紅框缺失 (內容空、明細有)
                    if content == "":
                        try:
                            detail = df_audit.iloc[r_idx+1, c_idx]
                            if detail != "": # 抓到 4/29 的死穴
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填菜名"
                                logs.append({"分頁": sn, "項目": label, "缺失": "菜名空白但明細有字"})
                        except: pass
                    
                    # 熱量專屬稽核
                    if "熱量" in label and content == "":
                        cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                        logs.append({"分頁": sn, "項目": label, "缺失": "數值漏填"})

            # 重量規格審核 (合約原則詳實記錄)
            specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
            for c_idx in data_indices:
                if c_idx >= max_cols: continue
                val = df_audit.iloc[r_idx, c_idx]
                for item, spec in specs.items():
                    if item in val and spec not in val.replace(" ", ""):
                        cell = ws.cell(row=r_idx+1, column=c_idx+1)
                        cell.fill, cell.font = STYLE_SPEC["fill"], STYLE_SPEC["font"]
                        logs.append({"分頁": sn, "項目": item, "缺失": f"未標註規格 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
up = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])
if up:
    logs, m, data = final_audit_process(up)
    if m == "INVALID_FILENAME":
        st.error("❌ 第一階段失敗：檔名不符。請包含「美食街」或「小學/幼兒園」。")
    else:
        st.info(f"📁 判定模式：{m}")
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項缺失，已標註於 Excel 內。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載標註退件檔", data, f"退件_{up.name}")
        else:
            st.success("✅ 內容完整，且符合放假與週一判斷邏輯。")
