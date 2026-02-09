import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 樣式設定
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=20, color="FFFFFF", bold=True)}
STYLE_SPEC = {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}

def robust_audit(file):
    fname = file.name
    # --- 第一階段：檔名身分判斷 ---
    if "美食街" in fname:
        mode, label_idx, data_indices = "美食街", 2, [3, 4, 5, 6, 7] # C欄標籤, D-H數據
    elif any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode, label_idx, data_indices = "教育學部", 0, [1, 2, 3, 4, 5] # A欄標籤, B-F數據
    else:
        return None, "INVALID_FILENAME", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    # --- 第二階段：內容深度稽核 ---
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 清理數據，預防 nan 或 0 騙過程式
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', '　'], '')
        max_rows, max_cols = df_audit.shape

        for r_idx in range(max_rows):
            # 確保 label_idx 沒有超出這頁的範圍
            if label_idx >= max_cols: continue
            
            label = df_audit.iloc[r_idx, label_idx].strip()
            
            # 關鍵字過濾（解決換行符號問題）
            target_tags = ["熱量", "主食", "主菜", "副菜", "套餐"]
            if any(t in label for t in target_tags):
                for c_idx in data_indices:
                    # 防踩空：確保資料欄位在這頁的範圍內
                    if c_idx >= max_cols: continue
                    
                    content = df_audit.iloc[r_idx, c_idx].strip()
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)

                    # 1. 抓熱量缺失 (4/28, 4/29 紅框)
                    if "熱量" in label and content == "":
                        cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                        logs.append({"分頁": sn, "項目": label, "缺失": "❌ 熱量數值漏填"})

                    # 2. 抓菜名漏填聯動 (4/29 副菜紅框)
                    elif content == "" and any(x in label for x in ["主菜", "副菜", "主食"]):
                        # 檢查下一行明細
                        if r_idx + 1 < max_rows:
                            detail = df_audit.iloc[r_idx+1, c_idx].strip()
                            if detail != "":
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "⚠️漏填菜名"
                                logs.append({"分頁": sn, "項目": label, "缺失": "⚠️ 菜名空白但下方有食材"})

            # 3. 規格審核：白帶魚/漢堡排 (針對 150g)
            specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
            for c_idx in data_indices:
                if c_idx >= max_cols: continue
                raw_txt = df_audit.iloc[r_idx, c_idx]
                for item, weight in specs.items():
                    if item in raw_txt and weight not in raw_txt.replace(" ", ""):
                        cell = ws.cell(row=r_idx+1, column=c_idx+1)
                        cell.fill, cell.font = STYLE_SPEC["fill"], STYLE_SPEC["font"]
                        logs.append({"分頁": sn, "項目": item, "缺失": f"未標註規格 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

st.title("🛡️ 團膳稽核系統｜核心對位版")
up = st.file_uploader("📂 請上傳菜單 (系統將自動識別檔名並進行兩階段審核)", type=["xlsx"])

if up:
    logs, m_detected, data_out = robust_audit(up)
    if m_detected == "INVALID_FILENAME":
        st.error(f"❌ 第一階段失敗：檔名『{up.name}』不含『美食街』或『小學/幼兒園』關鍵字。")
    else:
        st.info(f"📁 第一階段通過：判定為【{m_detected}】格式。")
        if logs:
            st.error(f"🚩 第二階段發現 {len(logs)} 項內容缺失！")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件標註檔 (查看黑色格子)", data_out, f"退件_{up.name}")
        else:
            st.success("✅ 內容完整，未發現紅框缺失。")
