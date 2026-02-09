import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 樣式定義 (依 Alison 規範) ---
# 缺失處：黑底白字
STYLE_ERR = {
    "fill": PatternFill("solid", fgColor="000000"), 
    "font": Font(name="微軟正黑體", size=18, color="FFFFFF", bold=True)
}

def alison_audit_engine(file):
    fname = file.name
    
    # --- 第一階段：檔名身分判斷 ---
    # 嚴格執行：美食街(C欄標籤)、教育學部(A欄標籤)
    if "美食街" in fname:
        mode, label_idx, data_indices = "美食街", 2, [3, 4, 5, 6, 7]
    elif any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode, label_idx, data_indices = "教育學部", 0, [1, 2, 3, 4, 5]
    elif "素食" in fname:
        mode, label_idx, data_indices = "素食專區", 2, [3, 4, 5, 6, 7]
    else:
        return None, "INVALID_FILENAME", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    # --- 第二階段：內容深度稽核 ---
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 清理所有偽裝空值 (0, 空格, nan)
        df_audit = df.astype(str).applymap(lambda x: "" if str(x).strip().lower() in ['nan', 'none', '0', '0.0', ''] else str(x).strip())
        max_rows, max_cols = df_audit.shape

        for r_idx in range(max_rows):
            if label_idx >= max_cols: continue
            
            # 標籤清理 (解決換行符號 \n 造成的偵測失敗)
            label = df_audit.iloc[r_idx, label_idx].replace('\n', '').strip()
            
            # 鎖定稽核目標：只要包含關鍵字就啟動
            targets = ["熱量", "主食", "主菜", "副菜", "湯品", "套餐"]
            if any(t in label for t in targets):
                for c_idx in data_indices:
                    if c_idx >= max_cols: continue
                    
                    # A. 放假判讀：若該欄(當日)數據全空，跳過不記缺失
                    col_data = "".join(df_audit.iloc[:, c_idx].tolist())
                    if len(col_data) == 0: continue
                    
                    # B. 週一特例：週一早上無早餐視為正常
                    if "早餐" in label and "一" in sn: continue

                    content = df_audit.iloc[r_idx, c_idx]
                    cell = ws.cell(row=r_idx+1, column=c_idx+1)

                    # C. 核心判斷：專抓 4/29 式的「菜名黑洞」
                    # 邏輯：菜名格為空，且下方(明細行)不為空
                    if content == "":
                        try:
                            detail = df_audit.iloc[r_idx+1, c_idx].strip()
                            if len(detail) > 0: # 下方明細有字，這格就是漏填
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填菜名"
                                logs.append({"分頁": sn, "項目": label, "原因": "內容不完整 (有食材無菜名)"})
                        except: pass
                    
                    # D. 熱量稽核
                    if "熱量" in label and content == "":
                        cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                        logs.append({"分頁": sn, "項目": label, "原因": "熱量數值缺失"})

    # 輸出標註後的檔案
    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- Streamlit 介面 (維持 Alison 要求之簡潔) ---
st.set_page_config(page_title="團膳稽核系統 - Alison", layout="wide")
st.title("🛡️ 團膳稽核系統 - 小學部 / 幼兒園 (細項模式)")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案 (系統將執行兩階段自動審核)", type=["xlsx"])

if up:
    with st.spinner("正在執行 Alison 規範稽核中..."):
        logs, m_detected, data_out = alison_audit_engine(up)
    
    if m_detected == "INVALID_FILENAME":
        st.error(f"❌ 第一階段失敗：檔名『{up.name}』未包含美食街/小學/幼兒園/素食等關鍵字。")
    else:
        st.info(f"✅ 第一階段通過：判定為【{m_detected}】模式。")
        if logs:
            st.error(f"🚩 第二階段發現 {len(logs)} 項不完整缺失（包含 4/29 式黑洞）。")
            st.table(pd.DataFrame(logs))
            st.download_button(
                label="📥 下載 Alison 專屬退件標註檔",
                data=data_out,
                file_name=f"退件_{up.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("🎉 內容完整！符合 Alison 稽核標準與放假判斷邏輯。")
