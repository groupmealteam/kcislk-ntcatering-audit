import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (維持 Alison 的原始標題)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：黑底白字 30 級 (專殺紅框空白) / 黃底紅字 (殺規格)
STYLE = {
    "BLACK_ERR": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_SPEC": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強迫 NaN 變成 MISSING，讓程式「看見」空白
        df_audit = df.fillna("MISSING")
        
        # 定位日期 Row (定錨點)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D-H 欄
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 關鍵修正：強制查核模式 ---
                # 偵測 A：熱量缺失 (4/28, 4/29 紅框)
                if "熱量" in label and (content == "MISSING" or content == "0" or content == ""):
                    cell.fill, cell.font = STYLE["BLACK_ERR"]["fill"], STYLE["BLACK_ERR"]["font"]
                    logs.append({"日期": date_val, "類別": "嚴重缺失", "原因": "⚠️ 熱量數據空白！"})

                # 偵測 B：菜名缺失但食材有填 (4/29 副菜紅框)
                # 邏輯：如果標籤是「副菜」，內容包含「+」號（通常是食材明細），代表沒寫菜名
                target_menu = ["主菜", "副菜", "套餐"]
                if any(t in label for t in target_menu):
                    if content == "MISSING" or "+" in content or "、" in content:
                        # 如果是空的，或者是誤把食材填進菜名欄
                        cell.fill, cell.font = STYLE["BLACK_ERR"]["fill"], STYLE["BLACK_ERR"]["font"]
                        logs.append({"日期": date_val, "類別": "嚴重缺失", "原因": f"❌ {label} 菜名漏填或填寫錯誤！"})

                # 偵測 C：原有規格稽核 (白帶魚、獅子頭等)
                check_list = {"白帶魚": "150g", "獅子頭": "60gX2", "漢堡排": "150g"}
                for item, spec in check_list.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                        logs.append({"日期": date_val, "類別": "規格不符", "原因": f"{item} 未標註 {spec}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳有缺失的菜單 Excel 進行驗證", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 抓到 {len(results)} 項缺失！(含紅框處空白/格式錯誤)")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔案", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完整，這次廠商沒逃掉！")
