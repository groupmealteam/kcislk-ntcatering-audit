import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (標題與註解完全鎖死)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式：黑底白字 30 級 (專抓 4/28-4/29 這種挖空的垃圾行為)
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
        # 修正核心 BUG：強制將所有空值變為 "VOID"，不准程式裝瞎跳過
        df_audit = df.fillna("VOID")
        
        # 尋找日期列 (定錨點)
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[2])), None)
        if d_row is None: continue

        for col in range(3, 8): # D 到 H 欄 (週一至週五)
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 抓包 A：熱量黑洞 (4/28, 4/29 晚餐熱量)
                if "熱量" in label:
                    if content in ["VOID", "", "nan", "0"]:
                        cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                        logs.append({"日期": date_val, "缺失": "數據缺失", "原因": "⚠️ 熱量欄位被挖空！"})

                # 抓包 B：幽靈菜單 (4/29 副菜：菜名空白但下格有食材)
                if label in ["主菜", "副菜", "青菜", "湯品"]:
                    if content == "VOID":
                        # 往下看一格，如果食材明細有字，這格就是漏填！
                        try:
                            detail_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if detail_val != "VOID":
                                cell.fill, cell.font = STYLE["BLACK_CRITICAL"]["fill"], STYLE["BLACK_CRITICAL"]["font"]
                                logs.append({"日期": date_val, "缺失": "結構缺失", "原因": f"❌ {label} 漏填菜名！"})
                        except: pass

                # 抓包 C：合約文字遊戲 (白帶魚 150g / 獅子頭 60gX2)
                specs = {"白帶魚": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW_CONTRACT"]["fill"], STYLE["YELLOW_CONTRACT"]["font"]
                        logs.append({"日期": date_val, "缺失": "規格不符", "原因": f"{item} 需標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單檔案進行「紅框」壓力測試", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 成功抓到 {len(results)} 項嚴重缺失！")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件建議檔", data, f"退件_{up.name}")
