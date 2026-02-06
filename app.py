import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
FONT_NAME = "微軟正黑體"
STYLE = {
    "MISSING": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF", bold=True)}, # 黑底：漏填/刪除
    "DATA_FAIL": {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name=FONT_NAME, size=30, color="FFFFFF")},       # 紅底：不符標準
    "CONTRACT": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name=FONT_NAME, size=30, color="FF0000", bold=True)}, # 黃底：規格不符
    "SPICY": {"fill": PatternFill("solid", fgColor="C6EFCE"), "font": Font(name=FONT_NAME, size=30, color="000000")}        # 綠底：禁辣違規
}

# 規格鎖死 (依據 SE1140803 增補協議書)
CONTRACT_SPECS = {"獅子頭": "60gX2", "漢堡排": "150g", "鯰魚片": "120g", "白蝦": "X3", "無刺白帶魚": "150g", "砂鍋魚丁": "250g"}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        df_audit = df.fillna("!!!MISSING!!!") # 這次直接用驚嘆號標註，絕對不漏
        
        # 搜尋關鍵字定位 row (不再用死板的數字)
        date_row_idx = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]) or "Date" in str(row[2]):
                date_row_idx = i
                break
        
        if date_row_idx is None: continue

        for col in range(3, 8): # 檢查 D 到 H 欄
            date_val = str(df_audit.iloc[date_row_idx, col]).split("\n")[0]
            day_text = str(df_audit.iloc[date_row_idx+1, col])

            # --- 1. 結構完整性 (主食/主菜/副菜/湯品) ---
            # 只要在 C 欄標籤對應的右側是空的，就是壞了
            for r_idx in range(date_row_idx + 2, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2])
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # A. 抓包：結構缺失 (妳刪掉的地方)
                target_labels = ["主食", "主菜", "副菜", "青菜", "湯品", "熱量"]
                if any(tl in label for tl in target_labels) and content in ["!!!MISSING!!!", "", "nan", "0"]:
                    cell.fill, cell.font = STYLE["MISSING"]["fill"], STYLE["MISSING"]["font"]
                    logs.append({"分頁": sn, "日期": date_val, "項目": "結構重大缺失", "原因": f"❌ {label} 被刪除或未填"})

                # B. 抓包：合約規格 (妳那份 4/2 寫無刺白帶魚，沒寫 150g 就退件)
                for item, spec in CONTRACT_SPECS.items():
                    if item in content and spec not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["CONTRACT"]["fill"], STYLE["CONTRACT"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "合約規格違規", "原因": f"{item}需標註 {spec}"})

                # C. 抓包：禁辣日 (週一、二、四)
                if any(d in day_text for d in ["(一)", "(二)", "(四)"]):
                    if any(x in content for x in ["🌶️", "●", "辣", "椒", "麻", "沙茶"]):
                        cell.fill, cell.font = STYLE["SPICY"]["fill"], STYLE["SPICY"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "項目": "禁辣違規", "原因": f"禁辣日出現: {content}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- 介面 ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison | 依據：114學年增補協議 & 審閱原則修訂2")

up = st.file_uploader("📂 請上傳美食街菜單 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if results:
        st.error(f"🚩 稽核發現 {len(results)} 處違規與缺失！")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔 (檢視黑底/黃底)", data, f"退件_{up.name}")
    else:
        st.success("🎉 通過稽核。")
