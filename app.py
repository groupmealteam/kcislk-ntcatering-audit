import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 頁面標題 (嚴格鎖定)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=20, color="FF0000", bold=True)}
}

def audit_process(file):
    fname = file.name
    # 第一步：檔名判讀
    if "美食街" in fname:
        mode = "美食街"
        label_col = 2  # C欄
        data_cols = range(3, 8) # D-H 欄
    elif "小學" in fname or "幼兒園" in fname:
        mode = "教育學部"
        label_col = 0  # A欄
        data_cols = range(1, 6) # B-F 欄
    else:
        mode = "未知格式"
        return None, mode

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制標記空值
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0', ' ', ''], 'MISSING_ERR')
        
        # 定錨：找日期列
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[label_col])), None)
        if d_row is None: continue

        for col in data_cols:
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, label_col]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # 第二步：條件判讀審核 (針對之前的紅框缺失)
                
                # A. 熱量黑洞：只要是熱量格卻是空的，必噴黑
                if "熱量" in label and content == "MISSING_ERR":
                    cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                    logs.append({"分頁": sn, "日期": date_val, "缺失": f"{label} 漏填數字"})

                # B. 菜名黑洞：副菜/主菜格子空，但下方有食材資訊 (4/29 紅框死穴)
                if any(t in label for t in ["主菜", "副菜", "套餐", "主食"]) and content == "MISSING_ERR":
                    try:
                        next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                        if next_val != "MISSING_ERR": # 下一列有食材
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"分頁": sn, "日期": date_val, "缺失": f"{label} 漏填菜名 (但有填食材)"})
                    except: pass

                # C. 規格嚴審：白帶魚 150g, 漢堡排 150g, 獅子頭 60gX2
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"分頁": sn, "日期": date_val, "缺失": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue(), mode

st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
up = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])

if up:
    logs, data, detected_mode = audit_process(up)
    
    if detected_mode == "未知格式":
        st.warning(f"⚠️ 檔名「{up.name}」無法辨識部別，請確認檔名包含『美食街』或『小學/幼兒園』。")
    else:
        st.info(f"📁 系統判定：**{detected_mode}** 菜單格式")
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項缺失（包含 4/28-4/29 空白處）。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
        else:
            st.success("✅ 檢查完畢，未發現缺失。")
