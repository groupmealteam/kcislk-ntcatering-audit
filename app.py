import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁配置
st.set_page_config(page_title="團膳區(新北食品) 多功能稽核系統", layout="wide")

# --- 樣式設定：黑底白字 30 級 (專殺空白) / 黃底紅字 (殺規格) ---
STYLE = {
    "BLACK": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

# 2. 側邊欄：選擇審核模式
st.sidebar.title("🔍 審核模式切換")
mode = st.sidebar.selectbox(
    "請選擇菜單類別：",
    ["小學部 / 幼兒園 (細項模式)", "美食街 (早午晚大雜燴模式)", "素食菜單"]
)

def audit_process(file, mode):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 強制字串化，確保 NaN 變成可辨識的標籤
        df_audit = df.astype(str).replace(['nan', 'None', 'NaN', '0', '0.0'], 'MISSING')
        
        # 定位日期 Row
        d_row = next((i for i, r in df_audit.iterrows() if "日期" in str(r[0]) or "日期" in str(r[2])), None)
        if d_row is None: continue

        # 根據模式設定掃描範圍
        if "美食街" in mode:
            cols = range(3, 8)  # 美食街通常是 D-H 欄
            target_tags = ["熱量", "主菜", "副菜", "套餐", "主食"]
        else:
            cols = range(1, 10) # 小學/幼兒園通常橫跨 A 欄開始
            target_tags = ["熱量", "主食", "主菜", "副菜", "下午點心"]

        for col in cols:
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 0 if "美食街" not in mode else 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心邏輯：黑洞偵測 ---
                if any(tag in label for tag in target_tags):
                    # 4/28-4/30 專用補丁：如果內容是空的 MISSING，直接噴黑
                    if content == "MISSING" or content == "":
                        is_fail = False
                        if "熱量" in label:
                            is_fail = True
                        else:
                            # 檢查下一行有沒有「食材明細」，有明細沒菜名就是漏填
                            try:
                                next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                                if next_val != "MISSING": is_fail = True
                            except: pass
                        
                        if is_fail:
                            cell.fill, cell.font = STYLE["BLACK"]["fill"], STYLE["BLACK"]["font"]
                            logs.append({"日期": date_val, "類別": label, "原因": "❌ 內容漏填！"})

                # --- 核心邏輯：規格稽核 ---
                specs = {"白帶魚": "150g", "漢堡排": "150g", "獅子頭": "60gX2"}
                for item, weight in specs.items():
                    if item in content and weight not in content.replace(" ", ""):
                        cell.fill, cell.font = STYLE["YELLOW"]["fill"], STYLE["YELLOW"]["font"]
                        logs.append({"日期": date_val, "類別": "規格缺失", "原因": f"{item} 未標註 {weight}"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# 3. 主頁面介面
st.title(f"🛡️ 團膳稽核系統 - {mode}")
st.caption("製作者：Alison | 專門處理新北康橋多格式菜單")

up = st.file_uploader(f"📂 請上傳【{mode}】的 Excel 檔案", type=["xlsx"])

if up:
    results, data = audit_process(up, mode)
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項缺失（包含 4/28-4/29 紅框）。")
        st.table(pd.DataFrame(results))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("✅ 結構完美，這次廠商沒逃過妳的法眼！")
