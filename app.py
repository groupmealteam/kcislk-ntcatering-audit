import streamlit as st
import pd as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 重新對齊【郁恩邏輯】的高強度字庫 (增加符號與變體)
TAGS_FRIED = ["炸", "酥", "裹粉", "爆", "脆", "可樂餅", "◎", "魚排", "雞排"] 
TAGS_PROCESSED = ["丸", "排", "素羊", "素火腿", "素肉", "獅子頭", "豆包", "炸豆腐", "△", "★", "成品", "肉燥", "捲"]
TAGS_FREQUENT = ["豆芽", "銀芽", "芽菜"]

def process_excel(uploaded_file):
    # 讀取 Excel 原始檔案
    wb = load_workbook(uploaded_file, data_only=True) # data_only 確保讀到的是值不是公式
    output_logs = []
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        fried_count = 0
        veg_count = 0
        
        # 尋找日期列 (擴大搜尋到前 20 列)
        date_row_idx = None
        days_in_week = 0
        for r in range(1, 21):
            row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 10)]
            if any("日期" in v or "Date" in v for v in row_vals):
                date_row_idx = r
                break
        
        if date_row_idx:
            # 計算當週天數
            for c in range(3, 8):
                v = str(ws.cell(row=date_row_idx, column=c).value)
                if "202" in v or "/" in v: days_in_week += 1

            # 遍歷整張表所有格子 (C欄到G欄 = 3到7)
            for c_idx in range(3, 8):
                for r_idx in range(1, ws.max_row + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    # 強制轉字串並過濾掉換行符號
                    val = str(cell.value).strip().replace('\n', '') if cell.value else ""
                    if len(val) < 2 or val == "None": continue

                    issue = None
                    # A. 炸物 (包含◎與關鍵字)
                    if any(f in val for f in TAGS_FRIED):
                        fried_count += 1
                        limit = 1 if days_in_week < 5 else 1
                        if fried_count > limit:
                            issue = f"🚩炸物超標(第{fried_count}次)"

                    # B. 加工品規格 (包含△與關鍵字)
                    elif any(p in val for p in TAGS_PROCESSED):
                        if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", val):
                            issue = "⚠️加工品缺規格"

                    # C. 豆芽重複
                    elif any(v in val for v in TAGS_FREQUENT):
                        veg_count += 1
                        if veg_count > 1:
                            issue = "❌豆芽重複"

                    # 只要有問題，就塗紅並紀錄
                    if issue:
                        cell.fill = red_fill
                        cell.font = white_font
                        output_logs.append({
                            "分頁": sheet_name,
                            "項目": val,
                            "問題": issue
                        })

    virtual_wb = BytesIO()
    wb.save(virtual_wb)
    virtual_wb.seek(0)
    return virtual_wb, output_logs

# --- Streamlit 介面 ---
st.set_page_config(page_title="康橋膳食稽核系統", layout="wide")
st.title("🛡️ 康橋林口：專業稽核標記系統")
st.subheader("⚠️ 修正版：強化掃描所有隱藏字元與符號")

f = st.file_uploader("📂 請上傳新北食品 Excel 菜單", type=["xlsx"])

if f:
    with st.spinner('正在地毯式掃描並塗紅違規格子...'):
        processed_file, logs = process_excel(f)
        
        if logs:
            st.error(f"🚨 偵測到 {len(logs)} 項問題！請下載下方標記版檔案。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載紅字標記版菜單", processed_file, "審核標記版.xlsx")
        else:
            st.success("✅ 檢查完畢，未發現違規。若仍有落差請檢查檔案分頁。")
