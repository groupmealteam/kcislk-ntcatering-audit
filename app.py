import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# --- 核心稽核字庫 ---
TAGS_FRIED = ["炸", "酥", "裹粉", "爆", "脆", "可樂餅", "◎", "魚排", "雞排", "春捲"] 
TAGS_PROCESSED = ["丸", "排", "素羊", "素火腿", "素肉", "獅子頭", "豆包", "炸豆腐", "△", "★", "成品", "肉燥", "捲", "甜不辣"]
TAGS_FREQUENT = ["豆芽", "銀芽", "芽菜"]

def process_excel(uploaded_file):
    wb = load_workbook(uploaded_file, data_only=True)
    output_logs = []
    
    # 設定震撼的紅牌樣式：字體 32, 粗體, 白字, 紅底
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    big_white_font = Font(color="FFFFFF", bold=True, size=16) 
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        fried_count = 0
        veg_count = 0
        
        # 抓取天數判定
        date_row_idx = None
        days_in_week = 0
        for r in range(1, 21):
            row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 10)]
            if any("日期" in v or "Date" in v for v in row_vals):
                date_row_idx = r
                break
        
        if date_row_idx:
            for c in range(3, 8):
                v = str(ws.cell(row=date_row_idx, column=c).value)
                if "202" in v or "/" in v: days_in_week += 1

            for c_idx in range(3, 8):
                for r_idx in range(1, ws.max_row + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    val = str(cell.value).strip().replace('\n', '') if cell.value else ""
                    if len(val) < 2 or val == "None": continue

                    issue = None
                    if any(f in val for f in TAGS_FRIED):
                        fried_count += 1
                        limit = 1 if days_in_week < 5 else 1 
                        if fried_count > limit: issue = f"🚩炸物超標({fried_count}次)"
                    elif any(p in val for p in TAGS_PROCESSED):
                        if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", val):
                            issue = "⚠️加工品缺規格"
                    elif any(v in val for v in TAGS_FREQUENT):
                        veg_count += 1
                        if veg_count > 1: issue = "❌豆芽重複"

                    if issue:
                        cell.fill = red_fill
                        cell.font = big_white_font
                        cell.alignment = center_align
                        output_logs.append({"分頁": sheet_name, "項目": val, "問題": issue})

    virtual_wb = BytesIO()
    wb.save(virtual_wb)
    virtual_wb.seek(0)
    return virtual_wb, output_logs

# --- 介面 ---
st.set_page_config(page_title="康橋膳食稽核系統", layout="wide")
st.title("🛡️ 康橋林口：膳食稽核標記系統 (高對比版)")
st.info("本系統依據《114增補協議》與營養師專業邏輯執行，標記字體已放大至 16 級。")

f = st.file_uploader("📂 上傳新北食品 Excel 菜單", type=["xlsx"])
if f:
    with st.spinner('標記中...'):
        processed_file, logs = process_excel(f)
        if logs:
            st.error(f"🚨 偵測到 {len(logs)} 項違規！請下載標記檔案退件。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載【字體放大標記版】退件 Excel", processed_file, "審核退件標記版.xlsx")
        else:
            st.success("✅ 檢查完畢，未發現違規項目。")
