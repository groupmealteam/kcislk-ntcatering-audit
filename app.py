import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# --- 1. 核心稽核字庫 (依據妳提供的 5 月菜單退件邏輯) ---
TAGS_FRIED = ["炸", "酥", "裹粉", "爆", "脆", "可樂餅", "◎", "魚排", "雞排", "春捲"] 
TAGS_PROCESSED = ["丸", "排", "素羊", "素火腿", "素肉", "獅子頭", "豆包", "炸豆腐", "△", "★", "成品", "肉燥", "捲", "甜不辣", "肉類加工"]
TAGS_FREQUENT = ["豆芽", "銀芽", "芽菜"]

def process_excel(uploaded_file):
    # 讀取 Excel 原始檔案 (使用 data_only 確保抓到數值而非公式)
    wb = load_workbook(uploaded_file, data_only=True)
    output_logs = []
    
    # --- 設定標記樣式 ---
    # 底色：亮紅色
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    # 字體：32 級、粗體、白色
    big_font = Font(color="FFFFFF", bold=True, size=32) 
    # 對齊：居中、自動換行
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        fried_count = 0
        veg_count = 0
        
        # 尋找日期列 (抓天數來判斷短週)
        date_row_idx = None
        days_in_week = 0
        for r in range(1, 21):
            row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 12)]
            if any("日期" in v or "Date" in v for v in row_vals):
                date_row_idx = r
                break
        
        if date_row_idx:
            for c in range(3, 10):
                v = str(ws.cell(row=date_row_idx, column=c).value)
                if "202" in v or "/" in v: days_in_week += 1

            # 遍歷所有供餐儲存格
            for c_idx in range(3, 10):
                for r_idx in range(1, ws.max_row + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    val = str(cell.value).strip().replace('\n', '') if cell.value else ""
                    if len(val) < 1 or val == "None": continue

                    issue = None
                    # A. 炸物 (短週/一般週嚴格限1次)
                    if any(f in val for f in TAGS_FRIED):
                        fried_count += 1
                        limit = 1 if days_count < 5 else 1 
                        if fried_count > limit:
                            issue = f"🚩炸物超標"

                    # B. 加工品/素料規格 (依增補協議：強制要求標註數量或克數)
                    elif any(p in val for p in TAGS_PROCESSED):
                        if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", val):
                            issue = "⚠️缺規格"

                    # C. 豆芽重複 (郁恩專業意見)
                    elif any(v in val for v in TAGS_FREQUENT):
                        veg_count += 1
                        if veg_count > 1:
                            issue = "❌豆芽重複"

                    # 發現問題：標色、放大字體至 32、改白字
                    if issue:
                        cell.fill = red_fill
                        cell.font = big_font
                        cell.alignment = center_align
                        # 將原因加在文字前面方便廠商看
                        cell.value = f"{issue}\n{cell.value}"
                        output_logs.append({"分頁": sheet_name, "項目": val, "問題": issue})

    virtual_wb = BytesIO()
    wb.save(virtual_wb)
    virtual_wb.seek(0)
    return virtual_wb, output_logs

# --- Streamlit 介面 ---
st.set_page_config(page_title="康橋膳食稽核系統", layout="wide")
st.title("🛡️ 康橋林口：膳食稽核系統 (字體 32 級標記版)")
st.warning("本系統依據《114增補協議》執行。違規項目將在 Excel 中以紅底白字 32 級大字標記。")

f = st.file_uploader("📂 請上傳新北食品 Excel 菜單", type=["xlsx"])

if f:
    with st.spinner('正在地毯式標記違規格子...'):
        processed_file, logs = process_excel(f)
        
        if logs:
            st.error(f"🚨 偵測到 {len(logs)} 項違規！請下載標記版檔案直接退回廠商。")
            st.table(pd.DataFrame(logs))
            st.download_button(
                label="📥 下載【32級大字標記版】退件 Excel",
                data=processed_file,
                file_name="菜單退件建議_請修正紅色標記處.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("✅ 檢查完畢，未發現違規項目。")
