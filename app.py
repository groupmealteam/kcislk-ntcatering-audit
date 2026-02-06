import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="NTCatering - Menu Audit System", layout="wide")

# 設定違規標色
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

# 2. 新北食品合約規格
CONTRACT_SPECS = {
    "現撈小卷": "80|100",
    "無刺白帶魚": "120|150",
    "手作獅子頭": "60",
    "手作漢堡排": "150",
    "手作烤肉串": "80"
}

def audit_logic(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except Exception:
        return ["❌ 檔案格式損壞，請上傳正確的 Excel 檔。"], None

    results = []
    output = BytesIO()
    is_menu_valid = False  # 驗證是否為真實菜單

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 尋找關鍵欄位：日期(M/D) 與 B 欄是否含「主食/副菜」
        date_row = next((i for i, row in df.iterrows() if any(re.search(r"\d{1,2}/\d{1,2}", str(c)) for c in row)), None)
        target_rows = [i for i, row in df.iterrows() if any(k in str(row[1]) for k in ["主食", "副菜", "主菜", "套餐"])]
        
        # 如果找不到日期或主食，這張表就不是菜單
        if date_row is None or len(target_rows) == 0:
            continue
        
        is_menu_valid = True # 只要有一張分頁符合，就視為菜單

        for col in range(2, len(df.columns)):
            date_val = str(df.iloc[date_row, col])
            if not re.search(r"\d{1,2}/\d{1,2}", date_val): continue
            
            day_processed = 0 
            day_fried = 0     
            
            for r_idx in target_rows:
                cell_val = str(df.iloc[r_idx, col]).strip()
                if not cell_val: continue

                # A. 檢核 NTCatering 合約克重
                for item, spec in CONTRACT_SPECS.items():
                    if item in cell_val and not re.search(spec, cell_val):
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": cell_val, "問題": f"⚠️ 規格錯誤：須標註 {spec}g"})

                # B. 檢核法規標示
                if "△" in cell_val: day_processed += 1
                if "◎" in cell_val: day_fried += 1

            if day_processed > 1:
                results.append({"分頁": sheet_name, "日期": date_val, "問題": f"🚫 原則五：加工品(△)超過 1 項"})
            if day_fried > 1:
                results.append({"分頁": sheet_name, "日期": date_val, "問題": f"🚫 原則七：油炸(◎)超過 1 次"})

    if not is_menu_valid:
        return ["❌ 偵測失敗：上傳檔案不含日期或主食欄位，請確認是否為正確菜單格式。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 網頁介面 ---
st.title("🛡️ NTCatering (新北食品) 菜單自主稽核系統")
st.warning("⚠️ 注意：系統僅接受包含『日期(M/D)』與『主食/副菜』欄位的正式菜單檔案。")

up = st.file_uploader("👉 上傳週菜單 Excel (.xlsx)", type=["xlsx"])

if up:
    logs, final_file = audit_logic(up)
    
    # 判斷是否為「報錯訊息」而非「審核結果」
    if logs and isinstance(logs[0], str) and logs[0].startswith("❌"):
        st.error(logs[0])
    elif logs:
        st.error(f"🚩 偵測到 {len(logs)} 處異常項目。")
        st.download_button("📥 下載標註檔 (請修正標紅處)", final_file, f"NTCatering_Check_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        # 只有在通過驗證後才顯示成功
        st.success("🎉 審核完成！該份正式菜單完全符合合約與法規。")
