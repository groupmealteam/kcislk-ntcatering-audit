import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定 (標題改為 NTCatering)
st.set_page_config(page_title="NTCatering - Menu Audit System", layout="wide")

# 設定違規標色
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

# 2. 新北食品(NTCatering) 合約規格規範
CONTRACT_SPECS = {
    "現撈小卷": "80|100",
    "無刺白帶魚": "120|150",
    "手作獅子頭": "60",
    "手作漢堡排": "150",
    "手作烤肉串": "80"
}

def audit_logic(file):
    wb = load_workbook(file)
    all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    results = []
    output = BytesIO()

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位日期與主副食行
        date_row = next((i for i, row in df.iterrows() if any(re.search(r"\d{1,2}/\d{1,2}", str(c)) for c in row)), None)
        target_rows = [i for i, row in df.iterrows() if any(k in str(row[1]) for k in ["主食", "副菜", "主菜"])]
        
        if date_row is None: continue

        for col in range(2, len(df.columns)):
            date_val = str(df.iloc[date_row, col])
            if not re.search(r"\d{1,2}/\d{1,2}", date_val): continue
            
            day_processed = 0 
            day_fried = 0     
            
            for r_idx in target_rows:
                cell_val = str(df.iloc[r_idx, col]).strip()
                if not cell_val: continue

                # A. 檢核 NTCatering 合約克重 (原則八)
                for item, spec in CONTRACT_SPECS.items():
                    if item in cell_val and not re.search(spec, cell_val):
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"日期": date_val, "項目": cell_val, "問題": f"⚠️ 規格不符：合約要求須標註 {spec}g"})

                # B. 檢核法規標示 (原則五、七)
                if "△" in cell_val: day_processed += 1
                if "◎" in cell_val: day_fried += 1

            # C. 檢核數量限制
            if day_processed > 1:
                results.append({"日期": date_val, "問題": f"🚫 違反原則五：加工食品(△)超過單日限制"})
            if day_fried > 1:
                results.append({"日期": date_val, "問題": f"🚫 違反原則七：油炸料理(◎)超過單日限制"})

    wb.save(output)
    return results, output.getvalue()

# --- 網頁介面佈局 ---
st.title("🛡️ NTCatering (新北食品) 菜單自主稽核系統")
st.markdown("---")
st.info("💡 請上傳您的週菜單 Excel，系統將根據《林口康橋菜單審閱原則》自動校閱合約規格與法規標示。")

up = st.file_uploader("👉 上傳菜單 Excel (.xlsx)", type=["xlsx"])

if up:
    with st.spinner("系統分析中..."):
        logs, final_file = audit_logic(up)
        if logs:
            st.error(f"🚩 偵測到 {len(logs)} 處異常項目。")
            st.download_button(
                label="📥 下載標註完成之 Excel (請修正標紅處)",
                data=final_file,
                file_name=f"NTCatering_Audit_{up.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 審核完成！該週菜單符合 NTCatering 合約與法規規範。")
