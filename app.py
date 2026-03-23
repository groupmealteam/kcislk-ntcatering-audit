import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ==========================================
# 1. 設定【合約紅線】關鍵字 (這裡妳以後可以自己加)
# ==========================================
TAGS_FRIED = ["炸", "酥", "裹粉", "薯條", "雞塊", "卡拉", "可樂餅"]
TAGS_PROCESSED = ["丸", "排", "火腿", "成品", "獅子頭", "肉燥", "捲", "鑫鑫腸"]
TAGS_HEAVY_SEASONING = ["沙茶", "咖哩", "腐乳", "三杯", "麻婆", "糖醋"]
# 幼兒園適口性警示
TAGS_KINDERGARTEN_LIMIT = ["粗米粉", "大丸子", "硬糖"]

# 顏色定義
STYLE_RED = {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(color="FFFFFF", bold=True)} # 違規
STYLE_BLACK = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(color="FFFFFF", bold=True)} # 缺失

def audit_contract_logic(df, sheet_name):
    logs = []
    fried_count = 0
    seasoning_used = []
    
    # 判斷這張表有幾天 (用來抓跨月短週)
    days_in_week = 0
    date_row = None
    for i in range(min(15, len(df))):
        if any(k in str(df.iloc[i, 2]) for k in ["日期", "Date"]):
            date_row = i
            break
    
    if date_row is not None:
        for col in range(2, 7): # C 到 G 欄
            if col < len(df.columns) and "202" in str(df.iloc[date_row, col]):
                days_in_week += 1

    # 逐格掃描 (包含午餐與點心區)
    for col in range(2, 7): 
        if col >= len(df.columns): break
        date_label = str(df.iloc[date_row, col]).split(" ")[0] if date_row is not None else "未知日期"
        
        for r_idx in range(len(df)):
            raw_val = str(df.iloc[r_idx, col]).strip()
            if raw_val in ["nan", "", "None"]: continue
            
            # --- A. 炸物累計 (點心+午餐合併計算) ---
            if any(word in raw_val for word in TAGS_FRIED):
                fried_count += 1
                # 執行妳的主張：短週(天數<5) 炸物只能 1 次
                limit = 1 if days_in_week < 5 else 1 # 這裡如果妳想放寬完整週可以改 2
                if fried_count > limit:
                    logs.append({"日期": date_label, "項目": raw_val, "原因": f"炸物超標(當週累計{fried_count}次)"})

            # --- B. 加工品【規格強制令】 (依增補協議) ---
            if any(word in raw_val for word in TAGS_PROCESSED):
                # 檢查有沒有 X數字 或 克數，沒寫就退件
                if not re.search(r"[xX*×]\d|克|g|G", raw_val):
                    logs.append({"日期": date_label, "項目": raw_val, "原因": "未標註規格(依增補協議需標數量/克數)"})

            # --- C. 調味重複性檢查 ---
            for s in TAGS_HEAVY_SEASONING:
                if s in raw_val:
                    if s in seasoning_used:
                        logs.append({"日期": date_label, "項目": raw_val, "原因": f"調味重複({s})"})
                    else:
                        seasoning_used.append(s)
            
            # --- D. 幼兒園適口性 (如果是幼兒園分頁) ---
            if "幼" in sheet_name or "小" in sheet_name:
                if any(word in raw_val for word in TAGS_KINDERGARTEN_LIMIT):
                    logs.append({"日期": date_label, "項目": raw_val, "原因": "不符適口性建議(請調整食材型態)"})

    return logs

# --- 介面呈現 ---
st.set_page_config(page_title="康橋林口膳食稽核", layout="wide")
st.title("🛡️ 康橋林口校區：合約防禦稽核系統")
st.subheader("適用對象：新北食品 (依據 114 學年增補協議)")

file = st.file_uploader("📂 請上傳菜單 Excel 檔案", type=["xlsx"])

if file:
    with st.spinner('合約比對中...'):
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
        final_errors = []
        
        for sn, df in all_sheets.items():
            results = audit_contract_logic(df, sn)
            for r in results:
                r['分頁'] = sn
                final_errors.append(r)
        
        if final_errors:
            st.error(f"🚩 發現 {len(final_errors)} 處不符合約或審核標準：")
            st.table(pd.DataFrame(final_errors))
            st.warning("💡 請依據上述理由退回新北食品修正。")
        else:
            st.success("✅ 檢查完畢！本週菜單符合炸物限制、加工品規格及調味多樣性原則。")
