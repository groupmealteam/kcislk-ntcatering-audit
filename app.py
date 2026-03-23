import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ==========================================
# 1. 定義【合約稽核紅線】關鍵字
# ==========================================
# 炸物關鍵字 (含午餐與點心)
TAGS_FRIED = ["炸", "酥", "裹粉", "薯條", "雞塊", "卡拉", "可樂餅", "春捲"]
# 加工品關鍵字 (需檢查規格)
TAGS_PROCESSED = ["丸", "排", "火腿", "成品", "獅子頭", "肉燥", "捲", "鑫鑫腸", "熱狗"]
# 強烈調味 (一週不重複)
TAGS_SEASONING = ["沙茶", "咖哩", "腐乳", "三杯", "麻婆", "糖醋"]
# 排除主菜屬性的湯品 (合約規範)
TAGS_SOUP = ["湯", "羹", "粥", "麵湯"]

def audit_contract_logic(df, sheet_name):
    logs = []
    fried_count = 0
    seasoning_tracker = {}
    
    # --- 步驟 A：判斷本週供餐天數 (解決跨月短週問題) ---
    days_in_week = 0
    date_row = None
    # 尋找日期列 (通常在前15列)
    for i in range(min(15, len(df))):
        if any(k in str(df.iloc[i, 2]) for k in ["日期", "Date"]):
            date_row = i
            break
    
    if date_row is not None:
        for col in range(2, 7): # 檢查 C-G 欄有幾個日期
            if col < len(df.columns) and "202" in str(df.iloc[date_row, col]):
                days_in_week += 1

    # --- 步驟 B：逐格掃描 (合併稽核午餐與點心) ---
    if date_row is None: return [] # 找不到日期則跳過該頁
    
    for col in range(2, 7): # 掃描週一到週五
        if col >= len(df.columns): break
        date_label = str(df.iloc[date_row, col]).split(" ")[0]
        
        for r_idx in range(len(df)):
            raw_val = str(df.iloc[r_idx, col]).strip()
            if raw_val in ["nan", "", "None"]: continue
            
            # 1. 炸物計次 (含午餐與點心，執行短週限1次原則)
            if any(f in raw_val for f in TAGS_FRIED):
                fried_count += 1
                # 妳的主張：短週(天數<5)炸物限1次
                limit = 1 if days_in_week < 5 else 1 
                if fried_count > limit:
                    logs.append({"日期": date_label, "項目": raw_val, "原因": f"🚩炸物超標(當週累計{fried_count}次)"})

            # 2. 加工品規格強制令 (依增補協議：沒標數量/克數就退件)
            if any(p in raw_val for p in TAGS_PROCESSED):
                if not re.search(r"[xX*×]\d|克|g|G|公克", raw_val):
                    logs.append({"日期": date_label, "項目": raw_val, "原因": "⚠️規格不詳(依合約需標數量或克數)"})

            # 3. 調味重複性檢查
            for s in TAGS_SEASONING:
                if s in raw_val:
                    if s in seasoning_tracker and seasoning_tracker[s] != date_label:
                        logs.append({"日期": date_label, "項目": raw_val, "原因": f"⚠️調味重複({s})"})
                    seasoning_tracker[s] = date_label

            # 4. 適口性警示 (幼兒園專用)
            if "幼" in sheet_name:
                if "粗米粉" in raw_val or "大丸子" in raw_val:
                    logs.append({"日期": date_label, "項目": raw_val, "原因": "💡建議調整食材型態(適口性)"})

    return logs

# --- Streamlit 介面 ---
st.set_page_config(page_title="康橋林口膳食稽核系統", layout="wide")
st.title("🛡️ 康橋林口校區：合約執行稽核系統")
st.info("本系統依據《114學年增補協議》執行。稽核紅線：短週炸物限1次、加工品強制標註規格、點心合併計次。")

uploaded_file = st.file_uploader("📂 上傳新北食品菜單 (Excel)", type=["xlsx"])

if uploaded_file:
    with st.spinner('正在比對合約規格...'):
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        audit_results = []
        
        for sn, df in all_sheets.items():
            errors = audit_contract_logic(df, sn)
            for e in errors:
                e['分頁'] = sn
                audit_results.append(e)
        
        if audit_results:
            st.error(f"🚨 偵測到 {len(audit_results)} 項不符合約規範：")
            st.table(pd.DataFrame(audit_results))
            st.warning("請依據上述結果退回廠商修正。")
        else:
            st.success("✅ 檢查完畢！本週菜單(含午餐與點心)均符合合約炸物限制與規格標註。")
