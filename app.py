import streamlit as st
import pandas as pd
import re

# --- 1. 定義【合約執行】紅線標籤 (妳的主張數位化) ---
TAGS_FRIED = ["炸", "酥", "裹粉", "薯條", "雞塊", "卡拉", "春捲", "排骨酥", "可樂餅"]
TAGS_PROCESSED = ["丸", "排", "火腿", "成品", "獅子頭", "肉燥", "捲", "鑫鑫腸", "甜不辣"]
TAGS_SEASONING = ["沙茶", "咖哩", "腐乳", "三杯", "麻婆", "糖醋"]

def audit_menu(df, sheet_name):
    logs = []
    fried_count = 0
    used_seasoning = set()
    
    # --- A. 偵測供餐天數 (精準判斷跨月短週) ---
    date_row = None
    days_count = 0
    for i in range(min(15, len(df))):
        if any(k in str(df.iloc[i, 2]) for k in ["日期", "Date"]):
            date_row = i
            break
    if date_row is not None:
        for col in range(2, 7):
            val = str(df.iloc[date_row, col])
            if "202" in val or "/" in val: # 偵測日期格式
                days_count += 1

    # --- B. 開始稽核 (掃描午餐與點心) ---
    if date_row is None: return []
    
    for col in range(2, 7):
        if col >= len(df.columns): break
        date_txt = str(df.iloc[date_row, col]).split(" ")[0]
        
        for r_idx in range(len(df)):
            cell = str(df.iloc[r_idx, col]).strip().replace('\n', '')
            if cell in ["nan", "", "None"] or len(cell) < 2: continue

            # 1. 炸物累計 (含午餐+點心)
            if any(f in cell for f in TAGS_FRIED):
                fried_count += 1
                # 執行主張：短週(<5天)炸物限 1 次
                limit = 1 if days_count < 5 else 1 
                if fried_count > limit:
                    logs.append({"日期": date_txt, "項目": cell, "原因": f"🚩 違反限制：當週炸物累計第 {fried_count} 次 (跨月短週/規範限1次)"})

            # 2. 加工品規格 (依增補協議：強制要求標註數量或克數)
            if any(p in cell for p in TAGS_PROCESSED):
                if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", cell):
                    logs.append({"日期": date_txt, "項目": cell, "原因": "⚠️ 規格不詳：請依增補協議標註數量規格(如X2顆)或克數"})

            # 3. 調味重複性
            for s in TAGS_SEASONING:
                if s in cell:
                    if s in used_seasoning:
                        logs.append({"日期": date_txt, "項目": cell, "原因": f"❌ 口味重複：當週已出現過「{s}」調味"})
                    used_seasoning.add(s)

    return logs

# --- 2. Streamlit 介面 ---
st.set_page_config(page_title="康橋膳食稽核系統", layout="wide")
st.title("🛡️ 康橋林口校區：膳食稽核系統")
st.subheader("廠商：新北食品 (依據 114 學年合約與增補協議)")

st.sidebar.markdown(f"""
### 📋 審核標準 (共識)
1. **短週炸物**：限 1 次
2. **加工規格**：強制標註
3. **點心規範**：合併計算
""")

file = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if file:
    with st.spinner('正在執行合約自動比對...'):
        sheets = pd.read_excel(file, sheet_name=None, header=None)
        all_logs = []
        for sn, df in sheets.items():
            res = audit_menu(df, sn)
            for r in res:
                r['分頁'] = sn
                all_logs.append(r)
        
        if all_logs:
            st.error(f"🚨 偵測到 {len(all_logs)} 項不符規範，請要求廠商修正後再審：")
            # 讓表格更漂亮
            error_df = pd.DataFrame(all_logs)[['分頁', '日期', '項目', '原因']]
            st.table(error_df)
        else:
            st.success("✅ 檢查完畢！本週菜單(含點心)符合合約各項要求。")
