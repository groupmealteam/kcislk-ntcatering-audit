import streamlit as st
import pandas as pd
import re

# 1. 定義【專業營養師眼光】關鍵字與符號
# 同步菜單定義：◎=油炸, △=加工品, ★=帶殼海鮮 [cite: 73, 113, 117]
TAGS_FRIED = ["炸", "酥", "裹粉", "爆", "脆", "可樂餅", "◎"] 
TAGS_PROCESSED = ["丸", "排", "素羊", "素火腿", "素肉", "獅子頭", "豆包", "炸豆腐", "△", "★"]
TAGS_FREQUENT = ["豆芽", "銀芽", "芽菜"] # 郁恩特別在意的頻率控管 [cite: 73]

def expert_audit(df, sheet_name):
    logs = []
    fried_indices = [] # 紀錄炸物出現的日期
    veg_counter = 0
    
    # 偵測本週天數
    date_row = None
    days_count = 0
    for i in range(min(15, len(df))):
        if any(k in str(df.iloc[i, 2]) for k in ["日期", "Date"]):
            date_row = i
            break
    if date_row is not None:
        for col in range(2, 7):
            if "202" in str(df.iloc[date_row, col]): days_count += 1

    for col in range(2, 7):
        if col >= len(df.columns): break
        date_txt = str(df.iloc[date_row, col]).split(" ")[0] if date_row else "未知"
        
        for r_idx in range(len(df)):
            cell = str(df.iloc[r_idx, col]).strip()
            if len(cell) < 1 or "nan" in cell: continue

            # A. 炸物累計 (包含◎符號) [cite: 113, 117]
            if any(f in cell for f in TAGS_FRIED):
                fried_indices.append(date_txt)
                limit = 1 if days_count < 5 else 1 # 堅持短週/一般週嚴格限制
                if len(fried_indices) > limit:
                    logs.append({"日期": date_txt, "項目": cell, "原因": f"🚩炸物超標(當週累計{len(fried_indices)}次)"})

            # B. 加工品規格 (包含△符號，依增補協議需標規格) [cite: 82, 117]
            if any(p in cell for p in TAGS_PROCESSED):
                if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", cell):
                    logs.append({"日期": date_txt, "項目": cell, "原因": "⚠️規格未標：加工品/素料/成品需標數量或克數"})

            # C. 豆芽頻率控管 (郁恩專業意見) 
            if any(v in cell for v in TAGS_FREQUENT):
                veg_counter += 1
                if veg_counter > 1:
                    logs.append({"日期": date_txt, "項目": cell, "原因": "❌食材重複：豆芽類當週出現過高"})

    return logs

# --- 介面呈現 ---
st.title("🛡️ 康橋林口：專業對齊版稽核系統")
st.markdown("##### 核心邏輯：將營養師專業直覺數位化，包含隱藏加工品與頻率控管。")

f = st.file_uploader("📂 請上傳菜單 Excel", type=["xlsx"])
if f:
    sheets = pd.read_excel(f, sheet_name=None, header=None)
    all_res = []
    for sn, df in sheets.items():
        res = expert_audit(df, sn)
        for r in res: r['分頁'] = sn; all_res.append(r)
    
    if all_res:
        st.error(f"🚨 系統依據營養師標準發現 {len(all_res)} 項疑義：")
        st.table(pd.DataFrame(all_res)[['分頁', '日期', '項目', '原因']])
    else:
        st.success("✅ 檢查完畢！本週菜單符合專業與合約標準。")
