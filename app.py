import streamlit as st
import pandas as pd
import re

# --- 1. 模擬營養師郁恩的【高強度關鍵字庫】 ---
TAGS_FRIED = ["炸", "酥", "裹粉", "爆爆", "薯", "雞塊", "酥", "脆", "可樂餅"]
# 擴大加工品認定：包含素肉系列
TAGS_PROCESSED = ["丸", "排", "素肉", "素羊", "素雞", "火腿", "獅子頭", "肉燥", "捲", "豆包", "炸豆腐", "甜不辣"]
# 調味與食材頻率監控
TAGS_SEASONING = ["沙茶", "咖哩", "腐乳", "三杯", "麻婆", "糖醋", "沙拉"]
TAGS_FREQUENT_VEG = ["豆芽", "銀芽", "芽菜"]

def audit_expert_logic(df, sheet_name):
    logs = []
    fried_count = 0
    veg_count = 0
    seasoning_set = set()
    
    # 偵測天數
    date_row = None
    days_count = 0
    for i in range(min(15, len(df))):
        if any(k in str(df.iloc[i, 2]) for k in ["日期", "Date"]):
            date_row = i
            break
    if date_row is not None:
        for col in range(2, 7):
            if "202" in str(df.iloc[date_row, col]): days_count += 1

    # 開始掃描
    for col in range(2, 7):
        if col >= len(df.columns): break
        date_label = str(df.iloc[date_row, col]).split(" ")[0] if date_row else "未知"
        
        for r_idx in range(len(df)):
            cell = str(df.iloc[r_idx, col]).strip()
            if len(cell) < 2 or "nan" in cell: continue

            # 1. 炸物嚴查 (跨月短週 1 次)
            if any(f in cell for f in TAGS_FRIED):
                fried_count += 1
                limit = 1 if days_count < 5 else 1 
                if fried_count > limit:
                    logs.append({"日期": date_label, "項目": cell, "原因": f"🚩炸物累計{fried_count}次(合約規範短週限1次)"})

            # 2. 加工品規格 (郁恩最在意的數量標註)
            if any(p in cell for p in TAGS_PROCESSED):
                if not re.search(r"(\d+[xX*×]\d+)|(\d+\s*[gG克])", cell):
                    logs.append({"日期": date_label, "項目": cell, "原因": "⚠️規格不詳(加工/成品類請標註數量或克數)"})

            # 3. 豆芽類頻率控制 (郁恩 5 月初審意見)
            if any(v in cell for v in TAGS_FREQUENT_VEG):
                veg_count += 1
                if veg_count > 1: # 當週不重複
                    logs.append({"日期": date_label, "項目": cell, "原因": "❌食材重複：豆芽/銀芽類當週頻率過高"})

            # 4. 強烈調味重複
            for s in TAGS_SEASONING:
                if s in cell:
                    if s in seasoning_set:
                        logs.append({"日期": date_label, "項目": cell, "原因": f"❌口味單調：當週已出現過「{s}」調味"})
                    seasoning_set.add(s)

    return logs

# --- 介面呈現 ---
st.title("🛡️ 康橋膳食稽核系統 (114學年合約規格版)")
st.info("已同步營養師初審意見：嚴查加工品規格、豆芽頻率、短週炸物限制。")

f = st.file_uploader("上傳菜單 Excel", type=["xlsx"])
if f:
    sheets = pd.read_excel(f, sheet_name=None, header=None)
    all_res = []
    for sn, df in sheets.items():
        res = audit_expert_logic(df, sn)
        for r in res: r['分頁'] = sn; all_res.append(r)
    
    if all_res:
        st.error(f"🚩 偵測到 {len(all_res)} 項不合規，請要求廠商依合約修正：")
        st.table(pd.DataFrame(all_res)[['分頁', '日期', '項目', '原因']])
    else:
        st.success("✅ 本週菜單符合合約與營養師初步審核標準。")
