import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 樣式定義
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)} # 真空
STYLE_LOW = {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)} # 份數不足
STYLE_CAL = {"fill": PatternFill("solid", fgColor="FFCCFF"), "font": Font(name="微軟正黑體", size=12, color="800000", bold=True)} # 熱量異常

def to_float(val):
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def alison_master_audit(file):
    fname = file.name
    if any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode = "新北食品-教育學部"
        nutri_map = {"熱量": 9, "全榖": 10, "豆魚": 11, "蔬菜": 12} # 假設的欄位索引
    elif any(kw in fname for kw in ["美食街", "素食"]):
        mode = "新北食品-美食街/素食"
        nutri_map = {"熱量": 3, "全榖": 4, "豆魚": 5, "蔬菜": 6}
    else:
        return None, "BLOCK", None, {}

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        stats = {"掃描總欄位": 0, "熱量檢核": 0, "份數檢核": 0}

        for sn, df in sheets_df.items():
            ws = wb[sn]
            df_audit = df.astype(str).replace(['nan', 'NaN', 'None'], '')
            
            for r_idx in range(len(df_audit)):
                label = str(df_audit.iloc[r_idx, 0]).strip()
                
                # 識別日期行
                if ("/" in label or "202" in label) and len(label) < 15:
                    for item_name, n_idx in nutri_map.items():
                        if n_idx >= len(df_audit.columns): continue
                        
                        raw_val = df_audit.iloc[r_idx, n_idx].strip()
                        stats["掃描總欄位"] += 1
                        cell = ws.cell(row=r_idx+1, column=n_idx+1)

                        # A. 檢查真空 (漏填)
                        if raw_val == "":
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌漏填"
                            logs.append({"日期": label, "項目": item_name, "原因": "真空漏填"})
                            continue

                        # B. 檢查具體指標內容
                        val = to_float(raw_val)
                        if item_name == "熱量":
                            stats["熱量檢核"] += 1
                            if val < 650 or val > 800:
                                cell.fill, cell.font = STYLE_CAL["fill"], STYLE_CAL["font"]
                                logs.append({"日期": label, "項目": "熱量", "原因": f"異常: {val} Kcal"})
                        
                        elif item_name in ["全榖", "豆魚", "蔬菜"]:
                            stats["份數檢核"] += 1
                            limit = 1.0 if item_name == "蔬菜" else 2.0
                            # 填 0 合法 (當天不供)，但大於 0 卻小於標準則報警
                            if 0 < val < limit:
                                cell.fill, cell.font = STYLE_LOW["fill"], STYLE_LOW["font"]
                                logs.append({"日期": label, "項目": item_name, "原因": f"份數不足: {val}"})

        return logs, mode, wb, stats
    except Exception as e:
        return None, f"ERROR: {str(e)}", None, {}

# --- Streamlit UI ---
st.set_page_config(page_title="新北食品進階稽核", layout="wide")
st.title("🛡️ 團膳區(新北食品) 菜單自主稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 上傳菜單 Excel", type=["xlsx"])
if up:
    logs, m, wb_out, stats = alison_master_audit(up)
    
    if m == "BLOCK":
        st.error("❌ 檔名不符！")
    else:
        # --- 確實度透明報告區 ---
        st.info("### 🔍 確實度稽核報告")
        col1, col2, col3 = st.columns(3)
        col1.metric("總掃描點", stats.get("掃描總欄位", 0))
        col2.metric("熱量符合性檢查", f"{stats.get('熱量檢核', 0)} 天")
        col3.metric("營養份數檢查", f"{stats.get('份數檢核', 0)} 項")

        if logs:
            st.error(f"🚩 偵測到 {len(logs)} 項法規與格式異常")
            st.table(pd.DataFrame(logs))
            # 下達下載
            out = BytesIO()
            wb_out.save(out)
            st.download_button("📥 下載 Alison 專業標註檔", out.getvalue(), f"退件_{up.name}")
        else:
            st.success("🎉 經『熱量、份數、真空』三大檢核點確認：數據完全符合新北規範！")
