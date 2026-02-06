import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁設定 (維持 Alison 原始標題與佈局)
st.set_page_config(page_title="團膳區(新北食品) 全方位稽核系統", layout="wide")

# --- 註解：製作者 Alison ---
# 樣式定義：黑底白字 30 級 (針對紅框空白缺失) / 黃底紅字 (針對規格缺失)
STYLE = {
    "BLACK_ALERT": {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=30, color="FFFFFF", bold=True)},
    "YELLOW_SPEC": {"fill": PatternFill("solid", fgColor="FFFF00"), "font": Font(name="微軟正黑體", size=14, color="FF0000", bold=True)}
}

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 關鍵：強迫程式看見空白，將 NaN 填補為 "VOID"
        df_audit = df.fillna("VOID")
        
        # 定位日期定位點 (通常在 C 欄)
        d_row = None
        for i, row in df_audit.iterrows():
            if "日期" in str(row[2]):
                d_row = i
                break
        if d_row is None: continue

        # 掃描週一到週五 (D-H 欄)
        for col in range(3, 8):
            # 取得該欄日期 (用於 Log 紀錄)
            date_val = str(df_audit.iloc[d_row, col]).split("\n")[0]
            
            # 從日期列往下開始地毯式搜索
            for r_idx in range(d_row + 1, len(df_audit)):
                label = str(df_audit.iloc[r_idx, 2]).strip()
                content = str(df_audit.iloc[r_idx, col]).strip()
                cell = ws.cell(row=r_idx+1, column=col+1)

                # --- 核心邏輯：強制偵測空白 (針對 4/28, 4/29) ---
                # 只要是關鍵欄位，內容是 VOID 或 只有空白字串，一律噴黑漆
                mandatory_labels = ["熱量", "主菜", "副菜", "套餐"]
                
                # 偵測 A：標籤存在但內容空白 (熱量缺失、菜名缺失)
                if any(m_label in label for m_label in mandatory_labels):
                    if content in ["VOID", "", "nan", "0"]:
                        # 檢查 4/29 特殊情況：如果菜名是空的，但下面那一格「食材明細」有字，那更是必抓！
                        is_structure_fail = False
                        try:
                            next_val = str(df_audit.iloc[r_idx+1, col]).strip()
                            if next_val != "VOID": is_structure_fail = True
                        except: pass
                        
                        if is_structure_fail or "熱量" in label:
                            cell.fill, cell.font = STYLE["BLACK_ALERT"]["fill"], STYLE["BLACK_ALERT"]["font"]
                            logs.append({"日期": date_val, "類別": "嚴重缺失", "原因": f"⚠️ {label} 欄位空白！"})

                # --- 核心邏輯：規格審核 (原本穩定的功能) ---
                if "白帶魚" in content and "150g" not in content:
                    cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                    logs.append({"日期": date_val, "類別": "規格不符", "原因": "白帶魚未標 150g"})
                
                if "獅子頭" in content and "60gX2" not in content:
                    cell.fill, cell.font = STYLE["YELLOW_SPEC"]["fill"], STYLE["YELLOW_SPEC"]["font"]
                    logs.append({"日期": date_val, "類別": "規格不符", "原因": "獅子頭未標 60gX2"})

    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue()

# --- Streamlit 介面 (完全依照 Alison 規範) ---
st.title("🛡️ 團膳區(新北食品) 全方位稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("📂 請上傳菜單 Excel 檔案進行審核", type=["xlsx"])
if up:
    with st.spinner("稽核系統執行中..."):
        results, data = audit_process(up)
        
    if results:
        st.error(f"🚩 抓到了！共發現 {len(results)} 項缺失（含紅框處空白與規格缺失）。")
        st.table(pd.DataFrame(results))
        st.download_button(
            label="📥 下載標註完成之退件檔",
            data=data,
            file_name=f"退件建議_{up.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 審核完畢，此菜單結構完整且規格正確！")
