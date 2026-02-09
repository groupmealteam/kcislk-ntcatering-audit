import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 樣式設定
STYLE_ERR = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=14, color="FFFFFF", bold=True)}

def alison_smart_audit(file):
    fname = file.name
    # 根據 Alison 規範判定模式與營養欄位位置
    if any(kw in fname for kw in ["小學", "幼兒園"]):
        mode = "教育學部"
        label_idx = 0
        data_indices = [1, 2, 3, 4, 5, 6, 7] # 午餐內容
        # 營養分析欄位座標 (全榖, 豆魚, 蔬菜, 油脂, 水果, 奶類, 熱量)
        nutri_indices = [9, 10, 11, 12, 13, 14, 15] 
    else:
        # 美食街模式 (熱量通常在 C 欄標籤下的固定列，邏輯略有不同)
        mode, label_idx, data_indices, nutri_indices = "美食街", 2, [3, 4, 5, 6, 7], [3, 4, 5, 6, 7]

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 轉為字串並清理空格
        df_audit = df.astype(str).applymap(lambda x: "" if str(x).strip().lower() in ['nan', 'none', '0', '0.0', ''] else str(x).strip())
        
        for r_idx in range(len(df_audit)):
            # 1. 抓日期行 (判斷當天是否有餐)
            # 小學部格式：日期在 A 欄，若 A 欄有日期，代表這整橫列都要審核
            cell_date = df_audit.iloc[r_idx, 0]
            if "/" in cell_date and "(" in cell_date:
                
                # --- A. 營養成分分析全檢 (Alison 要求的聰明審核) ---
                has_lunch = df_audit.iloc[r_idx, 1] != "" # 只要主食有填，右邊就不能空
                if has_lunch:
                    for n_idx in nutri_indices:
                        if n_idx >= df_audit.shape[1]: continue
                        val = df_audit.iloc[r_idx, n_idx]
                        
                        # 如果數值是空的，或者不是數字格式 (防止廠商填文字錯位)
                        is_numeric = val.replace('.','',1).isdigit() 
                        if val == "" or not is_numeric:
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌數據缺失"
                            logs.append({"分頁": sn, "日期": cell_date, "缺失": f"第{n_idx+1}欄營養數據異常"})

                # --- B. 垂直菜名黑洞檢查 ---
                for c_idx in data_indices:
                    content = df_audit.iloc[r_idx, c_idx]
                    if content == "":
                        try:
                            detail = df_audit.iloc[r_idx+1, c_idx]
                            if detail != "":
                                cell = ws.cell(row=r_idx+1, column=c_idx+1)
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填菜名"
                                logs.append({"分頁": sn, "日期": cell_date, "缺失": "有明細無菜名"})
                        except: pass

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# Streamlit 介面保持不變
