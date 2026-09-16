import pandas as pd
import numpy as np

# ==========================================
# 1. 讀取原始地獄 Excel 檔案
# ==========================================
# 🎯 請將 'FinalTest.xlsx' 改為妳電腦中實際的檔案路徑
file_path = '/Users/changting-chia/Desktop/FinalTest.xlsx' 

# 由於標頭在第 4 列 (Python index 從 0 開始算，所以是 header=3)
df_old = pd.read_excel(file_path, sheet_name='工作表1', header=3)

# 清除欄位名稱前後的空白與換行符號，防止比對失敗
df_old.columns = df_old.columns.str.strip().str.replace('\n', ' ')

# ==========================================
# 2. 核心大絕招：動態欄位模糊對照 
# ==========================================
col_map = {}
keywords = {
    'Date': ['Date', 'po日期'],
    'PC_code': ['PC code'],
    'BUSU': ['BUSU / CRESA'],
    'Applicant': ['Applicant', '申請人'],
    'Vendor': ['Vendor', '廠商'],
    'Service': ['Service', '服務摘要'],
    'Amount': ['Amount', '金額'],
    'Category': ['Category', '項目'],
    'CapexExpense': ['Capex/Expense'],
    'PO_NO_USER': ['PO. NO.', 'PO單號'],
    'Group': ['Group'],
    'InvoiceNo': ['發票號碼'],
    'ToFinance': ['To Finance'],
    'GLCode': ['GL Code'],
    'PaymentAmount': ['Payment Amount'],
    'GLDate': ['GL Date']
}

# 用妳最擅長的 For 迴圈，幫髒標頭建立對應字典
for std_name, kw_list in keywords.items():
    for actual_col in df_old.columns:
        if any(kw in actual_col for kw in kw_list):
            col_map[std_name] = actual_col
            break

# ==========================================
# 3. 雙重 For 迴圈：中台清洗與橫向展開 (Staging)
# ==========================================
stg_rows = [] # 準備存放清洗完的每一列資料 (像 Java 的 List<Map>)

prev_date = ""
date_counter = 1

# 外迴圈：逐列讀取資料
for index, row in df_old.iterrows():
    raw_date = row[col_map['Date']]
    if pd.isna(raw_date):
        continue # 遇到空行就 skip
        
    # 格式化日期為 YYYYMMDD
    current_date = pd.to_datetime(raw_date).strftime('%Y%m%dd')
    current_date_str = pd.to_datetime(raw_date).strftime('%Y%m%d')
    
    # 自動生成單號邏輯 (如 20260102-1)
    if current_date_str == prev_date:
        date_counter += 1
    else:
        date_counter = 1
    generated_no = f"{current_date_str}-{date_counter}"
    prev_date = current_date_str
    
    # 讀取 PC Code 並處理留白
    pc_code_input = str(row[col_map['PC_code']]).strip()
    if pc_code_input == 'nan' or not pc_code_input:
        pc_code_input = "EMPTY"
        
    # 🎯 用斜線拆分資料 (核心拆彈)
    code_array = [c.strip() for c in pc_code_input.split('/') if c.strip()]
    total_splits = len(code_array) if len(code_array) > 0 else 1
    
    # 内迴圈：處理拆分出來的每一個子 Code，進行攤提與貼標籤
    for single_code in code_array:
        
        # 🎯 貼標籤邏輯 (白名單判定)
        if single_code in ["00QH", "88QO", "8P13", "00QG", "8Q34", "18QB", "18QC"]:
            code_type = "BuildingCode"
        elif single_code in ["C669", "C631", "C634", "45CB"]:
            code_type = "PCcode"
        else:
            code_type = "Unknown"
            
        # 計算等比例攤提金額
        amt = float(row[col_map['Amount']]) / total_splits if not pd.isna(row[col_map['Amount']]) else 0
        pay_amt = float(row[col_map['PaymentAmount']]) / total_splits if not pd.isna(row[col_map['PaymentAmount']]) else 0
        
        # 組裝成乾淨的新列
        new_row = {
            "單號PO主鍵": generated_no,
            "PO日期": raw_date,
            "填寫者輸入的PC code": pc_code_input,
            "拆分局部的PC code": single_code,
            "Code類別": code_type,
            "BUSU / CRESA": row[col_map['BUSU']],
            "申請人": row[col_map['Applicant']],
            "廠商": row[col_map['Vendor']],
            "服務摘要": row[col_map['Service']],
            "金額（含稅）": amt,
            "項目": row[col_map['Category']],
            "Capex/Expense": row[col_map['CapexExpense']],
            "PO單號": row[col_map['PO_NO_USER']],
            "Group": row[col_map['Group']],
            "發票號碼": row[col_map['InvoiceNo']],
            "To Finance": row[col_map['ToFinance']],
            "GL Code": row[col_map['GLCode']],
            "Payment Amount": pay_amt,
            "GL Date": row[col_map['GLDate']]
        }
        stg_rows.append(new_row)

# ==========================================
# 4. 將暫存資料轉成 DataFrame 並倒回 Excel
# ==========================================
df_stg = pd.DataFrame(stg_rows)

# 使用 openpyxl 引擎，將新工作表「追加」到原有的 Excel 檔案中
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_stg.to_excel(writer, sheet_name='中台拆彈暫存', index=False)

print("🚀 【Python 報告長官】Mac 拆彈任務圓滿成功！'中台拆彈暫存' 工作表已完美產出！")