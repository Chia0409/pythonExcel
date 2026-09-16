import pandas as pd
from sqlalchemy import create_engine

# ==========================================
# STEP 1: 用 Python 字典定義資料並轉為 DataFrame
# ==========================================

customers_data = {
    'customer_id': ['C001', 'C002', 'C003', 'C004', 'C005', 'C006', 'C007', 'C008', 'C009', 'C010'],
    'customer_name': ['Alice', 'Bob', 'Charlie', 'David', 'Eva', 'Frank', 'Grace', 'Henry', 'Ivy', 'Jack'],
    'region': ['North', 'South', 'Central', 'North', 'East', 'North', 'South', 'Central', 'North', 'South'],
    'join_date': ['2025-01-15', '2025-03-22', '2024-11-05', '2026-02-10', '2025-07-19', '2026-04-01', '2025-09-30', '2025-12-25', '2026-05-12', '2026-05-20']
}
df_customers = pd.DataFrame(customers_data)

contracts_data = {
    'contract_id': ['T001', 'T002', 'T003', 'T004', 'T005', 'T006', 'T007', 'T008', 'T009', 'T010'],
    'customer_id': ['C001', 'C001', 'C002', 'C003', 'C004', 'C005', 'C006', 'C007', 'C008', 'C004'],
    'contract_type': ['Credit Card', 'Loan', 'Credit Card', 'Credit Card', 'Credit Card', 'Loan', 'Credit Card', 'Credit Card', 'Credit Card', 'Credit Card'],
    'status': ['Active', 'Active', 'Active', 'Expired', 'Active', 'Active', 'Active', 'Active', 'Active', 'Suspended'],
    'credit_limit': [350000, 1000000, 250000, 150000, 450000, 500000, 80000, 300000, 120000, 200000],
    'expiry_date': ['2026-10-31', '2029-12-31', '2026-08-15', '2025-12-31', '2026-12-15', '2028-05-20', '2026-11-30', '2027-04-18', '2026-06-30', '2026-09-01']
}
df_contracts = pd.DataFrame(contracts_data)

# ==========================================
# STEP 2: 匯出成 Excel 檔案儲存（需要安裝 openpyxl）
# ==========================================
df_customers.to_excel('customers.xlsx', index=False)
df_contracts.to_excel('contracts.xlsx', index=False)
print("🎯 成功生成 customers.xlsx 與 contracts.xlsx 檔案！")

# ==========================================
# STEP 3: 連線到你的本地資料庫並 Loading 進去
# ==========================================
# 💡 請根據你實際使用的資料庫種類，將下方連線字串（URL）進行修改：
# MySQL 範例: 'mysql+pymysql://帳號:密碼@localhost:3306/learninghub'


db_url = 'mysql+pymysql://root:90409991@localhost:3306/learninghub' 
engine = create_engine(db_url)

# 使用 pandas 的 to_sql 功能，if_exists='append' 代表資料會直接塞進你剛建好的 table 裡
df_customers.to_sql('customers', con=engine, if_exists='append', index=False)
df_contracts.to_sql('contracts', con=engine, if_exists='append', index=False)

print("🚀 資料已成功完美加載至 learninghub 資料庫！你可以開始寫 SQL 實測第 11-15 題了！")