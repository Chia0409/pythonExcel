import pandas as pd
import numpy as np
from sqlalchemy import create_engine
import random

# 設定隨機數種子以保持可重現性
np.random.seed(42)

# --- 生成 50 筆客戶 ---
regions = ['North', 'South', 'Central', 'East']
customers_list = []
for i in range(11, 51):
    c_id = f"C{i:03d}"
    customers_list.append({
        'customer_id': c_id,
        'customer_name': f"User_{i}",
        'region': np.random.choice(regions),
        'join_date': pd.to_datetime('2024-01-01') + pd.to_timedelta(np.random.randint(0, 500), unit='D')
    })
df_customers = pd.DataFrame(customers_list)

# --- 生成 50 筆合約 ---
contracts_list = []
types = ['Credit Card', 'Loan']
statuses = ['Active', 'Expired', 'Suspended']
for i in range(11, 51):
    c_id = f"T{i:03d}"
    cust_id = f"C{np.random.randint(1, 45):03d}" # 隨機選前45位客戶
    contracts_list.append({
        'contract_id': c_id,
        'customer_id': cust_id,
        'contract_type': np.random.choice(types),
        'status': np.random.choice(statuses),
        'credit_limit': np.random.randint(5, 50) * 10000,
        'expiry_date': pd.to_datetime('2026-01-01') + pd.to_timedelta(np.random.randint(0, 365), unit='D')
    })
df_contracts = pd.DataFrame(contracts_list)

# --- 存成 Excel ---
df_customers.to_excel('customers_v50.xlsx', index=False)
df_contracts.to_excel('contracts_v50.xlsx', index=False)
print("🎯 50 筆 Excel 資料已更新！")

# --- 寫入資料庫 (使用 if_exists='replace'/'append' 是附加在後面 直接更新 Table) ---
db_url = 'mysql+pymysql://root:90409991@localhost:3306/learninghub' 
engine = create_engine(db_url)

df_customers.to_sql('customers', con=engine, if_exists='append', index=False)
df_contracts.to_sql('contracts', con=engine, if_exists='append', index=False)

print("🚀 資料庫 Table (customers & contracts) 已更新為 50 筆記錄！")