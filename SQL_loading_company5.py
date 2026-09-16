import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# 1. 設定檔案路徑與資料庫連線
file_path = "simulate.xlsx"
db_connection_url = "mysql+pymysql://root:90409991@localhost:3306/company_simulate"

# 2. 自動掃描兩個工作頁，獲取日期區間 (Min Date & Max Date)
excel_file = pd.ExcelFile(file_path)
all_dates = []

for sheet in ['fact_reimbursements', 'fact_reimbursement_allocations']:
    if sheet in excel_file.sheet_names:
        df_temp = pd.read_excel(excel_file, sheet_name=sheet)
        # 自動尋找資料型態為 datetime 或欄位名稱帶有 date 的欄位
        for col in df_temp.columns:
            if 'date' in col.lower() or pd.api.types.is_datetime64_any_dtype(df_temp[col]):
                dates = pd.to_datetime(df_temp[col], errors='coerce').dropna()
                all_dates.extend(dates)

# 若抓不到日期，預設補齊 2025 全年區間
if all_dates:
    min_date = min(all_dates).replace(month=1, day=1)
    max_date = max(all_dates).replace(month=12, day=31)
else:
    min_date = pd.Timestamp("2025-01-01")
    max_date = pd.Timestamp("2025-12-31")

# 3. 建立連續的日期 Range
date_series = pd.date_range(start=min_date, end=max_date, freq='D')

# 4. 構建 dim_date 資料表欄位 (對應 SQL 表格結構)
dim_date = pd.DataFrame({'calendar_date': date_series})

dim_date['date_key'] = dim_date['calendar_date'].dt.strftime('%Y%m%d').astype(int)
dim_date['year'] = dim_date['calendar_date'].dt.year
dim_date['quarter'] = dim_date['calendar_date'].dt.quarter
dim_date['quarter_name'] = 'Q' + dim_date['quarter'].astype(str)
dim_date['month'] = dim_date['calendar_date'].dt.month
dim_date['month_name'] = dim_date['calendar_date'].dt.strftime('%B')
dim_date['year_month'] = dim_date['calendar_date'].dt.strftime('%Y-%m')
dim_date['day'] = dim_date['calendar_date'].dt.day
dim_date['day_of_week'] = dim_date['calendar_date'].dt.dayofweek + 1  # 1=Mon, 7=Sun
dim_date['day_name'] = dim_date['calendar_date'].dt.strftime('%A')

# 邏輯判斷 (工作日/假日/月底)
dim_date['is_workday'] = dim_date['day_of_week'].apply(lambda x: 'Y' if x <= 5 else 'N')
dim_date['is_holiday'] = dim_date['day_of_week'].apply(lambda x: 'Y' if x > 5 else 'N') # 可另行寫入台灣國定假日清單
dim_date['is_month_end'] = dim_date['calendar_date'].dt.is_month_end.apply(lambda x: 'Y' if x else 'N')

# 會計年度與會計季度 (預設與曆年相同，若銀行有特殊財年起始月份可調整偏移量)
dim_date['fiscal_year'] = dim_date['year']
dim_date['fiscal_quarter'] = 'FY-' + dim_date['quarter_name']

# 調整欄位順序以完全符合 SQL CREATE TABLE
columns_order = [
    'date_key', 'calendar_date', 'year', 'quarter', 'quarter_name',
    'month', 'month_name', 'year_month', 'day', 'day_of_week', 'day_name',
    'is_workday', 'is_holiday', 'is_month_end', 'fiscal_year', 'fiscal_quarter'
]
dim_date = dim_date[columns_order]

# 5. 寫入 MySQL 資料庫 (append 模式，避免蓋掉已經 create 好的 table schema)
engine = create_engine(db_connection_url)
dim_date.to_sql(name='dim_date', con=engine, if_exists='append', index=False)
print("Successfully written dim_date to MySQL database.")

# 6. 寫回 Excel (新增/蓋寫 dim_date 工作頁)
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    dim_date.to_excel(writer, sheet_name='dim_date', index=False)
print("Successfully added dim_date sheet to simulate.xlsx.")