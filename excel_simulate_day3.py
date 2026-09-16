import os
import numpy as np
import pandas as pd
from pypinyin import lazy_pinyin
from excel_tools.main_tool import excel_sheet_loader

dfs = excel_sheet_loader(excel_path_input="simulate/simulate.xlsx", target_sheet_input=[
        "fact_reimbursements",
        "fact_reimbursement_allocations",
        "dim_applicants",
        "dim_vendors",
        "dim_organization",
        "map_org_location_bridge",
        "bldg_dim_locations",
        "bldg_master",
        "dim_date",
    ])

fact_reimbursements_tdf = dfs["fact_reimbursements_df"]
dim_applicants_tdf= dfs["dim_applicants_df"]
dim_vendors_tdf = dfs["dim_vendors_df"]
dim_organization_tdf = dfs["dim_organization_df"]
bldg_dim_locations_tdf = dfs["bldg_dim_locations_df"]
bldg_master_tdf =  dfs["bldg_master_df"]
map_org_location_bridge_tdf = dfs["map_org_location_bridge_df"]
dim_date_tdf = dfs["dim_date_df"]

# 第 21 題 (日期格式解析與年月日拆解)：
# 請將 fact_reimbursements 中的 po_date 轉為 datetime 型態，並新增三個欄位：po_year (年份), po_month (月份), po_weekday (星期幾，0為週一)。
print(f"path= {os.path.abspath(__file__)}")

# 第 22 題 (時間間隔計算與天數差)：
# 請計算每筆請款單從 po_date 到 approve_date (核准日期) 經歷了多少天（計算審核耗時），並印出平均審核天數。


# 第 23 題 (字串擷取與格式清理)：
# 在 dim_applicants 中，email 欄位格式為 "username@company.com"。請使用字串方法 (.str) 擷取 username (帳號) 並新增為獨立欄位。


# 第 24 題 (字串替換與正規表示法)：
# 請將 dim_vendors 中的 vendor_name 進行清理：去除前後空白 (.str.strip())，並將名稱中的 "股份有限公司" 或 "有限公司" 統一替換為 "Corp" (可使用 .str.replace())。


# 第 25 題 (字串模糊搜尋與篩選)：
# 請篩選出 fact_reimbursements 中 description (請款說明) 包含 "差旅" 或 "出差" 的所有紀錄（提示：使用 .str.contains()）。



# 第 26 題 (單條件分類 np.where)：
# 請使用 np.where() 在 fact_reimbursements 新增欄位 is_large_amount，若 payment_amount >= 100000 標註為 'High'，否則標註為 'Normal'。


# 第 27 題 (多條件複雜邏輯 np.select)：
# 請根據 payment_amount 對請款單進行金額評級 (amount_grade)：
# 1. > 200,000 -> 'Tier 1'
# 2. 50,000 ~ 200,000 -> 'Tier 2'
# 3. < 50,000 -> 'Tier 3'


# 第 28 題 (連續變數分箱 pd.cut)：
# 請使用 pd.cut() 將 payment_amount 依據 [0, 10000, 50000, 200000, np.inf] 切分為 4 個區間標籤：['微額', '小額', '中額', '大額']，並統計各區間筆數。


# 第 29 題 (分位數等頻分箱 pd.qcut)：
# 請使用 pd.qcut() 將 saving_pct 按照資料百分位數均等分為 4 個等級 (Quartiles: Q1, Q2, Q3, Q4)，並計算每個等級的平均 saving_pct。


# 第 30 題 (布林遮罩與多重條件填補)：
# 在 fact_reimbursements 中，若 project_no 為空值且 capex_expense 為 'CAPEX'，則將 project_no 填補為 'CAPEX未指派'；若為 'OPEX' 則填補為 'OPEX常態費用'。




# 第 31 題 (數據排序與重置索引)：
# 請將 fact_reimbursements 依照 applicant_id 與 po_date 由舊到新排序，並重置索引 (reset_index(drop=True))。



# 第 32 題 (位移函數 shift)：
# 接續第 31 題排序後的資料，請以 applicant_id 為分組，新增欄位 prev_payment_amount，代表該申請人「上一筆請款的金額」（提示：groupby + shift(1)）。



# 第 33 題 (計算前後累積差額)：
# 接續第 32 題，請計算每筆請款與該申請人上一筆請款的金額差額 (payment_amount - prev_payment_amount)。



# 第 34 題 (組內累積求和 cumsum)：
# 請以 applicant_id 為分組，計算每位申請人隨時間推進的「累積請款總金額」(allocated_amount 或 payment_amount 的 cumsum())。



# 第 35 題 (滑動視窗平均 rolling mean)：
# 請將資料按 po_date 排序後，計算全公司每日請款金額的 7 天移動平均值 (7-day rolling mean)。




# 第 36 題 (基礎透視表 pivot_table)：
# 請建立透視表，列 (index) 為 user_group (部門)，欄 (columns) 為 capex_expense (CAPEX/OPEX)，值為 payment_amount 的總和 (sum)。



# 第 37 題 (多聚合函數透視表)：
# 請建立透視表，統計每個 user_group 的 payment_amount 的「筆數 (count)」、「總金額 (sum)」與「平均金額 (mean)」，並自動加上小計列 (margins=True)。



# 第 38 題 (交叉表 crosstab)：
# 請使用 pd.crosstab() 計算 user_group 與 capex_expense 的「請款筆數次數分配表」，並顯示欄位比例 (normalize='index')。



# 第 39 題 (資料長寬轉換 melt)：
# 假設有一張寬表 (Month, CAPEX, OPEX)，請使用 pd.melt() 將其轉換為長表 (Month, Expense_Type, Amount)。
# 模擬資料可自行建立或使用 36 題結果。



# 第 40 題 (綜合實戰：異常請款行為偵測)：
# 請找出「單日請款次數 > 3 次」或是「單筆金額超過該部門平均金額 3 倍」的極端請款紀錄（提示：groupby + transform('mean')）。