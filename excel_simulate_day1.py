import os
import pandas as pd
import numpy as np

imput_file = "simulate.xlsx"
output_file = "simulate_output.xlsx"
target_sheets = ["fact_reimbursements", "dim_applicants"]

def select_df(input_df, cols:list):
    selected_df = input_df[cols] 
    return selected_df


excel_file = pd.ExcelFile(imput_file)


if all (sheet in excel_file.sheet_names for sheet in target_sheets):
    factReim_df = pd.read_excel(excel_file, sheet_name = "fact_reimbursements")
    app_df = pd.read_excel(excel_file, sheet_name = "dim_applicants")
    print("sheet all here")
    # 1. 資料讀取與合併（JOIN）：
    # 讀取 fact_reimbursements（請款主表）與 dim_applicants（申請人主表）。 將兩張表透過 applicant_id 進行 LEFT JOIN，取得申請人的姓名 (applicant_name) 與組別 (user_group)。
    factremi_app_m_df = pd.merge(factReim_df, app_df, "left", on = "applicant_id")
    factremi_app_m_df2 = (
        factremi_app_m_df[["applicant_name","user_group"]]
        .head(10)
    )
    # print(factremi_app_m_df2)

    # 2. 條件篩選與計算（WHERE & CASE WHEN）：
    # 篩選出 支出類型 (capex_expense) 為 'CAPEX' 且 實付金額 (payment_amount) 大於等於 100,000 的請款紀錄。新增一個欄位 discount_flag：若節省比例 saving_pct > 0 : 標註為 '有折扣'其他狀況: 標註為 '無折扣'
    factReim_capex_df = factReim_df.loc[ (factReim_df["capex_expense"]=="CAPEX") & (factReim_df["payment_amount"] >= 100000)]
    # print(factReim_capex_df.head(5))
    # np.where 
    factReim_capex_df["discount_flag"] = np.where(
        factReim_capex_df["saving_pct"] > 0, '有折扣', '無折扣'
    )
    # print(factReim_capex_df.head(5))
    # np.select
    conds = [
        factReim_capex_df["saving_pct"] > 10,
        factReim_capex_df["saving_pct"] > 0
    ]
    choice = ["高折扣","低折扣"]
    factReim_capex_df["discount_level"] = np.select(
        conds, choice, default = "no"
    )
    # print(factReim_capex_df.head(5))
    # df["col"].apply
    factReim_capex_df["payment_level"] = factReim_capex_df["payment_amount"].apply(lambda x: "high level" if x >= 300000 else "low level")
    # print(factReim_capex_df.head(5))

    # 定義一個傳統的 Python 判斷函式
    def check_payment_level(row):
        if row["payment_amount"] > 300000:
            return "high"
        else:
            return "low"
    # df.apply 套用至每一列
    factReim_capex_df["payment_level2"] = factReim_capex_df.apply(check_payment_level, axis=1)
    
    # df.apply + lambda 套用至每一列
    factReim_capex_df["payment_level3"] = factReim_capex_df.apply(
        lambda row: "high" if row["payment_amount"] > 300000 else "low", axis=1
    )
    # print(factReim_capex_df.head(5))

    # 4.排序與指定欄位輸出（ORDER BY & SELECT）：
    # 依照 payment_amount 由大到小（降冪） 排序。
    # 僅保留以下欄位：reimburse_id, applicant_name, user_group, service_summary, payment_amount, saving_pct, discount_flag
    factremi_app_m_df["discount_flag"] = np.where(
        factremi_app_m_df["saving_pct"] > 0, "有折扣", "無折扣"
    )

    factremi_app_m_df3 = (
        factremi_app_m_df[["reimburse_id","applicant_name","user_group","service_summary","payment_amount","saving_pct","discount_flag"]]
        .sort_values(by="payment_amount", ascending=False)
    )
    # print(factremi_app_m_df3.head(5))
    
    # 5.安全匯出至「新 Excel 檔案」：
    # 將處理好的 DataFrame 寫入一個全新的檔案 contracts_v50_result.xlsx 中，工作表命名為 CAPEX大額請款，並且不要印出預設索引列（index=False）。
    # if os.path.exists(output_file):
    #     with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    #         factremi_app_m_df3.to_excel(writer, sheet_name="CAPEX大額請款", index=False)
    #         print("excel output done")
    # else:
    #     factremi_app_m_df3.to_excel(output_file, sheet_name= "CAPEX大額請款", index=False)
    #     print("excel output done")


    # 第 6 題 (SELECT 多欄位)：
    # 請從 fact_reimbursements 集中選取 reimburse_id, po_date, vendor_id, payment_amount 這四個欄位，並取出前 8 筆資料。
    factReim_df2 = (
        factReim_df[["reimburse_id", "po_date", "vendor_id", "payment_amount"]]
        .head(8)
    )
    # print(factReim_df2)

    # 第 7 題 (ORDER BY 多重欄位排序)：
    # 請將 fact_reimbursements 依照 capex_expense（資本/營運支出）進行升冪排序；若支出類型相同，則依照 payment_amount（實付金額）進行降冪排序。
    factReim_df3 = factReim_df.sort_values(by=["capex_expense", "payment_amount"] , ascending=[True,False])
    # print(factReim_df3.head(5))

    # 第 8 題 (WHERE 數字比較與不等於)：
    # 請找出 fact_reimbursements 中，節省比例 saving_pct 不等於 0 且實付金額 payment_amount 低於 50,000 的所有請款紀錄。
    factReim_df4 = factReim_df.loc[(factReim_df["saving_pct"]!= 0) & (factReim_df["payment_amount"]< 50000) ]
    # print(factReim_df4.head(5))

    # 第 9 題 (WHERE 空值判斷 IS NULL / NOT NULL)：
    # 在 fact_reimbursements 中，有些紀錄沒有專案編號 (project_no)。請篩選出 project_no 為空值 (NaN) 的請款單，並計算共有幾筆。
    factReim_df5 = factReim_df.loc[ factReim_df["project_no"].isnull() ]
    # print(factReim_df5["project_no"].isnull())
    factReim_df5_row_len = len(factReim_df5)
    # print(factReim_df5_row_len)

    # 第 10 題 (WHERE 空值過濾)：
    # 承上題，請篩選出 project_no 有值（非空值 / NOT NULL） 的紀錄，並只顯示 reimburse_id 與 project_no 兩個欄位。
    factReim_df6 = factReim_df.loc[ factReim_df["project_no"].notnull() ]
    factReim_df6_row_len = len(factReim_df6)
    # print(factReim_df6_row_len)

    # 第 11 題 (IN 運算子)：
    # 請篩選出供應商 ID (vendor_id) 屬於 1, 3, 5（分別代表台電、中華電信、台灣大哥大等）的所有請款紀錄。（提示：使用 .isin()）
    factReim_df7 = (
        factReim_df.loc[ factReim_df["vendor_id"].isin([1,3,5]), ["reimburse_id", "payment_amount", "applicant_id", "vendor_id"] ]
        .head(10)
    )
    # print(factReim_df7)

    # 第 12 題 (NOT IN 運算子)：
    # 請篩選出申請人 ID (applicant_id) 不屬於 ['EMP001', 'EMP002', 'EMP003'] 的請款紀錄。
    factReim_df8 = (
        factReim_df.loc[ ~factReim_df["applicant_id"].isin(['EMP001', 'EMP002', 'EMP003']), ["reimburse_id", "payment_amount", "applicant_id", "vendor_id"] ]
        .head(10)
    )
    # print(factReim_df8)

    # 第 13 題 (BETWEEN 數值範圍)：
    # 請找出實付金額 payment_amount 介於 100,000 到 200,000 之間（包含邊界） 的請款單。（提示：使用 .between()）
    factReim_df9 = (
        factReim_df.loc[ factReim_df["payment_amount"].between(100000, 200000,inclusive="both"), ["reimburse_id", "payment_amount", "applicant_id", "vendor_id"] ]
        .head(10)
    )
    # print(factReim_df9)

    # 第 14 題 (BETWEEN 日期範圍)：
    # 請篩選出採購日期 po_date 在 2025-01-01 到 2025-06-30 之間（2025上半年）的所有請款紀錄。
    factReim_df10 = (
        factReim_df.loc[ factReim_df["po_date"].between("2025-01-01", "2025-06-30"), ["reimburse_id", "po_date", "payment_amount", "applicant_id"] ]
        .head(10)
    )
    # print(factReim_df10)

    # 第 15 題 (LIKE 'keyword%')：
    # 請篩選出服務摘要 service_summary 中，開頭為 '維修' 或包含 '保養' 的請款單。（提示：使用 .str.contains() 或 .str.startswith()）
    factReim_df11 = (
        factReim_df.loc[ (factReim_df["service_summary"].str.startswith("維修")) | (factReim_df["service_summary"].str.contains("保養")),  ["reimburse_id", "po_date", "service_summary", "applicant_id"]  ]
    )
    # print(factReim_df11)

    # 第 16 題 (LIKE '%keyword%')：
    # 請找出發票號碼 invoice_no 中包含 '819' 字串的所有請款紀錄。
    factReim_df12 = (
        factReim_df.loc[ factReim_df["invoice_no"].str.contains("819"), ["reimburse_id", "po_date", "service_summary", "invoice_no"] ]
    )
    # print(factReim_df12)

    # 第 17 題 (LIKE 忽略大小寫)：
    # 請從 dim_applicants（申請人表）中，找出組別 user_group 包含 'fm' 的資料（需不區分大小寫，如 FM 或 fm 都算）。（提示：case=False）
    app_df2 = (
        app_df.loc[ app_df["user_group"].str.contains("fm", case=False) ]
        .head(10)
    )
    # print(app_df2)

    # 第 18 題 (CASE WHEN 單一條件 - np.where)：請在 fact_reimbursements 中新增一個欄位 is_large_amount：若 payment_amount >= 200,000 標記為 '大額'，否則標記為 '一般'。
    factReim_df14 = factReim_df.copy()
    factReim_df14["is_large_amount"] = np.where(factReim_df14["payment_amount"] >= 200000, "大額", "一般")
    factReim_df14 = (
        factReim_df14[["reimburse_id", "po_date", "payment_amount", "is_large_amount"]]
    )
    # print(factReim_df14)

    # 第 19 題 (CASE WHEN 多重條件 - np.select)：請新增一個欄位 amount_grade，依照 payment_amount 進行級距分類：>= 200,000 :'S', >= 100,000 且 < 200,000 :'A', < 100,000:'B'
    factReim_df15 = factReim_df.copy()
    conditions = [
        factReim_df15["payment_amount"] >= 200000,
        factReim_df15["payment_amount"] >= 100000,
        factReim_df15["payment_amount"] < 100000,
    ]
    choice = [
        "S", "A", "B"
    ]
    factReim_df15["amount_grade"] = np.select(conditions, choice, "no")
    factReim_df15 = factReim_df15[ ["reimburse_id", "po_date", "payment_amount", "amount_grade"] ]
    # print(factReim_df15)

    # 第 20 題 (綜合實戰：篩選 + 新增欄位 + 寫入新 Sheet)：請篩選出 2025 年下半年（po_date >= '2025-07-01'） 且 capex_expense == 'OPEX' 的紀錄，新增欄位 saving_status（若 saving_pct > 5 為 '高省利'，否則為 '常規'），並將結果寫入 simulate_output.xlsx 中的 OPEX下半年分析 Sheet 中！
    factReim_df16 = factReim_df.loc[ (factReim_df["po_date"]>= '2025-07-01') & (factReim_df["capex_expense"]== 'OPEX'), ["reimburse_id", "po_date", "payment_amount", "saving_pct", "capex_expense"] ]
    # factReim_df16["saving_status"] = np.where(factReim_df16["saving_pct"]>5, "高省利", "常規")
    # factReim_df16["saving_status"] = factReim_df16["saving_pct"].apply(lambda v: "高省利" if v > 5 else "常規")
    factReim_df16["saving_status"] = factReim_df16.apply(lambda df: "高省利" if df["saving_pct"] > 5 else "常規", axis=1 )

    print(factReim_df16)
    if os.path.exists(output_file):
        with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            factReim_df16.to_excel(writer, sheet_name="OPEX下半年分析", index=False)
            print("OPEX下半年分析已寫入")
    else:
        factReim_df16.to_excel(output_file, sheet_name="OPEX下半年分析", index=False)
        print("新excel已創建，OPEX下半年分析已寫入")
else:
    print(f"sheet is not present")




