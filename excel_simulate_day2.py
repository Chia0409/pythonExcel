import os
import numpy as np
import pandas as pd

input_file = "simulate.xlsx"
out_file = "simulate_output.xlsx"

targetsheet = ["fact_reimbursements", "fact_reimbursement_allocations", "dim_applicants", "dim_vendors", "dim_organization", "map_org_location_bridge", "bldg_dim_locations", "bldg_master", "dim_date"]
excel_file = pd.ExcelFile(input_file)
# sheet_in_list = []
# for sheet in targetsheet:
#     exist = sheet in targetsheet
#     sheet_in_list.append(exist)
# if all(sheet_in_list):
#     print("所有工作表都存在")

if all(sheet in excel_file.sheet_names for sheet in targetsheet):
    
    fact_reimbursements_df = pd.read_excel(excel_file, sheet_name="fact_reimbursements")
    fact_reimbursement_allocations_df = pd.read_excel(excel_file, sheet_name="fact_reimbursement_allocations")
    dim_applicants_df= pd.read_excel(excel_file, sheet_name="dim_applicants")
    dim_vendors_df = pd.read_excel(excel_file, sheet_name="dim_vendors")
    dim_organization_df = pd.read_excel(excel_file, sheet_name="dim_organization")
    bldg_dim_locations_df = pd.read_excel(excel_file, sheet_name="bldg_dim_locations") 
    bldg_master_df = pd.read_excel(excel_file, sheet_name="bldg_master")
    map_org_location_bridge_df = pd.read_excel(excel_file, sheet_name="map_org_location_bridge")
    dim_date_df = pd.read_excel(excel_file, sheet_name="dim_date")

# 第 1 題 (雙表 LEFT JOIN)：
# 請將請款主表 fact_reimbursements 與申請人表 dim_applicants 進行 LEFT JOIN（關鍵字：applicant_id），並選取 reimburse_id、applicant_name 與 user_group 前 5 筆資料。
    fr_app_df = pd.merge(fact_reimbursements_df, dim_applicants_df, how="left", on="applicant_id")
    # print(fr_app_df)

# 第 2 題 (三表連環 JOIN)：
# 請將 fact_reimbursements 串接 dim_applicants（取 applicant_name）以及 dim_vendors（取 vendor_name），並只顯示 reimburse_id、payment_amount、applicant_name 與 vendor_name。
    fr_app_v_df = pd.merge(fr_app_df, dim_vendors_df, how="left", on="vendor_id")
    fr_app_v_df2 = fr_app_v_df[["reimburse_id","payment_amount","applicant_name" ,"vendor_name"]]
    # print(fr_app_v_df2)

# 第 3 題 (異名鍵 JOIN - left_on / right_on)：
# 假設 dim_vendors 的欄位名稱改為 supplier_id，請使用 left_on='vendor_id' 與 right_on='supplier_id' 將其與 fact_reimbursements 串接，並在串接後丟棄多餘的 supplier_id 欄位。
    dim_vendors_df2 = dim_vendors_df.copy()
    dim_vendors_df2 = dim_vendors_df2.rename(columns={"vendor_id":"supplier_id"})
    fr_v_df = pd.merge(fact_reimbursements_df, dim_vendors_df2, how="left", left_on="vendor_id",right_on="supplier_id")
    # print(fr_v_df)


# 第 4 題 (INNER JOIN 數據過濾)：
# 請將分攤表 fact_reimbursement_allocations 與組織表 dim_organization 進行 INNER JOIN（關鍵字：pc_code），只保留能夠完美對應到的組織分攤紀錄。
    fra_org_df = pd.merge(fact_reimbursement_allocations_df, dim_organization_df, how="inner", on="pc_code")
    # print(fra_org_df)

# 第 5 題 (多對多橋接表 4 表 JOIN)：
# 請從分攤表 fact_reimbursement_allocations 出發，依序串接 map_org_location_bridge（依 pc_code）、bldg_dim_locations（依 location_key）與 bldg_master（依 building_code），找出每一筆分攤金額對應的繁體大樓名稱 building_name_cn。
    # step1 = pd.merge(fact_reimbursement_allocations_df, map_org_location_bridge_df, how="left", on="pc_code", suffixes=("", "_bridge"))
    step1_df = pd.merge(fact_reimbursement_allocations_df, map_org_location_bridge_df, how="left", on=["pc_code", "location_key"])
    step2_df = pd.merge(step1_df, bldg_dim_locations_df, how="left", on="location_key")
    step3_df = pd.merge(step2_df, bldg_master_df, how="left", on="building_code")
    final_df = step3_df[["allocation_id", "reimburse_id", "location_key", "pc_code", "building_name_cn"]]
    # print(final_df)

# 第 6 題 (isnull 檢查與統計)：
# 請檢查 fact_reimbursements 中每一個欄位的空值數量（提示：使用 .isnull().sum()），並印出有空值出現的欄位清單。
    fr_vaCount = fact_reimbursements_df.isnull().sum()
    # print(fr_vaCount)
    fr_vaCount_list1 = [col for col in fr_vaCount if col != 0 ]
    fr_vaCount_list2 = [col for col, val in fr_vaCount.items() if val > 0]
    # print(fr_vaCount_list2)


# 第 7 題 (fillna 指定值補全)：
# fact_reimbursements 中的專案編號 project_no 有許多空值。請使用 .fillna('無專案編號') 填補該欄位的空值，並覆蓋原欄位或新增欄位。
    fa2 = fact_reimbursements_df.copy()
    fa2["project_no"] = fa2["project_no"].fillna("無專案編號")
    # print(fa2[["reimburse_id", "project_no"]])

# 第 8 題 (fillna 跨欄位或統計值補全)：
# 請計算 fact_reimbursements 中 saving_pct 的平均值（Mean），並將 saving_pct 欄位中的空值用該平均值進行填補。
    fa3 = fact_reimbursements_df.copy()
    saving_pct_M_value = fa3["saving_pct"].mean(axis=0)
    # print(saving_pct_M_value)
    fa3["saving_pct"] = fa3["saving_pct"].apply(lambda x: saving_pct_M_value if x== 0.00 else x)
    # print(fa3[["reimburse_id", "saving_pct"]])

# 第 9 題 (fillna 前後值補全 method='ffill'/'bfill')：
# 在 dim_organization 中，busu_sub_unit 存在許多空值。請練習使用前值補全（method='ffill' 或 .ffill()）來模擬向下填補部門資料。
    org2_df = dim_organization_df.copy()
    # 在 Pandas 中，前值補全（向下填補上一個非空值）可以使用 .ffill() 方法，或是 .fillna(method='ffill')。
    org2_df["busu_sub_unit"] = org2_df["busu_sub_unit"].fillna(method="ffill")
    # print(org2_df[["pc_code","busu_sub_unit"]])

# 第 10 題 (dropna 刪除空值列)：
# 請剔除 fact_reimbursements 中 project_no 為空值 (NaN) 的紀錄，並比對剔除前後的資料筆數差異（提示：使用 len()）。
    fa4 = fact_reimbursements_df.copy()
    # print(f"fact_reimbursements_df: {len(fact_reimbursements_df)}")
    # 正確作法：指定 subset，只剔除 project_no 為 NaN 的列，但保留整張 DataFrame
    fa4_cleaned = fa4.dropna(subset=["project_no"])
    # print(f"fa4: {len(fa4)}")

# 第 11 題 (dropna 複合條件刪除)：
# 請刪除 fact_reimbursements 中 project_no 與 saving_pct 同時或任一為空值的列（提示：觀察 subset 參數）。
    fa5 = fact_reimbursements_df.copy()

    # how='any'：只要 subset 指定的欄位中「任一」出現 NaN 就刪除（預設值）
    # how='all'：必須 subset 指定的欄位「全部」同時都是 NaN 才刪除
    fa5_cleaned = fa5.dropna(
        subset=["project_no", "saving_pct"], how="any"
    )
    # print(f"原本筆數: {len(fa5)}")
    # print(f"刪除空值後筆數: {len(fa5_cleaned)}")


# 第 12 題 (pd.concat 垂直拼接 UNION ALL)：
# 請將 fact_reimbursements 拆分為 capex_expense == 'CAPEX' 與 capex_expense == 'OPEX' 兩張子表，最後再使用 pd.concat([df1, df2], axis=0) 將兩表垂直拼接回來，並重新編排索引 (ignore_index=True)。
    fr_cap_df = (
        fact_reimbursements_df.loc[ fact_reimbursements_df["capex_expense"]== 'CAPEX', ["reimburse_id", "payment_amount", "capex_expense"] ]
        .head(10)
    )
    
    fr_ope_df = (
        fact_reimbursements_df.loc[ fact_reimbursements_df["capex_expense"]== 'OPEX', ["reimburse_id", "payment_amount", "capex_expense"] ]
        .head(10)
    )
    cap_ope_df = pd.concat([fr_cap_df, fr_ope_df], axis=0, ignore_index=True)
    # print(cap_ope_df)

# 第 13 題 (pd.concat 帶有來源標籤 keys)：
# 接續第 12 題，在拼接時加入 keys=['CAPEX資料', 'OPEX資料'] 參數，觀察產生的多重索引（MultiIndex）結構。
    # 不清楚 第 13 題 使用keys的用意與方法
    cap_ope_df2 = pd.concat([fr_cap_df, fr_ope_df], axis=0, keys=['CAPEX資料', 'OPEX資料'])
    # print(cap_ope_df2)

# 第 14 題 (pd.concat 水平拼接)：
# 請將 fact_reimbursements 的前 10 筆資料拆成「基本資訊 (reimburse_id, po_date)」與「金額資訊 (payment_amount, saving_pct)」兩張表，再使用 pd.concat([df_base, df_amount], axis=1) 水平拼合。
    fr_cap_df2 = (
        fact_reimbursements_df[["reimburse_id", "po_date"]]
        .head(10)
    )
    fr_ope_df = (
        fact_reimbursements_df[["payment_amount", "saving_pct"]]
        .head(10)
    )
    cap_ope_df3 = pd.concat([fr_cap_df2, fr_ope_df], axis=1)
    # print(cap_ope_df3)

# 第 15 題 (批流目錄合併模組)：
# 假設資料夾中有 3 張結構相同的月份報表清單，請寫一個 for 迴圈搭配列表推導式（List Comprehension）與 pd.concat() 將它們合併成一張完整的大表。
    column_list = ["week","payment_amount","saving_pct"]
    Jan_df = pd.DataFrame([
        ["one", 100000, 10],
        ["two", 220000, 0],
        ["three", 120000, 4],
        ["four", 320000, 3]
    ], columns=column_list)
    Feb_df = pd.DataFrame([
        ["one", 120000, 10],
        ["two", 210000, 0],
        ["three", 150000, 4],
        ["four", 325000, 3]
    ], columns=column_list)
    Mar_df = pd.DataFrame([
        ["one", 140000, 10],
        ["two", 410000, 0],
        ["three", 50000, 4],
        ["four", 300000, 3]
    ], columns=column_list)
    Month_list = [Jan_df, Feb_df, Mar_df]
    month_df = pd.concat([month for month in Month_list], axis=0)
    # print(month_df)

# 第 16 題 (JOIN 缺漏值與分組統計)：
# 請將 fact_reimbursements LEFT JOIN dim_applicants 後，統計每個組別 user_group 的總請款金額 payment_amount（若串接後 user_group 為空值，請先填補為 '未指定部門'）。
    fa_app_df = pd.merge(fact_reimbursements_df, dim_applicants_df, how="left", on="applicant_id")
    fa_app_df["user_group"]= fa_app_df["user_group"].fillna("未指定部門")
    fa_app_df = fa_app_df[["reimburse_id", "applicant_id" ,"user_group", "payment_amount"]]
    fa_app_df2 = fa_app_df.groupby("user_group")["payment_amount"].sum().reset_index()
    # print(fa_app_df2)

# 第 17 題 (日曆表與請款單關聯分析)：
# 請將 fact_reimbursements（以 po_date）與 dim_date（以 calendar_date 轉字串或日期型態）進行 JOIN，篩選出「工作日 (is_workday == 'Y')」 的請款總金額。
    # 如何轉換日期，以讓兩表相連接？
    fr17 = fact_reimbursements_df.copy()
    dim_date17 = dim_date_df.copy()

    # 1. 統一轉為 datetime 型態 (Pandas 最推薦作法)
    fr17["po_date"] = pd.to_datetime(fr17["po_date"])
    dim_date17["calendar_date"] = pd.to_datetime(dim_date17["calendar_date"])

    # 2. 異名鍵串接 (po_date 對應 calendar_date)
    fa_d_df = pd.merge(
        fr17, dim_date17, how="left", left_on="po_date", right_on="calendar_date"
    )

    # 3. 篩選工作日 (is_workday == 'Y') 並計算總金額
    workday_total_amount = fa_d_df.loc[
        fa_d_df["is_workday"] == "Y", "payment_amount"
    ].sum()

    print(f"工作日請款總金額: {workday_total_amount:,.2f}")

# 第 18 題 (JOIN 數據驗證 - 避免笛卡爾積爆炸)：
# 請比較 fact_reimbursements 原始列數與串接 fact_reimbursement_allocations 後的列數，說明為何列數會大幅增加（提示：一對多關係）。
    # print(len(fact_reimbursements_df))
    fr_fra_df = pd.merge(fact_reimbursements_df, fact_reimbursement_allocations_df, how="left", on="reimburse_id")
    # print(len(fr_fra_df))
    # 因為每筆fact_reimbursements_df的資料列都至少會對到fact_reimbursement_allocations的一筆資料，有的是多筆

# 第 19 題 (清理+關聯+匯出Excel)：
# 請將 fact_reimbursements 中的 project_no 空值補為 'N/A'，接著 LEFT JOIN dim_vendors 取得廠商名稱 vendor_name，最後篩選 payment_amount >= 150000 的資料匯出至 simulate_output.xlsx 的 大額廠商請款 Sheet。
    # fr_df = fact_reimbursements_df.copy()
    # fr_df["project_no"] = fr_df["project_no"].fillna("N/A")
    # fr2_v_df = pd.merge(fr_df, dim_vendors_df, how="left", on="vendor_id")
    # fr2_v_df = fr2_v_df.loc[ fr2_v_df["payment_amount"]>= 150000]
    # print(fr2_v_df)
    # if os.path.exists(out_file):
    #     with pd.ExcelWriter(out_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    #         fr2_v_df.to_excel(writer, sheet_name="大額廠商請款")
    # else:
    #     fr2_v_df.to_excel(out_file, sheet_name="大額廠商請款")

# 第 20 題 (全資料鏈串接 - 總表生成挑戰)：
# 請將 fact_reimbursement_allocations（分攤表）作為主體，同時串接 fact_reimbursements（請款主表）、dim_applicants（申請人表）、dim_vendors（廠商表），產出一張包含 reimburse_id、applicant_name、vendor_name、allocated_amount 的完整大表！
    fr_df1 = pd.merge(fact_reimbursement_allocations_df, fact_reimbursements_df, how="left", on="reimburse_id")
    fr_df2 = pd.merge(fr_df1, dim_applicants_df, how="left", on="applicant_id")
    fr_final_df = pd.merge(fr_df2, dim_vendors_df, how="left", on="vendor_id")
    fr_final_df = fr_final_df[[ "reimburse_id", "applicant_name", "vendor_name", "allocated_amount"]]
    # print(fr_final_df)

else:
    print( list(set(excel_file.sheet_names) - set(targetsheet) ) )
    print([sheet for sheet in targetsheet if sheet not in excel_file.sheet_names])
    not_exist_list = []
    for sheet in targetsheet:
        if sheet not in excel_file.sheet_names:
            not_exist_sheet = sheet
            not_exist_list.append(sheet)
    print(not_exist_list)