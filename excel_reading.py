import pandas as pd
import numpy as np

# 1. pd.concat > 將同為dataframe的資料做上下串接
# file_path = "contracts_v50.xlsx"
# target_sheet = "工作表1"

# excel_file = pd.ExcelFile(file_path)
# if target_sheet in excel_file.sheet_names:
#         df_temp = pd.read_excel(excel_file, sheet_name=target_sheet)
#         # print(df_temp)
#         df_new = pd.DataFrame([
#                 ["dunk", 22],
#                 ["emma", 23]
#                 ], columns=df_temp.columns)
#         # print(df_new)
#         df_merge = pd.concat([df_temp,df_new], ignore_index=True)
#         print(df_merge)

#         with pd.ExcelWriter(file_path, engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
#                 df_merge.to_excel(writer, sheet_name=target_sheet, index=False)
        
#         print(f"成功更新 {target_sheet}！其他工作表均完好保留。")

# 2. .values.tolist() > 將讀取到的資料翻成二維list後做後續資料的增減

file_path = "contracts_v50.xlsx"
target_sheet = "工作表1"

excel_file = pd.ExcelFile(file_path)
if target_sheet in excel_file.sheet_names:
    temp_df = pd.read_excel(excel_file, sheet_name=target_sheet)
    tempt_list = temp_df.values.tolist()

    newData_list = [["frank", 24], ["gary", 25]]
    for elements in newData_list:
        tempt_list.append(elements)
    print(tempt_list)
    
    merge_df = pd.DataFrame(tempt_list)
    merge_df.columns = temp_df.columns
    print(merge_df)

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        merge_df.to_excel(writer, sheet_name=target_sheet, index=False)
    print(f"成功更新 {target_sheet}！其他工作表均完好保留。")

