import numpy as np
import pandas as pd

def excel_sheet_loader(excel_path_input:str, target_sheet_input:list):
    excel_file = pd.ExcelFile(excel_path_input)
    
    if all(sheet in excel_file.sheet_names for sheet in target_sheet_input):
        sheet_count = 0
        data_dict = {}
        df_name_list = []

        for sheet in target_sheet_input:
            if sheet in excel_file.sheet_names:
                df_key = f"{sheet}_df"
                data_dict[df_key] = pd.read_excel(excel_path_input, sheet_name=sheet)
                sheet_count += 1
                df_name_list.append(df_key)
    
        # print(f"sheet_count: {sheet_count}, df_name_list: {df_name_list}")
        return data_dict
    
    else:
        miss_list = [sheet for sheet in target_sheet_input if sheet not in excel_file.sheet_names] 
        raise FileNotFoundError(f"缺少以下工作表：{miss_list}")
    

# ===========================testing=============================================

if __name__ == "__main__" :
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

    fact_reimbursements_df = dfs["fact_reimbursements_df"]
    print(fact_reimbursements_df.head(3))
    

        



