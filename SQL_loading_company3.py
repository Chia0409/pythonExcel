# ==========================================
# 📌 變數宣告區（Variable Declarations & Type Hints）
# 說明：集中宣告本腳本使用的全域變數、資料型態與預設值
# ==========================================
from typing import List, Dict, Any
import pandas as pd
import random
from sqlalchemy import create_engine

# 1. 檔案與資料庫設定變數
excel_filename: str = 'simulate.xlsx'                           # 欲追加工作頁的 Excel 活頁簿檔名
db_connection_url: str = "mysql+pymysql://root:90409991@localhost:3306/company_simulate" # SQL 連線字串
vendor_target_count: int = 20                                   # 廠商資料目標筆數
applicant_target_count: int = 30                                 # 同仁資料目標筆數

# 2. 模擬數據清單 (List of Strings)
vendor_names: List[str] = [                                     # 20 家實體機電/工程/電信廠商名稱
    "台電工程股份有限公司", "大同綜合電訊", "中華電信股份有限公司", "遠傳電信", "台灣大哥大",
    "東元電機", "三菱電機台灣分公司", "施耐德電機", "飛利浦照明", "漢唐集成",
    "聖暉工程", "亞翔工程", "洋基工程", "中興電工", "華城電機",
    "士林電機", "三洋電機", "松下電器", "日立冷氣", "大金空調"
]

group_options: List[str] = ["FM", "PM", "Finance"]             # 同仁所屬組別選項列表

applicant_names: List[str] = [                                  # 30 位同仁真實姓名列表
    "陳志明", "林淑芬", "張建國", "黃麗華", "李偉婷", "王俊傑", "吳美玲", "劉家豪", "蔡佩君", "楊宗翰",
    "許雅婷", "鄭冠宇", "謝佳蓉", "郭哲瑋", "洪詩涵", "曾彥廷", "廖心怡", "賴柏翰", "邱怡君", "周宇軒",
    "葉雅文", "黃健豪", "莊惠雯", "江致遠", "簡妙如", "藍承恩", "魏嘉玲", "游博勝", "羅敏君", "柯威廷"
]

# 3. 資料處理中間容器 (List of Dicts)
vendors_data: List[Dict[str, Any]] = []                          # 存放廠商列資料的字典串列
applicants_data: List[Dict[str, Any]] = []                       # 存放同仁列資料的字典串列

# 4. Pandas DataFrame 物件
df_dim_vendors: pd.DataFrame                                    # 廠商維度表 DataFrame
df_dim_applicants: pd.DataFrame                                 # 同仁維度表 DataFrame


# ==========================================
# 1. 生成 dim_vendors (廠商維度表 - 20 筆)
# ==========================================
vendors_data = [
    {"vendor_id": idx, "vendor_name": name}
    for idx, name in enumerate(vendor_names, start=1)
]
df_dim_vendors = pd.DataFrame(vendors_data)


# ==========================================
# 2. 生成 dim_applicants (同仁維度表 - 30 筆)
# ==========================================
random.seed(42)  # 固定隨機種子，確保多次執行結果一致
for idx in range(1, applicant_target_count + 1):
    app_id: str = f"EMP{idx:03d}"                               # 生成格式化工號 (如 EMP001, EMP002)
    app_name: str = applicant_names[idx - 1]
    grp: str = group_options[(idx - 1) % len(group_options)]     # 均勻輪流分配組別 (FM, PM, Finance)
    
    applicants_data.append({
        "applicant_id": app_id,
        "applicant_name": app_name,
        "user_group": grp
    })

df_dim_applicants = pd.DataFrame(applicants_data)


# ==========================================
# 3. 落地作業 A：使用追加模式 (mode='a') 寫入 Excel
# ==========================================
# 🎯 關鍵優化：利用 openpyxl 引擎追加頁籤，不必重新讀取舊 DataFrame
with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_dim_vendors.to_excel(writer, sheet_name='dim_vendors', index=False)
    df_dim_applicants.to_excel(writer, sheet_name='dim_applicants', index=False)

print(f"✅ Excel 追加成功！檔案 [{excel_filename}] 已補齊 dim_vendors 與 dim_applicants 頁籤。")


# ==========================================
# 4. 落地作業 B：寫入 SQL 資料庫
# ==========================================
try:
    engine = create_engine(db_connection_url)
    
    df_dim_vendors.to_sql(name='dim_vendors', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] dim_vendors 20 筆資料已成功寫入資料庫！")
    
    df_dim_applicants.to_sql(name='dim_applicants', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] dim_applicants 30 筆資料已成功寫入資料庫！")

except Exception as e:
    print(f"💡 提醒：資料已安全寫入 Excel 活頁簿。SQL 寫入提示（若未開起資料庫連線可忽略）: {e}")