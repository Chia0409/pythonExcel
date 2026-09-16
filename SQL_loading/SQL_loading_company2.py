import pandas as pd
import numpy as np
import random
from sqlalchemy import create_engine

# ==========================================
# 0. 基礎設定與隨機種子 (確保結果可複現)
# ==========================================
random.seed(42)
np.random.seed(42)

excel_filename = 'simulate.xlsx'

# 讀取上一階段已建置好的大樓主表與地點維度表
df_bldg_master = pd.read_excel(excel_filename, sheet_name='bldg_master')
df_dim_locations = pd.read_excel(excel_filename, sheet_name='bldg_dim_locations')

# ==========================================
# 1. 生成 dim_organization (組織維度表 - 100 筆)
# ==========================================
# 包含之前拆彈核心 PC Code，其餘自動補足至 100 筆
special_pc_codes = ["C669", "C631", "C634", "45CB", "88TD", "00TD"]
pc_codes = special_pc_codes.copy()
i = 1
while len(pc_codes) < 100:
    code = f"PC{i:03d}"
    if code not in pc_codes:
        pc_codes.append(code)
    i += 1

# 固定的 12 個 BUSU 核心值
busu_list = ['CBG', 'IBG', 'T&O', 'Leg', 'Audit', 'HQ', 'IT', 'Media', 'TDSP', 'GTS', 'CRESA', 'GPS']

org_data = []
for idx, code in enumerate(pc_codes):
    busu = busu_list[idx % len(busu_list)]
    
    # 🎯 業務邏輯：只有 CBG & T&O 有次級單位，其餘填 None (SQL NULL)
    if busu == 'CBG':
        sub_unit = f"CBG-{(idx % 4) + 1}"
    elif busu == 'T&O':
        sub_unit = f"T&O-{(idx % 4) + 1}"
    else:
        sub_unit = None
        
    dept_name = f"Department-{idx + 1}"
    status = 'Active' if (idx % 10 != 0) else 'Closed' # 90% Active / 10% Closed
    
    org_data.append({
        "pc_code": code,
        "busu": busu,
        "busu_sub_unit": sub_unit,
        "department_name": dept_name,
        "status": status
    })

df_dim_organization = pd.DataFrame(org_data)

# ==========================================
# 2. 生成 map_org_location_bridge (橋接表 - 200 筆)
# ==========================================
# 🎯 約束過濾：只篩選地點表中狀態為 '自用' 或 '承租' 的地點
valid_locations = df_dim_locations[df_dim_locations['status'].isin(['自用', '承租'])][['location_key', 'status']].to_dict('records')

bridge_data = []
unique_pairs = set() # 用於保證 uq_pc_location (pc_code, location_key) 唯一性

bridge_id = 1
while len(bridge_data) < 200:
    pc = random.choice(pc_codes)
    loc_obj = random.choice(valid_locations)
    loc_key = loc_obj['location_key']
    loc_status = loc_obj['status'] # 100% 綁定真實地點狀態
    
    pair = (pc, loc_key)
    if pair in unique_pairs:
        continue # 若重複則重新抽樣
        
    unique_pairs.add(pair)
    
    is_primary = 'Y' if random.random() < 0.3 else 'N'
    allocation_ratio = random.choice([100.00, 50.00, 25.00, 75.00, 33.33])
    
    bridge_data.append({
        "bridge_id": bridge_id,
        "pc_code": pc,
        "location_key": loc_key,
        "location_status": loc_status,
        "is_primary_site": is_primary,
        "allocation_ratio": allocation_ratio
    })
    bridge_id += 1

df_map_bridge = pd.DataFrame(bridge_data)

# ==========================================
# 3. 落地作業 A：更新寫入 Excel 檔 (simulate.xlsx)
# ==========================================
with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
    df_bldg_master.to_excel(writer, sheet_name='bldg_master', index=False)
    df_dim_locations.to_excel(writer, sheet_name='bldg_dim_locations', index=False)
    df_dim_organization.to_excel(writer, sheet_name='dim_organization', index=False)
    df_map_bridge.to_excel(writer, sheet_name='map_org_location_bridge', index=False)

print(f"✅ Excel 檔案更新成功：{excel_filename}")
print(f"   - dim_organization 工作頁：{len(df_dim_organization)} 筆")
print(f"   - map_org_location_bridge 工作頁：{len(df_map_bridge)} 筆")

# ==========================================
# 4. 落地作業 B：順序匯入 SQL 資料庫
# ==========================================
db_connection_url = "mysql+pymysql://root:90409991@localhost:3306/company_simulate"

try:
    engine = create_engine(db_connection_url)
    
    # ⚠️ 外鍵相依順序：先寫入組織主表，再寫入橋接表
    df_dim_organization.to_sql(name='dim_organization', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] dim_organization 資料成功寫入！")
    
    df_map_bridge.to_sql(name='map_org_location_bridge', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] map_org_location_bridge 資料成功寫入！")

except Exception as e:
    print(f"💡 提示：若未開起 SQL 連線，請替換連線字串。模擬資料已安全寫入 Excel！\n訊息: {e}")