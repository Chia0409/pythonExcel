import pandas as pd
import numpy as np
import datetime
import random
from sqlalchemy import create_engine

# ==========================================
# 0. 基礎設定與固定隨機種子 (確保資料可重複驗證)
# ==========================================
random.seed(42)
np.random.seed(42)

# 🎯 包含妳之前核心使用的實體大樓代碼，並補充至 60 筆
special_codes = ["00QH", "88QO", "8P13", "00QG", "8Q34", "18QB", "18QC"]
bldg_codes = special_codes.copy()
i = 1
while len(bldg_codes) < 60:
    code = f"B{i:03d}K"
    if code not in bldg_codes:
        bldg_codes.append(code)
    i += 1

# ==========================================
# 1. 生成 bldg_master (大樓主表 - 60 筆)
# ==========================================
cities_districts = [
    ("台北市中山區", "Zhongshan Dist., Taipei City"),
    ("台北市信義區", "Xinyi Dist., Taipei City"),
    ("台北市大安區", "Da'an Dist., Taipei City"),
    ("台北市中正區", "Zhongzheng Dist., Taipei City"),
    ("新北市板橋區", "Banqiao Dist., New Taipei City"),
    ("台中市西屯區", "Xitun Dist., Taichung City"),
    ("高雄市苓雅區", "Lingya Dist., Kaohsiung City"),
    ("桃園市桃園區", "Taoyuan Dist., Taoyuan City")
]

road_names_cn = ["中山北路", "南京東路", "忠孝東路", "敦化南路", "民生東路", "仁愛路", "復興南路", "中正路"]
road_names_eng = ["Zhongshan N. Rd.", "Nanjing E. Rd.", "Zhongxiao E. Rd.", "Dunhua S. Rd.", "Minsheng E. Rd.", "Ren'ai Rd.", "Fuxing S. Rd.", "Zhongzheng Rd."]
bldg_cn_prefixes = ["港都", "安和", "松江", "板橋", "信義", "南京", "敦南", "台中", "高雄", "桃興"]
bldg_cn_suffixes = ["金融大樓", "資訊大樓", "營運中心", "資訊中心", "備援大樓", "分行大樓", "企業大樓"]

bldg_master_data = []
for idx, code in enumerate(bldg_codes):
    cd_idx = idx % len(cities_districts)
    city_cn, city_eng = cities_districts[cd_idx]
    road_idx = idx % len(road_names_cn)
    sec_num = (idx % 5) + 1
    lane_num = (idx * 3 % 80) + 10
    no_num = (idx * 7 % 300) + 1
    
    addr_cn = f"{city_cn}{road_names_cn[road_idx]}{sec_num}段{lane_num}巷{no_num}號"
    addr_eng = f"No. {no_num}, Ln. {lane_num}, Sec. {sec_num}, {road_names_eng[road_idx]}, {city_eng}"
    
    prefix = bldg_cn_prefixes[idx % len(bldg_cn_prefixes)]
    suffix = bldg_cn_suffixes[idx % len(bldg_cn_suffixes)]
    name_cn = f"{prefix}{suffix}"
    name_eng = f"{prefix} Building {idx+1:02d}"
    
    bldg_master_data.append({
        "building_code": code,
        "address": addr_cn,
        "building_name_cn": name_cn,
        "building_name_eng": name_eng
    })

df_bldg_master = pd.DataFrame(bldg_master_data)

# ==========================================
# 2. 生成 bldg_dim_locations (大樓維度表 - 130 筆)
# ==========================================
floors = ["1樓", "2樓", "2樓之1", "2樓之2", "3樓", "4樓", "5樓", "B1", "全部"]
statuses = ["自用", "出租", "承租", "已售出", "暫時停用"]
snapshot_dates = {
    2025: datetime.date(2025, 12, 31),
    2026: datetime.date(2026, 1, 1),
    2027: datetime.date(2027, 1, 1)
}

dim_locations_data = []
unique_combos = set() # 確保符合 UNIQUE (version_year, building_code, floor_area)
location_key = 1

while len(dim_locations_data) < 130:
    # 比例配比：讓 2026 與 2027 年份有充足快照切片
    year = 2026 if len(dim_locations_data) < 60 else (2027 if len(dim_locations_data) < 115 else 2025)
    snap_date = snapshot_dates[year]
    b_code = random.choice(bldg_codes)
    flr = random.choice(floors)
    
    combo = (year, b_code, flr)
    if combo in unique_combos:
        continue # 若違反複合唯一約束則重新抽樣
    
    unique_combos.add(combo)
    st = random.choice(statuses)
    
    dim_locations_data.append({
        "location_key": location_key,
        "version_year": year,
        "snapshot_date": snap_date,
        "building_code": b_code,
        "floor_area": flr,
        "status": st
    })
    location_key += 1

df_dim_locations = pd.DataFrame(dim_locations_data)

# ==========================================
# 3. 落地作業 A：寫入 Excel 檔 (simulate.xlsx)
# ==========================================
excel_filename = "simulate.xlsx"
with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
    df_bldg_master.to_excel(writer, sheet_name='bldg_master', index=False)
    df_dim_locations.to_excel(writer, sheet_name='bldg_dim_locations', index=False)

print(f"✅ Excel 檔案產出成功：{excel_filename}")
print(f"   - bldg_master 工作頁：{len(df_bldg_master)} 筆")
print(f"   - bldg_dim_locations 工作頁：{len(df_dim_locations)} 筆")

# ==========================================
# 4. 落地作業 B：寫入 SQL 資料庫 (MySQL / PostgreSQL)
# ==========================================
# 🎯 請將此處修改為妳的 SQL 資料庫連線字串 (如 MySQL)
# 格式: mysql+pymysql://帳號:密碼@localhost:3306/company_simulate
db_connection_url = "mysql+pymysql://root:90409991@localhost:3306/company_simulate"

try:
    engine = create_engine(db_connection_url)
    
    # ⚠️ 順序極度重要：外鍵約束 (FK) 規定必須先寫入主表 (Master)，再寫入維度表 (Dim)
    df_bldg_master.to_sql(name='bldg_master', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] bldg_master 資料成功寫入資料庫！")
    
    df_dim_locations.to_sql(name='bldg_dim_locations', con=engine, if_exists='append', index=False)
    print("🚀 [SQL] bldg_dim_locations 資料成功寫入資料庫！")

except Exception as e:
    print(f"💡 提醒：若尚未連線真實 SQL 資料庫，請更新 db_connection_url 密碼與設定。程式已產出 Excel 備份！\n錯誤訊息: {e}")