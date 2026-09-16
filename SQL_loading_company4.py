# ==============================================================================
# 📌 1. 全域變數宣告與型態標註區 (Variable Declarations & Type Hints)
# 說明：集中宣告檔案路徑、資料庫連線字串、資料筆數與模擬參數
# ==============================================================================
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import pandas as pd
import numpy as np
import random
from sqlalchemy import create_engine

# 1.1 檔案與資料庫連線設定
excel_filename: str = 'simulate.xlsx'                           # Excel 落地檔案名稱
db_connection_url: str = "mysql+pymysql://root:90409991@localhost:3306/company_simulate"
# 💡 註：請將上面的 password、localhost、3306 與 company_db 替換為您的真實 SQL 連線資訊

# 1.2 模擬目標筆數與時間
reimbursements_count: int = 150                                 # 總表目標筆數 (150筆)
start_date: datetime = datetime(2025, 1, 1)                    # 模擬資料起始日期

# 1.3 外鍵參照與選單清單
gl_codes: List[str] = ["510101", "510102", "510201", "510301", "510401", "610101", "610201"]
capex_opex_list: List[str] = ["OPEX", "OPEX", "OPEX", "CAPEX"]   # 75% 費用化, 25% 資本化
services: List[str] = [
    "高低壓配電盤保養與定期檢測", "中央空調冰水主機維修與濾網更換", 
    "辦公大樓消防設施安全檢查", "電梯每月定期維護與零件更換", 
    "門禁系統軟體升級與卡機更換", "不斷電系統(UPS)電池組更換", 
    "樓層水管漏水修復與管線重敷", "數據中心機房冷氣保養", 
    "全行照明LED燈具更換專案", "發電機定期保養與試運轉"
]

# 1.4 資料儲存容器 (List of Dicts)
reimbursement_data: List[Dict[str, Any]] = []                  # 總表字典列表
allocation_data: List[Dict[str, Any]] = []                     # 分攤表字典列表

# 1.5 計數器與中間變數
alloc_id_counter: int = 1                                       # 分攤表 PK 流水號

# 1.6 DataFrame 變數宣告
df_dim_locations: pd.DataFrame
df_dim_organization: pd.DataFrame
df_map_bridge: pd.DataFrame
df_dim_vendors: pd.DataFrame
df_dim_applicants: pd.DataFrame
df_reimbursements: pd.DataFrame
df_allocations: pd.DataFrame


# ==============================================================================
# 🔄 2. 讀取維度表與資料準備 (Data Preparation)
# ==============================================================================
random.seed(42)
np.random.seed(42)

# 讀取既有維度表
df_dim_locations = pd.read_excel(excel_filename, sheet_name='bldg_dim_locations')
df_dim_organization = pd.read_excel(excel_filename, sheet_name='dim_organization')
df_map_bridge = pd.read_excel(excel_filename, sheet_name='map_org_location_bridge')
df_dim_vendors = pd.read_excel(excel_filename, sheet_name='dim_vendors')
df_dim_applicants = pd.read_excel(excel_filename, sheet_name='dim_applicants')

# 提取有效外鍵清單
applicant_ids: List[str] = df_dim_applicants['applicant_id'].tolist()
vendor_ids: List[int] = df_dim_vendors['vendor_id'].tolist()
bridge_pairs: List[Dict[str, Any]] = df_map_bridge[['location_key', 'pc_code']].to_dict('records')


# ==============================================================================
# ⚙️ 3. 核心 For 迴圈：生成總表與分攤表資料
# ==============================================================================
for r_id in range(1, reimbursements_count + 1):
    # 3.1 生成日期（採購 -> 送財務 -> 入帳）
    po_days_offset: int = random.randint(0, 360)
    po_date: datetime = start_date + timedelta(days=po_days_offset)
    to_fin_date: datetime = po_date + timedelta(days=random.randint(2, 10))
    gl_date: datetime = to_fin_date + timedelta(days=random.randint(1, 5))
    
    app_id: str = random.choice(applicant_ids)
    v_id: int = random.choice(vendor_ids)
    
    proj_no: Optional[str] = f"PRJ-2025-{random.randint(100, 999)}" if random.random() < 0.6 else None
    po_no: str = f"PO2025{random.randint(10000, 99999)}"
    inv_no: str = f"FU-{random.randint(10000000, 99999999)}"
    gl_code: str = random.choice(gl_codes)
    capex_opex: str = random.choice(capex_opex_list)
    summary: str = random.choice(services)
    
    # 3.2 算金額與議價折扣
    total_amount: float = round(random.uniform(10000, 500000), 2)
    saving_pct: float = round(random.uniform(0.0, 15.0), 2) if random.random() < 0.4 else 0.00
    payment_amount: float = round(total_amount * (1 - saving_pct / 100), 2)
    
    # 3.3 組裝總表列資料
    reimbursement_data.append({
        "reimburse_id": r_id,
        "po_date": po_date.strftime("%Y-%m-%d"),
        "to_finance_date": to_fin_date.strftime("%Y-%m-%d"),
        "gl_date": gl_date.strftime("%Y-%m-%d"),
        "applicant_id": app_id,
        "vendor_id": v_id,
        "project_no": proj_no,
        "po_no": po_no,
        "invoice_no": inv_no,
        "gl_code": gl_code,
        "capex_expense": capex_opex,
        "service_summary": summary,
        "total_amount_tax": total_amount,
        "payment_amount": payment_amount,
        "saving_pct": saving_pct
    })
    
    # 3.4 計算分攤明細 (1~4 個地點/部門拆分)
    num_splits: int = random.choices([1, 2, 3, 4], weights=[0.35, 0.35, 0.20, 0.10])[0]
    selected_bridges: List[Dict[str, Any]] = random.sample(bridge_pairs, num_splits)
    
    if num_splits == 1:
        pcts: List[float] = [100.00]
    else:
        raw_pcts = np.random.dirichlet(np.ones(num_splits)) * 100
        pcts = [round(p, 2) for p in raw_pcts]
        pcts[0] = round(pcts[0] + round(100.00 - sum(pcts), 2), 2)
        
    for i in range(num_splits):
        b = selected_bridges[i]
        pct: float = pcts[i]
        alloc_amt: float = round(total_amount * (pct / 100.0), 2)
        
        allocation_data.append({
            "allocation_id": alloc_id_counter,
            "reimburse_id": r_id,
            "location_key": b['location_key'],
            "pc_code": b['pc_code'],
            "allocated_amount": alloc_amt,
            "allocated_pct": pct,
            "remark": f"{summary} - 分攤 ({pct}%)"
        })
        alloc_id_counter += 1

# 轉成 DataFrame
df_reimbursements = pd.DataFrame(reimbursement_data)
df_allocations = pd.DataFrame(allocation_data)

# 3.5 財務精確度校正：修正四捨五入微小誤差，確保分攤金額加總完全等於總表金額
for r_id in range(1, reimbursements_count + 1):
    header_amt = df_reimbursements.loc[df_reimbursements['reimburse_id'] == r_id, 'total_amount_tax'].values[0]
    alloc_indices = df_allocations[df_allocations['reimburse_id'] == r_id].index
    sum_alloc = df_allocations.loc[alloc_indices, 'allocated_amount'].sum()
    diff = round(header_amt - sum_alloc, 2)
    if diff != 0:
        df_allocations.loc[alloc_indices[-1], 'allocated_amount'] = round(
            df_allocations.loc[alloc_indices[-1], 'allocated_amount'] + diff, 2
        )


# ==============================================================================
# 💾 4. 落地端 A：Excel 檔案追加寫入 (Mode = 'a')
# ==============================================================================
with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_reimbursements.to_excel(writer, sheet_name='fact_reimbursements', index=False)
    df_allocations.to_excel(writer, sheet_name='fact_reimbursement_allocations', index=False)

print(f"📊 [Excel 落地成功] 檔案 [{excel_filename}] 已成功更新：")
print(f"   └─ fact_reimbursements: {len(df_reimbursements)} 筆")
print(f"   └─ fact_reimbursement_allocations: {len(df_allocations)} 筆")


# ==============================================================================
# 🗄️ 5. 落地端 B：SQL 資料庫批次寫入 (SQLAlchemy)
# ==============================================================================
try:
    # 建立資料庫連線引擎
    engine = create_engine(db_connection_url)
    
    # ⚠️ 注意外鍵順序：必須先寫入總表 (Fact Reimbursements)
    df_reimbursements.to_sql(
        name='fact_reimbursements', 
        con=engine, 
        if_exists='append', 
        index=False,
        chunksize=500  # 批次寫入提升效能
    )
    print(f"🚀 [SQL 落地成功] 150 筆總表資料已寫入資料庫！")
    
    # 再寫入明細分攤表 (Fact Reimbursement Allocations)
    df_allocations.to_sql(
        name='fact_reimbursement_allocations', 
        con=engine, 
        if_exists='append', 
        index=False,
        chunksize=500
    )
    print(f"🚀 [SQL 落地成功] {len(df_allocations)} 筆分攤表資料已寫入資料庫！")

except Exception as e:
    print(f"\n💡 [SQL 連線提示] 資料已安全備份至 Excel。若要自動導入 SQL，請確認 db_connection_url 設定。錯誤細節: {e}")