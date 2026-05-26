import pandas as pd, json, re
from datetime import datetime, timedelta

today = datetime.today().date()

# ── Load files ──────────────────────────────────────────────────────────────
df_mem = pd.read_excel('/sessions/nice-eager-goodall/mnt/uploads/MemberInfo-2026-05-25 11_25_38.xlsx')
df_wl  = pd.read_excel('/sessions/nice-eager-goodall/mnt/uploads/All Members-2026-05-25 11_11_48.xlsx')
df_dep = pd.read_excel('/sessions/nice-eager-goodall/mnt/uploads/Deposit-2026-05-25 11_11_58.xlsx')
df_wd  = pd.read_excel('/sessions/nice-eager-goodall/mnt/uploads/Withdrawal-2026-05-25 11_42_09.xlsx')

# ── Amount parser ────────────────────────────────────────────────────────────
def parse_amount(s):
    if pd.isna(s): return 0.0
    m = re.search(r'[\d,]+\.?\d*', str(s).replace(',',''))
    return float(m.group().replace(',','')) if m else 0.0

# ── Win/loss map (use last column = winLoss) ─────────────────────────────────
last_col = df_wl.columns[-1]
print(f"Using win/loss column: '{last_col}'")
wl_map = df_wl.set_index('member')[[last_col]].copy()
wl_map.columns = ['winLoss']

# ── Process deposits ──────────────────────────────────────────────────────────
df_dep['ts']  = pd.to_datetime(df_dep['Created Timestamp'], errors='coerce')
df_dep['amt'] = df_dep['Amount'].apply(parse_amount)
df_dep['Login ID'] = df_dep['Login ID'].astype(str).str.strip()

# Approved deposits only for financial calcs
dep_ok = df_dep[df_dep['Status'] == 'Approved'].copy()

# All deposits (approved + rejected) for last deposit status
dep_all_sorted = df_dep.sort_values('ts')
last_dep_any = dep_all_sorted.groupby('Login ID').last()[['ts','Status','amt','Payment Method']]

# Date windows
feb_s, feb_e = datetime(2026,2,1), datetime(2026,2,28,23,59,59)
mar_s, mar_e = datetime(2026,3,1), datetime(2026,3,31,23,59,59)
apr_s, apr_e = datetime(2026,4,1), datetime(2026,4,30,23,59,59)
may_s, may_e = datetime(2026,5,1), datetime(2026,5,31,23,59,59)
week_s       = datetime.combine(today - timedelta(days=7), datetime.min.time())

def dep_window(lid, s, e):
    m = dep_ok[(dep_ok['Login ID']==lid) & (dep_ok['ts']>=s) & (dep_ok['ts']<=e)]
    return round(m['amt'].sum(),2), len(m)

# FTD per login (first approved deposit)
ftd_df = dep_ok.sort_values('ts').groupby('Login ID').first()[['ts','amt']]

# ── Process withdrawals ───────────────────────────────────────────────────────
df_wd['amt'] = df_wd['Amount'].apply(parse_amount)
df_wd['Login ID'] = df_wd['Login ID'].astype(str).str.strip()
wd_ok = df_wd[df_wd['Status'] == 'Approved'].copy()
wd_total = wd_ok.groupby('Login ID')['amt'].sum()

# ── Registration date parser ──────────────────────────────────────────────────
def parse_reg_date(reginfo):
    if pd.isna(reginfo): return ''
    m = re.match(r'(\d+/\d+/\d+)', str(reginfo))
    if not m: return ''
    try:
        return datetime.strptime(m.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')
    except:
        return ''

VIP_INACTIVE_DAYS = 30

# ── Build records ─────────────────────────────────────────────────────────────
records = []
for _, row in df_mem.iterrows():
    lid = str(row['loginId']).strip()
    mid = int(row['id']) if not pd.isna(row['id']) else 0

    # Registration date from registerInfo
    reg_date = parse_reg_date(row.get('registerInfo',''))

    membership = str(row.get('memberShip','Normal')).strip()
    is_vip = membership != 'Normal'

    # FTD
    ftd_date, ftd_amt = '', 0.0
    if lid in ftd_df.index:
        ftd_date = ftd_df.loc[lid,'ts'].strftime('%Y-%m-%d') if pd.notna(ftd_df.loc[lid,'ts']) else ''
        ftd_amt  = round(float(ftd_df.loc[lid,'amt']),2)

    # Monthly deposits
    feb_dep, feb_cnt = dep_window(lid, feb_s, feb_e)
    mar_dep, mar_cnt = dep_window(lid, mar_s, mar_e)
    apr_dep, apr_cnt = dep_window(lid, apr_s, apr_e)
    may_dep, may_cnt = dep_window(lid, may_s, may_e)

    # Weekly
    week_m = dep_ok[(dep_ok['Login ID']==lid) & (dep_ok['ts']>=week_s)]
    week_dep, week_cnt = round(week_m['amt'].sum(),2), len(week_m)

    # Lifetime
    life_m   = dep_ok[dep_ok['Login ID']==lid]
    life_dep = round(life_m['amt'].sum(),2)
    life_cnt = len(life_m)
    life_wd  = round(float(wd_total.get(lid, 0.0)),2)

    # Lifetime win/loss from winLoss column
    life_wl = 0.0
    if lid in wl_map.index:
        try: life_wl = round(float(wl_map.loc[lid,'winLoss']),2)
        except: pass

    # Monthly loss (deposit - withdrawal proxy using winLoss ratio — use dep - wd for simplicity)
    # For monthly loss: dep - (proportional wd) … simpler: just show deposit and use winLoss for lifetime
    # Monthly "loss" shown = deposits - net (approximation); keep as dep for now, actual loss not in monthly data
    # Use same approach as before: monthly loss = monthly dep * (life_wl / life_dep) ratio if available
    def monthly_loss(dep):
        if life_dep > 0 and life_wl != 0:
            return round(dep * (life_wl / life_dep), 2)
        return 0.0

    feb_loss = monthly_loss(feb_dep)
    mar_loss = monthly_loss(mar_dep)
    apr_loss = monthly_loss(apr_dep)
    may_loss = monthly_loss(may_dep)

    # Last deposit (all statuses)
    last_dep_date, last_dep_status, last_dep_ts, last_dep_method, last_dep_amt = '','','','',0.0
    if lid in last_dep_any.index:
        ld = last_dep_any.loc[lid]
        last_dep_ts     = ld['ts'].strftime('%Y-%m-%d') if pd.notna(ld['ts']) else ''
        last_dep_date   = last_dep_ts
        last_dep_status = str(ld['Status']) if pd.notna(ld['Status']) else ''
        last_dep_method = str(ld['Payment Method']) if pd.notna(ld['Payment Method']) else ''
        last_dep_amt    = round(float(ld['amt']),2)

    # VIP inactivity
    vip_inactive, vip_inactive_days = False, 0
    if is_vip and last_dep_ts:
        try:
            days_since = (today - datetime.strptime(last_dep_ts,'%Y-%m-%d').date()).days
            vip_inactive_days = days_since
            vip_inactive = days_since >= VIP_INACTIVE_DAYS
        except: pass

    # Alerts (Normal members only)
    alerts = []
    if not is_vip:
        if ftd_amt >= 1000:      alerts.append({'type':'FTD',  'label':'FTD ≥ 1K'})
        if week_dep >= 10000:    alerts.append({'type':'WEEK', 'label':'7D ≥ 10K'})
        if may_dep >= 10000:     alerts.append({'type':'MONTH','label':'30D ≥ 10K'})
        if life_cnt >= 20:       alerts.append({'type':'FREQ', 'label':f'Hi-Freq ({life_cnt}x)'})

    records.append({
        'id': mid, 'loginId': lid,
        'name': str(row.get('name','')).strip(),
        'regDate': reg_date,
        'membership': membership,
        'status': str(row.get('status','Active')).strip(),
        'ftd_date': ftd_date, 'ftd_amt': ftd_amt,
        'feb_dep': feb_dep, 'feb_cnt': feb_cnt, 'feb_loss': feb_loss,
        'mar_dep': mar_dep, 'mar_cnt': mar_cnt, 'mar_loss': mar_loss,
        'apr_dep': apr_dep, 'apr_cnt': apr_cnt, 'apr_loss': apr_loss,
        'may_dep': may_dep, 'may_cnt': may_cnt, 'may_loss': may_loss,
        'life_dep': life_dep, 'life_cnt': life_cnt, 'life_wd': life_wd,
        'life_winloss': life_wl,
        'week_dep': week_dep, 'week_cnt': week_cnt,
        'last_dep_date': last_dep_date,
        'last_dep_status': last_dep_status,
        'last_dep_ts': last_dep_ts,
        'last_dep_method': last_dep_method,
        'last_dep_amt': last_dep_amt,
        'vip_inactive': vip_inactive,
        'vip_inactive_days': vip_inactive_days,
        'alerts': alerts
    })

with open('/sessions/nice-eager-goodall/mnt/outputs/player_data2.json','w') as f:
    json.dump(records, f)

print(f"Done. {len(records)} records written.")
# Verify sample
sample = next((r for r in records if r['regDate']), None)
if sample:
    print(f"Sample regDate: {sample['loginId']} → {sample['regDate']}")
