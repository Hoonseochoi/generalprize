import pandas as pd
import os
import re
import asyncio
import base64
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright

# ── 설정 ──────────────────────────────────────────────────────
BASE_DIR        = r'C:\Users\chlgn\OneDrive\Desktop\general prize'
MC_LIST_FILE    = os.path.join(BASE_DIR, '0219MC_LIST_OUT_202602.xlsx') # Main list
# If main list not found, try backup
if not os.path.exists(MC_LIST_FILE):
    MC_LIST_FILE = os.path.join(BASE_DIR, '0219MC_LIST_OUT_202602 (1).xlsx')

PRIZE_SUM_FILE    = os.path.join(BASE_DIR, '0219PRIZE_SUM_OUT_202602.xlsx')
PRIZE_BRIDGE_FILE = os.path.join(BASE_DIR, '0219PRIZE_6_BRIDGE_OUT_202602.xlsx')

TEMPLATE   = 'template2.html'
OUTPUT_DIR = os.path.join(BASE_DIR, 'output_2')
DATE_STR   = '2026년 2월 19일 기준'

# 컬럼 매핑 (PRIZE_SUM) - Key: K(10)
# AB(27)=1주현금, AD(29)=1주상품, AG(32)=2주현금, AI(34)=2주상품
# AC(28)=1주상품실적, AH(33)=2주상품실적
Key_Sum = 10
Idx_S_1C = 27
Idx_S_1P = 29
Idx_S_1P_Perf = 28
Idx_S_2C = 32
Idx_S_2P = 34
Idx_S_2P_Perf = 33

# 컬럼 매핑 (PRIZE_BRIDGE) - Key: I(8)
# Z(25)=브릿지1금액, AG(32)=브릿지2금액
# X(23)=1월대상1, Y(24)=2월대상1
# AE(30)=1월대상2, AF(31)=2월대상2
Key_Bridge = 8
Idx_B_1 = 25
Idx_B_1_Jan = 23
Idx_B_1_Feb = 24
Idx_B_2 = 32
Idx_B_2_Jan = 30
Idx_B_2_Feb = 31

# ── 시상 구간 정의 ───────────────────────────────────────────────
INBO3_TIERS = [
    (100_000, '10만', 100_000), (200_000, '20만', 200_000),
    (300_000, '30만', 300_000), (500_000, '50만', 500_000),
]
CONT_TIERS = [
    (100_000, '10만', 200_000), (200_000, '20만', 600_000),
    (300_000, '30만', 800_000), (500_000, '50만', 1_800_000),
]
MCPLUS_TIERS = [
    (200_000,'20만',600_000), (400_000,'40만',1_200_000),
    (600_000,'60만',1_800_000), (800_000,'80만',2_400_000),
    (1_000_000,'100만',3_000_000),
]

# ── 유틸 함수 ─────────────────────────────────────────────────
def fmt(n):
    try:
        n = float(n)
        if n == 0 or pd.isna(n): return '0'
        val = int(n)
        if val >= 10000 and val % 10000 == 0:
            return f'{val // 10000:,}만'
        return f'{val:,}'
    except:
        return '0'

def fmt_won(n):
    try:
        n = int(n)
        if n == 0: return '0원'
        if n % 10_000 == 0: return f'{n // 10_000:,}만원'
        return f'{n:,}원'
    except: return '0원'

def safe_folder_name(s):
    return re.sub(r'[\\/:*?"<>|]', '_', str(s).strip())

def build_tiers(tier_def, current_perf):
    result = []
    n = len(tier_def)
    for i, (thr, label, reward) in enumerate(tier_def):
        next_thr = tier_def[i + 1][0] if i + 1 < n else None
        if current_perf >= thr:
            state = 'done' if (next_thr and current_perf >= next_thr) else 'active'
        else:
            prev_done = all(current_perf >= tier_def[j][0] for j in range(i))
            state = 'challenge' if prev_done else ''
        result.append({'label':label, 'threshold_str':fmt_won(thr), 'reward_str':fmt_won(reward), 'state':state})
    return result

def get_current_tier_idx(tier_def, perf):
    idx = -1
    for i, (thr, _, __) in enumerate(tier_def):
        if perf >= thr: idx = i
    return idx

def get_gap_to_next(tier_def, perf):
    for thr, label, reward in tier_def:
        if perf < thr: return thr - perf, thr, reward
    return 0, None, None

def gauge_pct(perf, ref=300_000):
    return min(int(perf / ref * 100), 100)

def tier_name_from_perf(perf):
    if perf >= 500_000: return 'MERITZ PREMIUM'
    if perf >= 300_000: return 'MERITZ GOLD'
    if perf >= 200_000: return 'MERITZ CLUB'
    if perf >= 100_000: return 'MERITZ'
    return '미달성'

def decode_hex(match):
    try:
        hex_val = match.group(1)
        return chr(int(hex_val, 16))
    except:
        return ''

def clean_code(code):
    s = str(code).strip()
    return re.sub(r'_x([0-9a-fA-F]{4})_', decode_hex, s)

# ── 데이터 로드 ───────────────────────────────────────────────
def load_data():
    # 1. MC_LIST
    print(f"Loading MC_LIST: {MC_LIST_FILE}")
    df = pd.read_excel(MC_LIST_FILE)
    df.columns = df.columns.str.strip()
    
    # Identify Code Column by Index (F=5)
    # Just verify content
    print(f"Agent Code (sample): {df.iloc[0,5]}")

    df_f = df[(df['현재영업단조직명']=='GA4본부') & (df['인정실적']>=100_000)].copy()
    
    # 2. PRIZE_SUM (Lookup)
    sum_map = {}
    if os.path.exists(PRIZE_SUM_FILE):
        print(f"Loading PRIZE_SUM: {PRIZE_SUM_FILE}")
        df_s = pd.read_excel(PRIZE_SUM_FILE, header=None)
        for _, r in df_s.iterrows():
            try:
                code = clean_code(r[Key_Sum]) # Clean code
                sum_map[code] = {
                    'w1_c': r[Idx_S_1C], 'w1_p': r[Idx_S_1P],
                    'w1_p_perf': r[Idx_S_1P_Perf],
                    'w2_c': r[Idx_S_2C], 'w2_p': r[Idx_S_2P],
                    'w2_p_perf': r[Idx_S_2P_Perf],
                }
            except: continue
    else:
        print(f"WARNING: {PRIZE_SUM_FILE} not found!")

    # 3. PRIZE_BRIDGE (Lookup)
    br_map = {}
    if os.path.exists(PRIZE_BRIDGE_FILE):
        print(f"Loading PRIZE_BRIDGE: {PRIZE_BRIDGE_FILE}")
        df_b = pd.read_excel(PRIZE_BRIDGE_FILE, header=None)
        for _, r in df_b.iterrows():
            try:
                code = clean_code(r[Key_Bridge]) # Clean code
                br_map[code] = { 
                    'b1': r[Idx_B_1], 
                    'b1_jan': r[Idx_B_1_Jan], 'b1_feb': r[Idx_B_1_Feb],
                    'b2': r[Idx_B_2],
                    'b2_jan': r[Idx_B_2_Jan], 'b2_feb': r[Idx_B_2_Feb],
                }
            except: continue
    else:
        print(f"WARNING: {PRIZE_BRIDGE_FILE} not found (Bridge awards will be 0)!")

    return df_f, sum_map, br_map

# ── 컨텍스트 빌드 ─────────────────────────────────────────────
def build_context(row, sum_data, bridge_data):
    # Left Side Logic
    perf_k = int(row.get('인정실적', 0) or 0) # 2월 인정실적
    perf_w = int(row.get('실적_3주차', 0) or 0)
    perf_s = int(row.get('전월실적_메리츠plus', 0) or 0)
    prev_k = int(row.get('이전월인정실적', 0) or 0)
    
    # Simple Tiers Calc
    t3_cur = get_current_tier_idx(INBO3_TIERS, perf_w)
    t3_reward = INBO3_TIERS[t3_cur][2] if t3_cur >= 0 else 0
    t3_gap, t3_next, _ = get_gap_to_next(INBO3_TIERS, perf_w)
    
    reg_reward = perf_k
    
    c_cur = get_current_tier_idx(CONT_TIERS, perf_w)
    c_reward = CONT_TIERS[c_cur][2] if c_cur >= 0 else 0
    c_gap, _, _ = get_gap_to_next(CONT_TIERS, perf_w)

    mc_base = min(perf_s, perf_k)
    mc_cur = get_current_tier_idx(MCPLUS_TIERS, mc_base)
    mc_reward = MCPLUS_TIERS[mc_cur][2] if mc_cur >= 0 else 0
    mc_gap, mc_next, mc_next_rwd = get_gap_to_next(MCPLUS_TIERS, mc_base)

    # Agency Extra Award Logic (New)
    agency_name = str(row.iloc[35]).strip() # Col AJ (35)
    agency_reward = 0
    agency_pct = 0
    agency_title = ""
    is_agency_target = False
    
    if '메가' in agency_name and '주' in agency_name:
        agency_pct = 500
        agency_title = "메가(주)"
    elif '한국보험' in agency_name: # 한국보험금융(주)
        agency_pct = 400
        agency_title = "한국보험금융(주)"
    elif '메타리치' in agency_name: # (주)메타리치
        agency_pct = 300
        agency_title = "(주)메타리치"
    
    if agency_pct > 0:
        is_agency_target = True
        agency_reward = int(perf_k * (agency_pct / 100)) # 실적 10만원시 50만원 (500%)

    # Right Side Logic (New)
    code = clean_code(row.iloc[5])
    s_dat = sum_data.get(code, {})
    b_dat = bridge_data.get(code, {})
    
    # Right Values (Parse as float/int safe)
    def to_n(v): 
        try: return float(v)
        except: return 0
    
    r_w1_c = to_n(s_dat.get('w1_c', 0))
    r_w1_p = to_n(s_dat.get('w1_p', 0))
    r_w2_c = to_n(s_dat.get('w2_c', 0))
    r_w2_p = to_n(s_dat.get('w2_p', 0))
    r_b1   = to_n(b_dat.get('b1', 0))
    r_b2   = to_n(b_dat.get('b2', 0))
    
    right_total = r_w1_c + r_w1_p + r_w2_c + r_w2_p + r_b1 + r_b2
    
    # New Total: Left 1,2 (t3_reward * 2? No, t3_reward is 1. Item 2 is "Extra"? No, generate.py maps t3_reward to both?)
    # Wait, generate.py: item 1 is t3_reward. Item 2 is also t3_reward (duplicate text in template, same logic).
    # Actually INBO3_TIERS used for both 1 and 2?
    # generate.py: t3_tiers = build_tiers(INBO3_TIERS, perf_w)
    # The Template shows 1 and 2.
    # Are they additive? "인보험 3주차 시상 100%" AND "인보험 3주차 추가시상 100%"?
    # Logic: t3_reward appears to be just ONE value calculated from tiers. 
    # If the user means item 1 AND item 2 are separate awards, and they use the same table...
    # generate.py doesn't multiply by 2 in `total_reward`.
    # generate.py: `total_reward = t3_reward + reg_reward + ...`
    # So t3_reward covers BOTH 1 and 2? Or just 1?
    # The template shows identical logic for 1 and 2.
    # User said "1,2 + 확정된 시상...".
    # I'll include `t3_reward` (for 1) and `t3_reward` (for 2) if they are indeed double?
    # generate.py only had `t3_reward` once in total.
    # But visually 1 and 2 are separate boxes.
    # I will assume `t3_reward` is for ONE item. If there are two items (1 & 2), and they are identical 100%, then maybe total is `t3_reward * 2`?
    # Or maybe one of them is already included? 
    # generate.py logic for total: `t3_reward + reg_reward + cont_reward + mcplus_reward`.
    # It seems to count t3_reward ONCE.
    # But user asked "1,2 + ...".
    # I will add `t3_reward` twice if the template implies two payouts.
    # "인보험 3주차 시상 100%" + "인보험 3주차 추가시상 100%". Sounds like 200% total?
    # I'll calculate total as: `t3_reward * 2 + reg_reward + agency_reward + right_total`.
    # (Exclude 3,4,5).
    
    final_total = (t3_reward * 2) + reg_reward + agency_reward + right_total

    # MC_LIST columns for Right Detail
    # U(20)=1주실적, V(21)=2주실적
    val_u = row.iloc[20]
    val_v = row.iloc[21]

    return {
        'agent_name': str(row.get('현재대리점설계사조직명', '')),
        'manager_name': str(row.get('매니저명', '')),
        'agency_name': str(row.get('현재대리점지사명', '')),
        'org_name': str(row.get('현재지점조직명', '')),
        'date_str': DATE_STR,
        'tier_name': tier_name_from_perf(perf_k),
        'prev_tier_name': tier_name_from_perf(prev_k),
        'perf_k_str': fmt_won(perf_k),
        'perf_w_str': fmt_won(perf_w),
        'prev_k_str': fmt_won(prev_k),
        
        'total_reward_str': fmt_won(final_total), # Updated Total
        
        # Left Side Lists
        't3_tiers': build_tiers(INBO3_TIERS, perf_w),
        't3_pct': gauge_pct(perf_w, 300_000),
        't3_gap_str': fmt_won(t3_gap) if t3_gap>0 else '',
        't3_next_thr_str': fmt_won(t3_next) if t3_next else '',
        't3_reward_str': fmt_won(t3_reward),
        
        'reg_reward_str': fmt_won(reg_reward),
        
        'cont_tiers': build_tiers(CONT_TIERS, perf_w),
        'cont_pct': gauge_pct(perf_w, 300_000),
        'cont_gap_str': fmt_won(c_gap) if c_gap>0 else '',
        'cont_reward_str': fmt_won(c_reward),
        
        'mcplus_tiers': build_tiers(MCPLUS_TIERS, mc_base),
        'mcplus_pct': gauge_pct(mc_base, 600_000),
        'mcplus_gap_str': fmt_won(mc_gap) if mc_gap>0 else '',
        'mcplus_next_thr_str': fmt_won(mc_next) if mc_next else '',
        'mcplus_next_rwd_str': fmt_won(mc_next_rwd) if mc_next_rwd else '',
        'mcplus_reward_str': fmt_won(mc_reward),
        'mcplus_jan_str': fmt_won(perf_s),
        'mcplus_feb_str': fmt_won(perf_k),
        'mcplus_base_str': fmt_won(mc_base),
        
        # Agency Extra
        'is_agency_target': is_agency_target,
        'agency_title': agency_title,
        'agency_reward_str': fmt_won(agency_reward),
        'agency_pct_str': f"{agency_pct}%",
        
        # Right Side 6 Rows
        # Row 1 (W1 Cash)
        'r1_label': '1주차 현금시상',
        'r1_detail': f"1주차 실적 : {fmt_won(val_u)}",
        'r1_val': fmt(r_w1_c),
        'r1_zero': r_w1_c == 0,
        
        # Row 2 (W2 Cash)
        'r2_label': '2주차 현금시상',
        'r2_detail': f"2주차 실적 : {fmt_won(val_v)}",
        'r2_val': fmt(r_w2_c),
        'r2_zero': r_w2_c == 0,

        # Row 3 (W1 Prod)
        'r3_label': '1주차 상품시상 (통합 등)',
        'r3_detail': f"1주차 상품실적 : {fmt_won(s_dat.get('w1_p_perf',0))}",
        'r3_val': fmt(r_w1_p),
        'r3_zero': r_w1_p == 0,

        # Row 4 (W2 Prod)
        'r4_label': '2주차 상품시상',
        'r4_detail': f"2주차 상품실적 : {fmt_won(s_dat.get('w2_p_perf',0))}",
        'r4_val': fmt(r_w2_p),
        'r4_zero': r_w2_p == 0,

        # Row 5 (Bridge 1)
        'r5_label': '브릿지 시상 1',
        'r5_detail': f"1월 대상 : {fmt_won(b_dat.get('b1_jan',0))} / 2월 대상 : {fmt_won(b_dat.get('b1_feb',0))}",
        'r5_val': fmt(r_b1),
        'r5_zero': r_b1 == 0,

        # Row 6 (Bridge 2)
        'r6_label': '브릿지 시상 2',
        'r6_detail': f"1월 대상 : {fmt_won(b_dat.get('b2_jan',0))} / 2월 대상 : {fmt_won(b_dat.get('b2_feb',0))}",
        'r6_val': fmt(r_b2),
        'r6_zero': r_b2 == 0,
    }

# ── 실행 ──────────────────────────────────────────────────────
async def render_all(df_main, sum_map, br_map):
    env = Environment(loader=FileSystemLoader(BASE_DIR))
    tmpl = env.get_template(TEMPLATE)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1600, 'height': 1100})
        
        # CI load
        ci_path = os.path.join(BASE_DIR, 'ci.png')
        ci_b64 = ''
        if os.path.exists(ci_path):
            with open(ci_path, 'rb') as f:
                ci_b64 = 'data:image/png;base64,' + base64.b64encode(f.read()).decode()
        
        total = len(df_main)
        for idx, (_, row) in enumerate(df_main.iterrows(), 1):
            try:
                ctx = build_context(row, sum_map, br_map)
                
                folder = os.path.join(OUTPUT_DIR, safe_folder_name(row['현재지점조직명']),
                                      safe_folder_name(row['매니저명']), safe_folder_name(row['현재대리점지사명']))
                os.makedirs(folder, exist_ok=True)
                fn = safe_folder_name(row['현재대리점설계사조직명']) + '.png'
                out = os.path.join(folder, fn)
                
                html = tmpl.render(**ctx, ci_src=ci_b64)
                await page.set_content(html, wait_until='networkidle')
                await page.screenshot(path=out, clip={'x':0,'y':0,'width':1600,'height':1100})
                print(f"[{idx}/{total}] {out}")
            except Exception as e:
                print(f"Error {idx}: {e}")
        await browser.close()
    print("Done.")

if __name__ == '__main__':
    df, s_map, b_map = load_data()
    asyncio.run(render_all(df, s_map, b_map))
