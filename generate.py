import pandas as pd
import os
import re
import asyncio
import base64
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright

# ── 설정 ──────────────────────────────────────────────────────
BASE_DIR   = r'C:\Users\chlgn\OneDrive\Desktop\general prize'
EXCEL_FILE = os.path.join(BASE_DIR, '0219MC_LIST_OUT_202602 (1).xlsx')
TEMPLATE   = 'template.html'
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
DATE_STR   = '2026년 2월 19일 기준'

# 컬럼 매핑
# W열: 실적_3주차  (인보험 3주차 시상 ①②)
# K열: 인정실적    (2월 인보험 정규시상 ③)
# S열: 전월실적_메리츠plus  (1월 실적 for MC PLUS+ ⑥)

# ── 시상 구간 정의 ────────────────────────────────────────────

# ①② 인보험 3주차 시상 / 추가시상 (W열 기준)
INBO3_TIERS = [
    (100_000, '10만', 100_000),
    (200_000, '20만', 200_000),
    (300_000, '30만', 300_000),
    (500_000, '50만', 500_000),
]

# ③ 2월 인보험 정규시상 (K열 기준) - 정률 100%
#    구간 없음, 실적 그대로 지급

# ④ 2~3월 연속시상 1 (2/19~28 구간 달성 후 3/1~15 10만 달성 시 지급)
# ⑤ 2~3월 연속시상 2 (2/19~28 구간 달성 후 3/1~8  10만 달성 시 지급)
CONT_TIERS = [
    (100_000, '10만', 200_000),
    (200_000, '20만', 600_000),
    (300_000, '30만', 800_000),
    (500_000, '50만', 1_800_000),
]

# ⑥ MC PLUS+ 구간 (1월·2월·3월 연속, 최저구간 기준)
# 달성 조건: 20/40/60/80/100만 → 시상금 60/120/180/240/300만
MCPLUS_TIERS = [
    (200_000,  '20만',  600_000),
    (400_000,  '40만', 1_200_000),
    (600_000,  '60만', 1_800_000),
    (800_000,  '80만', 2_400_000),
    (1_000_000,'100만',3_000_000),
]

# ── 유틸 함수 ─────────────────────────────────────────────────
def fmt(n):
    n = int(n)
    if n == 0:
        return '0원'
    if n % 10_000 == 0:
        return f'{n // 10_000:,}만원'
    return f'{n:,}원'

def safe_folder_name(s):
    return re.sub(r'[\\/:*?"<>|]', '_', str(s).strip())

def build_tiers(tier_def, current_perf):
    """티어 리스트 빌드 (done / active / challenge / '')"""
    result = []
    n = len(tier_def)
    for i, (thr, label, reward) in enumerate(tier_def):
        next_thr = tier_def[i + 1][0] if i + 1 < n else None
        if current_perf >= thr:
            state = 'done' if (next_thr and current_perf >= next_thr) else 'active'
        else:
            prev_done = all(current_perf >= tier_def[j][0] for j in range(i))
            state = 'challenge' if prev_done else ''
        result.append({
            'label':         label,
            'threshold_str': fmt(thr),
            'reward_str':    fmt(reward),
            'state':         state,
        })
    return result

def get_current_tier_idx(tier_def, perf):
    idx = -1
    for i, (thr, _, __) in enumerate(tier_def):
        if perf >= thr:
            idx = i
    return idx

def get_gap_to_next(tier_def, perf):
    """(부족금액, 다음구간threshold, 다음구간reward)"""
    for thr, label, reward in tier_def:
        if perf < thr:
            return thr - perf, thr, reward
    return 0, None, None

def gauge_pct(perf, ref=300_000):
    return min(int(perf / ref * 100), 100)

def tier_name_from_perf(perf):
    if perf >= 500_000:   return 'MERITZ PREMIUM'
    if perf >= 300_000:   return 'MERITZ GOLD'
    if perf >= 200_000:   return 'MERITZ CLUB'
    if perf >= 100_000:   return 'MERITZ'
    return '미달성'

# ── 데이터 로드 ───────────────────────────────────────────────
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()
    df_f = df[
        (df['현재영업단조직명'] == 'GA4본부') &
        (df['인정실적'] >= 100_000)
    ].copy()
    print(f"[INFO] 전체 {len(df)}행 → 필터 후 {len(df_f)}행")
    return df_f

# ── 렌더 컨텍스트 빌드 ───────────────────────────────────────
def build_context(row):
    # 실적 컬럼
    perf_k   = int(row.get('인정실적', 0) or 0)          # K열: 2월 인보험
    perf_w   = int(row.get('실적_3주차', 0) or 0)        # W열: 3주차 실적
    perf_s   = int(row.get('전월실적_메리츠plus', 0) or 0) # S열: 1월 실적(MC+)
    prev_k   = int(row.get('이전월인정실적', 0) or 0)

    # 주차 실적
    w1 = int(row.get('실적_1주차', 0) or 0)
    w2 = int(row.get('실적_2주차', 0) or 0)
    w3 = perf_w
    max_week = max(w1, w2, w3, 1)

    # ①② 인보험 3주차 시상 (W열)
    t3_tiers = build_tiers(INBO3_TIERS, perf_w)
    t3_cur   = get_current_tier_idx(INBO3_TIERS, perf_w)
    t3_gap, t3_next_thr, t3_next_rwd = get_gap_to_next(INBO3_TIERS, perf_w)
    t3_reward = INBO3_TIERS[t3_cur][2] if t3_cur >= 0 else 0
    t3_pct    = gauge_pct(perf_w, 300_000)

    # ③ 2월 인보험 정규시상 (K열 100%)
    reg_reward = perf_k  # 정률 100%
    reg_gap, reg_next_thr, _ = get_gap_to_next(INBO3_TIERS, perf_k)
    reg_pct = gauge_pct(perf_k, 500_000)

    # ④⑤ 연속시상 (W열 기준 2/19~28 구간)
    cont_tiers   = build_tiers(CONT_TIERS, perf_w)
    cont_cur_idx = get_current_tier_idx(CONT_TIERS, perf_w)
    cont_reward  = CONT_TIERS[cont_cur_idx][2] if cont_cur_idx >= 0 else 0
    cont_gap, cont_next_thr, _ = get_gap_to_next(CONT_TIERS, perf_w)
    cont_pct = gauge_pct(perf_w, 300_000)

    # ⑥ MC PLUS+ (S열=1월, K열=2월, 1·2월 중 최소 구간 기준)
    # 연속 달성 조건: 1·2·3월 모두 해당 구간 이상
    # 현재는 1·2월만 알 수 있으므로, min(S, K)로 달성 가능 최대 구간 계산
    mcplus_base       = min(perf_s, perf_k)  # 1·2월 min 실적
    mcplus_tiers      = build_tiers(MCPLUS_TIERS, mcplus_base)
    mcplus_cur_idx    = get_current_tier_idx(MCPLUS_TIERS, mcplus_base)
    mcplus_reward     = MCPLUS_TIERS[mcplus_cur_idx][2] if mcplus_cur_idx >= 0 else 0
    mcplus_gap, mcplus_next_thr, mcplus_next_rwd = get_gap_to_next(MCPLUS_TIERS, mcplus_base)
    mcplus_pct        = gauge_pct(mcplus_base, 600_000)
    mcplus_jan_str    = fmt(perf_s)
    mcplus_feb_str    = fmt(perf_k)
    mcplus_base_str   = fmt(mcplus_base)

    # 총 시상금
    total_reward = t3_reward + reg_reward + cont_reward + mcplus_reward

    return {
        # 인적정보
        'agent_name':      str(row.get('현재대리점설계사조직명', '')),
        'manager_name':    str(row.get('매니저명', '')),
        'agency_name':     str(row.get('현재대리점지사명', '')),
        'org_name':        str(row.get('현재지점조직명', '')),
        'branch_name':     str(row.get('현재영업단조직명', '')),
        'date_str':        DATE_STR,
        'tier_name':       tier_name_from_perf(perf_k),
        'prev_tier_name':  tier_name_from_perf(prev_k),

        # 실적
        'perf_k_str':      fmt(perf_k),
        'perf_w_str':      fmt(perf_w),
        'prev_k_str':      fmt(prev_k),
        'total_reward_str': fmt(total_reward),

        # 주차
        'weeks': [
            {'label': '1주차', 'amount_str': fmt(w1), 'pct': min(int(w1 / max_week * 100), 100)},
            {'label': '2주차', 'amount_str': fmt(w2), 'pct': min(int(w2 / max_week * 100), 100)},
            {'label': '3주차', 'amount_str': fmt(w3), 'pct': min(int(w3 / max_week * 100), 100)},
        ],

        # ①② 3주차 시상
        't3_tiers':        t3_tiers,
        't3_pct':          t3_pct,
        't3_gap_str':      fmt(t3_gap) if t3_gap > 0 else '',
        't3_next_thr_str': fmt(t3_next_thr) if t3_next_thr else '',
        't3_reward_str':   fmt(t3_reward),

        # ③ 정규시상
        'reg_reward_str':  fmt(reg_reward),
        'reg_pct':         reg_pct,
        'reg_gap_str':     fmt(reg_gap) if reg_gap > 0 else '',
        'reg_next_thr_str': fmt(reg_next_thr) if reg_next_thr else '',

        # ④⑤ 연속시상
        'cont_tiers':      cont_tiers,
        'cont_pct':        cont_pct,
        'cont_gap_str':    fmt(cont_gap) if cont_gap > 0 else '',
        'cont_reward_str': fmt(cont_reward),
        'cont_achieved':   cont_reward > 0,

        # ⑥ MC PLUS+
        'mcplus_tiers':          mcplus_tiers,
        'mcplus_pct':            mcplus_pct,
        'mcplus_gap_str':        fmt(mcplus_gap) if mcplus_gap > 0 else '',
        'mcplus_next_thr_str':   fmt(mcplus_next_thr) if mcplus_next_thr else '',
        'mcplus_next_rwd_str':   fmt(mcplus_next_rwd) if mcplus_next_rwd else '',
        'mcplus_reward_str':     fmt(mcplus_reward),
        'mcplus_achieved':       mcplus_reward > 0,
        'mcplus_jan_str':        mcplus_jan_str,
        'mcplus_feb_str':        mcplus_feb_str,
        'mcplus_base_str':       mcplus_base_str,
    }

# ── PNG 생성 ──────────────────────────────────────────────────
async def render_all(df_filtered):
    env  = Environment(loader=FileSystemLoader(BASE_DIR))
    tmpl = env.get_template(TEMPLATE)

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page    = await browser.new_page(viewport={'width': 1600, 'height': 1100})

        total = len(df_filtered)

        # img.jpg + ci.png → base64 데이터 URL (1회 로드)
        img_path = os.path.join(BASE_DIR, 'img.jpg')
        with open(img_path, 'rb') as f:
            img_b64 = 'data:image/jpeg;base64,' + base64.b64encode(f.read()).decode()
        ci_path = os.path.join(BASE_DIR, 'ci.png')
        with open(ci_path, 'rb') as f:
            ci_b64 = 'data:image/png;base64,' + base64.b64encode(f.read()).decode()

        for idx, (_, row) in enumerate(df_filtered.iterrows(), 1):
            try:
                ctx = build_context(row)

                folder = os.path.join(
                    OUTPUT_DIR,
                    safe_folder_name(row['현재지점조직명']),
                    safe_folder_name(row['매니저명']),
                    safe_folder_name(row['현재대리점지사명']),
                )
                os.makedirs(folder, exist_ok=True)

                filename = safe_folder_name(row['현재대리점설계사조직명']) + '.png'
                out_path  = os.path.join(folder, filename)

                html_str = tmpl.render(**ctx, img_src=img_b64, ci_src=ci_b64)
                await page.set_content(html_str, wait_until='networkidle')
                await page.screenshot(path=out_path, clip={'x': 0, 'y': 0, 'width': 1600, 'height': 1100})
                print(f"[{idx}/{total}] {out_path}")

            except Exception as e:
                print(f"[ERROR] {idx}번째 행 오류: {e}")

        await browser.close()
    print(f"\n✅ 완료! 총 {total}장 생성 → {OUTPUT_DIR}")

if __name__ == '__main__':
    df_filtered = load_data()
    asyncio.run(render_all(df_filtered))