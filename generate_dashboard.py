"""
메리츠화재 GA4본부 시상 대시보드 자동 생성기
조건: AF열(현재영업단조직명)='GA4본부', K열(인정실적)>=100,000원
출력: AH열/E열/AL열/G열.png
"""

import os
import re
import time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright

# ─── 경로 설정 ────────────────────────────────────────
BASE_DIR    = Path(r'C:\Users\chlgn\OneDrive\Desktop\general prize')
XLSX_PATH   = BASE_DIR / '0219MC_LIST_OUT_202602 (1).xlsx'
OUTPUT_ROOT = BASE_DIR

# TEST_MODE = True이면 첫 3명만 생성
TEST_MODE = False
TEST_COUNT = 3

# ─── 상수 ────────────────────────────────────────────
WEEK_THRESHOLDS = [100_000, 200_000, 300_000, 500_000]  # 오름차순
WEEK_LABELS_KO  = ['10만↑', '20만↑', '30만↑', '50만↑']
WEEK_REWARDS    = [100_000, 200_000, 300_000, 500_000]
CONN_MAP        = {1: 200_000, 2: 600_000, 3: 800_000, 4: 1_800_000}
MC_TIERS        = [200_000, 400_000, 600_000, 800_000, 1_000_000]
MC_REWARDS      = [600_000, 1_200_000, 1_800_000, 2_400_000, 3_000_000]
MC_LABELS       = ['20만↑', '40만↑', '60만↑', '80만↑', '100만↑']


# ─── 유틸 ────────────────────────────────────────────
def safe_int(v):
    try: return int(v) if pd.notna(v) else 0
    except: return 0

def safe_str(v):
    return str(v).strip() if pd.notna(v) else ''

def sanitize_path(name):
    return re.sub(r'[\\/:*?"<>|]', '_', str(name).strip()) or 'unknown'

def fmt(amount):
    if amount == 0: return '0원'
    if amount % 10_000 == 0: return f'{amount // 10_000}만원'
    return f'{amount:,}원'

def fmt_comma(amount):
    return f'{amount:,}원'

def pct(current, target):
    if target <= 0: return 100
    return min(100, round(current * 100 / target))

def get_week_tier(k):
    """10만/20만/30만/50만 → tier 1~4, 미달 0"""
    for i in range(3, -1, -1):
        if k >= WEEK_THRESHOLDS[i]:
            return i + 1
    return 0

def get_gap_to_next(k, tier):
    """다음 구간까지 남은 금액, 다음 구간 라벨"""
    if tier == 0:
        gap = WEEK_THRESHOLDS[0] - k
        return gap, '10만'
    if tier >= 4:
        return 0, ''
    gap = WEEK_THRESHOLDS[tier] - k
    return max(0, gap), ['20만', '30만', '50만'][tier - 1]

def get_prev_grade(s):
    if s >= 1_000_000: return 'MERITZ VVIP', 'tier-vvip'
    if s >= 500_000:   return 'MERITZ VIP',  'tier-vip'
    if s >= 300_000:   return 'MERITZ GOLD', 'tier-gold'
    if s >= 200_000:   return 'MERITZ CLUB', 'tier-club'
    return 'MERITZ ENTRY', 'tier-entry'

def get_mc_tier_idx(target_amount):
    try: return MC_TIERS.index(int(target_amount))
    except: return -1


# ─── HTML 생성 ────────────────────────────────────────
def generate_html(d):
    perf_k = d['perf_k']
    perf_s = d['perf_s']

    tier = get_week_tier(perf_k)
    tier_gap, next_label = get_gap_to_next(perf_k, tier)

    grade_name, grade_css = get_prev_grade(perf_s)

    weekly_reward = WEEK_REWARDS[tier - 1] if tier > 0 else 0
    mc_rate = d['mc_rate'] or 100
    regular_reward = round(perf_k * mc_rate / 100)
    conn_reward = CONN_MAP.get(tier, 0)

    mc_tier_consec = d['mc_plus_tier_consec']
    mc_target = d['mc_plus_target']
    mc_gap = d['mc_plus_gap']
    mc_tier_idx = get_mc_tier_idx(mc_tier_consec)
    mc_reward = MC_REWARDS[mc_tier_idx] if mc_tier_idx >= 0 else 0
    mc_plus_achieved = (mc_gap == 0 and mc_tier_idx >= 0)
    mc_gauge_pct = pct(perf_k - mc_gap, mc_target) if mc_target > 0 else 0

    total_expected = (weekly_reward * 2) + regular_reward + (conn_reward * 2) + mc_reward

    # ── 구간 셀 4칸 ──
    def tier_cells(cur):
        out = ''
        for i in range(4):
            t = i + 1
            if cur > t:
                cls = 'tc done'; chk = '<div class="tc-check">✓</div>'
            elif cur == t:
                cls = 'tc current'; chk = ''
            else:
                cls = 'tc'; chk = ''
            out += f'''<div class="{cls}">{chk}
              <div class="tc-target">{WEEK_LABELS_KO[i]}</div>
              <div class="tc-reward">{WEEK_REWARDS[i]//10000}만원</div>
            </div>'''
        return out

    # ── 게이지 안내 텍스트 ──
    if tier == 0:
        gauge_w = pct(perf_k, WEEK_THRESHOLDS[0])
        gap_html = f'<span class="g-gap">{fmt(tier_gap)} 더 필요</span>'
    elif tier >= 4:
        gauge_w = 100
        gap_html = '<span class="g-ok">✓ 최고 구간 달성!</span>'
    else:
        gauge_w = pct(perf_k, WEEK_THRESHOLDS[tier])
        gap_html = f'다음 {next_label}까지 <span class="g-gap">{fmt(tier_gap)} 남음</span>'

    # ── 연속가동 셀 ──
    def conn_cells(cur):
        conn_rewards_c = ['20만', '60만', '80만', '180만']
        out = ''
        for i in range(4):
            t = i + 1
            cls = 'cc cc-active' if t == cur else 'cc'
            chk = ' ✓' if cur > t else ''
            out += f'''<div class="{cls}">
              <div class="cc-src">{WEEK_LABELS_KO[i]}{chk}</div>
              <div class="cc-plus">+3월 10만↑</div>
              <div class="cc-reward">{conn_rewards_c[i]}</div>
            </div>'''
        return out

    # ── MC PLUS 티어 그리드 ──
    def mc_grid(mi):
        out = ''
        for i in range(5):
            if mi < 0:
                cls = 'pt'; lbl = ''
            elif i < mi:
                cls = 'pt pt-done'; lbl = ' ✓'
            elif i == mi:
                cls = 'pt pt-current'; lbl = ''
            else:
                cls = 'pt pt-ahead'; lbl = ''
            out += f'''<div class="{cls}">
              <span class="pt-target">{MC_LABELS[i]}{lbl}</span>
              <div class="pt-reward">{MC_REWARDS[i]//10000}만</div>
              <div class="pt-dot"></div>
            </div>'''
        return out

    conn_gap_text = f'{fmt(tier_gap)} 추가 시 {WEEK_LABELS_KO[tier]} 구간!' if tier_gap > 0 else '현재 구간 달성 완료'
    if mc_tier_idx >= 0:
        mc_gl = (f'<span class="g-cur">현재 {fmt_comma(perf_k)}</span>'
                 f'<span class="g-{"ok" if mc_plus_achieved else "gap"}">'
                 f'{"✓ 달성 완료" if mc_plus_achieved else fmt(mc_gap) + " 더 필요"}</span>')
    else:
        mc_gl = '<span class="g-cur">MC PLUS+ 해당 없음</span><span style="color:var(--text-soft);">1·2·3월 연속 달성 조건</span>'

    # 정규시상 다음구간 note
    if tier_gap > 0:
        reg_note = f'다음 {next_label} 달성 시 <span class="note-gap">+{fmt(round(tier_gap * mc_rate / 100))} 추가</span>'
    else:
        reg_note = '<span style="color:var(--red);">최고 구간 달성 완료</span>'

    html = f'''<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<style>
:root{{--red:#EF3B24;--red-dark:#C42E1A;--red-pale:#FFF0EE;
  --gray:#8A8D91;--gray-light:#C8CACD;--black:#1A1A1A;--white:#FFFFFF;
  --bg:#F0EEE9;--card:#FAFAF8;--surface:#F3F1ED;--surface2:#EAE8E3;
  --border:#E0DDD7;--border-mid:#D0CCC5;--text:#1A1A1A;--text-mid:#4A4A4A;--text-soft:#8A8A8A;}}
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:var(--bg);font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;
  width:1600px;height:1100px;overflow:hidden;display:flex;flex-direction:column;color:var(--text);}}
.header{{width:100%;height:60px;background:var(--white);border-bottom:1px solid var(--border);
  display:flex;align-items:center;justify-content:space-between;padding:0 40px;
  flex-shrink:0;position:relative;}}
.header::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;background:var(--red);}}
.ci-box{{width:34px;height:34px;background:var(--red);display:flex;align-items:center;justify-content:center;}}
.ci-main{{font-size:22px;font-weight:900;letter-spacing:3px;color:var(--black);}}
.ci-sub{{font-size:9px;letter-spacing:3px;color:var(--gray);display:block;}}
.ci-wrap{{display:flex;align-items:center;gap:10px;}}
.hdr-center{{position:absolute;left:50%;transform:translateX(-50%);text-align:center;}}
.hdr-title{{font-size:22px;font-weight:900;letter-spacing:6px;color:var(--black);}}
.hdr-sub{{font-size:9px;letter-spacing:2.5px;color:var(--gray);margin-top:1px;}}
.hdr-date{{font-size:10px;color:var(--gray);text-align:right;letter-spacing:1.5px;}}
.hdr-date strong{{display:block;font-size:13px;font-weight:700;color:var(--text-mid);}}
.body{{flex:1;display:grid;grid-template-columns:800px 800px;overflow:hidden;}}
.left{{background:var(--card);border-right:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden;}}
.agent-block{{padding:22px 36px 18px;border-bottom:1px solid var(--border);display:flex;
  align-items:flex-start;justify-content:space-between;flex-shrink:0;background:var(--white);}}
.tier-tag{{display:inline-block;font-size:10px;font-weight:700;letter-spacing:2px;padding:4px 12px;margin-bottom:9px;}}
.tier-vvip{{background:#C9A84C;color:#fff;}}
.tier-vip{{background:#9E9E9E;color:#fff;}}
.tier-gold{{background:#fff;color:#A08840;border:1.5px solid #C9A84C;}}
.tier-club{{background:var(--red);color:white;}}
.tier-entry{{background:var(--surface);color:var(--gray);border:1px solid var(--border);}}
.agent-name{{font-size:34px;font-weight:900;color:var(--black);letter-spacing:-1.5px;line-height:1;}}
.agent-name small{{font-size:14px;font-weight:300;color:var(--gray);margin-left:4px;}}
.agent-unit{{margin-top:8px;font-size:11px;color:var(--gray);line-height:1.9;}}
.rank-box{{text-align:center;flex-shrink:0;padding-top:2px;}}
.rank-label{{font-size:9px;letter-spacing:2.5px;color:var(--gray-light);text-align:center;margin-top:4px;}}
.perf-row{{display:grid;grid-template-columns:1fr 1fr 1fr;border-bottom:1px solid var(--border);flex-shrink:0;background:var(--white);}}
.perf-cell{{padding:13px 18px;border-right:1px solid var(--border);}}
.perf-cell:last-child{{border-right:none;}}
.perf-label{{font-size:8.5px;letter-spacing:2px;color:var(--gray);margin-bottom:5px;display:block;}}
.perf-val{{font-size:18px;font-weight:900;color:var(--black);letter-spacing:-0.5px;}}
.perf-val.red{{color:var(--red);}}
.perf-grade{{font-size:11px;color:var(--red);font-weight:700;margin-top:4px;}}
.sec-label{{padding:9px 36px 8px;font-size:8.5px;letter-spacing:3px;color:var(--red);font-weight:700;
  border-bottom:1px solid var(--border);flex-shrink:0;background:var(--surface);}}
.tier-table-wrap{{padding:11px 36px;border-bottom:1px solid var(--border);flex-shrink:0;background:var(--white);}}
.tier-cols{{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:8px;}}
.tc{{border:1.5px solid var(--border);padding:7px 5px;text-align:center;background:var(--surface);position:relative;}}
.tc.done{{background:#FFF5F4;border-color:rgba(239,59,36,0.4);}}
.tc.current{{background:var(--red-pale);border:2px solid var(--red);}}
.tc-target{{font-size:9px;color:var(--gray);margin-bottom:4px;}}
.tc.done .tc-target,.tc.current .tc-target{{color:var(--red);font-weight:700;}}
.tc-reward{{font-size:15px;font-weight:900;color:var(--gray-light);}}
.tc.done .tc-reward,.tc.current .tc-reward{{color:var(--red);}}
.tc-check{{position:absolute;top:3px;right:5px;font-size:10px;color:var(--red);}}
.gauge-info{{display:flex;justify-content:space-between;font-size:10px;color:var(--gray);margin-bottom:5px;}}
.g-cur{{color:var(--text-mid);font-weight:500;}}
.g-gap{{color:var(--red);font-weight:700;}}
.g-ok{{color:var(--red);font-weight:700;}}
.gauge-track{{height:7px;background:var(--surface2);overflow:hidden;}}
.gauge-fill{{height:100%;background:linear-gradient(90deg,var(--red-dark),var(--red));}}
.regular-wrap{{padding:13px 36px;flex-shrink:0;border-bottom:1px solid var(--border);background:var(--white);}}
.regular-amount{{font-size:26px;font-weight:900;color:var(--black);letter-spacing:-1px;margin:5px 0 4px;}}
.regular-note{{font-size:9.5px;color:var(--gray);}}
.note-gap{{color:var(--red);font-weight:700;}}
.total-box{{margin:auto 36px 18px;background:var(--red);padding:13px 22px;
  display:flex;align-items:center;justify-content:space-between;flex-shrink:0;}}
.total-label{{font-size:10px;letter-spacing:2px;color:rgba(255,255,255,0.7);}}
.total-sub{{font-size:9px;color:rgba(255,255,255,0.4);margin-top:2px;}}
.total-val{{font-size:29px;font-weight:900;color:#fff;letter-spacing:-1px;}}
.right{{background:var(--bg);display:flex;flex-direction:column;overflow:hidden;}}
.right-grid{{display:grid;grid-template-rows:1fr 1fr 1fr;flex:1;overflow:hidden;}}
.r-sec{{border-bottom:1px solid var(--border);display:flex;flex-direction:column;
  padding:14px 34px;overflow:hidden;background:var(--card);}}
.r-sec:nth-child(even){{background:var(--white);}}
.r-sec:last-child{{border-bottom:none;}}
.r-hdr{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;flex-shrink:0;}}
.r-title{{font-size:13px;font-weight:700;color:var(--black);}}
.r-period{{font-size:9.5px;color:var(--gray);margin-top:2px;}}
.r-badge{{font-size:9px;font-weight:700;padding:3px 10px;flex-shrink:0;border:1px solid transparent;}}
.r-amount{{font-size:18px;font-weight:900;color:var(--red);margin-top:3px;text-align:right;}}
.r-amount.locked{{color:var(--gray-light);}}
.bd-done{{background:#FFF0EE;color:var(--red);border-color:rgba(239,59,36,0.3);}}
.bd-progress{{background:#FFF0EE;color:var(--red);border-color:rgba(239,59,36,0.3);}}
.bd-pending{{background:var(--surface);color:var(--gray);border-color:var(--border);}}
.conn-cols{{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:8px;flex-shrink:0;}}
.cc{{background:var(--surface);border:1.5px solid var(--border);padding:6px 5px;text-align:center;}}
.cc.cc-active{{background:var(--red-pale);border-color:var(--red);}}
.cc-src{{font-size:9px;color:var(--gray);margin-bottom:2px;}}
.cc.cc-active .cc-src{{color:var(--red);font-weight:700;}}
.cc-plus{{font-size:8px;color:var(--gray-light);margin-bottom:2px;}}
.cc-reward{{font-size:14px;font-weight:900;color:var(--gray-light);}}
.cc.cc-active .cc-reward{{color:var(--red);}}
.flow-box{{background:var(--surface);border:1px solid var(--border);padding:8px 12px;
  display:flex;align-items:center;margin-bottom:8px;flex-shrink:0;}}
.flow-step{{flex:1;text-align:center;}}
.flow-lbl{{font-size:8.5px;letter-spacing:0.5px;color:var(--gray);margin-bottom:3px;}}
.flow-val{{font-size:13px;font-weight:900;color:var(--black);}}
.flow-val.ok,.flow-val.need{{color:var(--red);}}
.flow-arrow{{width:28px;text-align:center;color:var(--gray-light);font-size:16px;flex-shrink:0;}}
.flow-reward{{flex:1.1;text-align:center;border-left:1.5px solid var(--border);padding-left:12px;}}
.flow-rlbl{{font-size:8.5px;letter-spacing:0.5px;color:var(--red);margin-bottom:3px;}}
.flow-rval{{font-size:17px;font-weight:900;color:var(--red);}}
.gap-note{{display:flex;align-items:center;gap:8px;margin-bottom:6px;flex-shrink:0;}}
.gap-chip{{background:var(--red-pale);border:1px solid rgba(239,59,36,0.3);padding:3px 10px;
  font-size:10px;font-weight:700;color:var(--red);white-space:nowrap;}}
.gap-desc{{font-size:9.5px;color:var(--gray);}}
.plus-tiers{{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin-bottom:8px;flex-shrink:0;}}
.pt{{background:var(--surface);border:1.5px solid var(--border);padding:7px 4px;text-align:center;}}
.pt.pt-done{{background:#FFF5F4;border-color:rgba(239,59,36,0.3);}}
.pt.pt-current{{background:var(--red-pale);border-color:var(--red);}}
.pt-target{{font-size:9px;color:var(--gray-light);margin-bottom:4px;display:block;}}
.pt.pt-done .pt-target,.pt.pt-current .pt-target{{color:var(--red);font-weight:700;}}
.pt-reward{{font-size:14px;font-weight:900;color:var(--gray-light);}}
.pt.pt-done .pt-reward,.pt.pt-current .pt-reward{{color:var(--red);}}
.pt-dot{{width:5px;height:5px;border-radius:50%;background:var(--border);margin:5px auto 0;}}
.pt.pt-done .pt-dot{{background:rgba(239,59,36,0.3);}}
.pt.pt-current .pt-dot{{background:var(--red);}}
.r-foot{{margin-top:auto;padding:8px 12px;background:var(--surface);border:1px solid var(--border);
  display:flex;justify-content:space-between;align-items:center;flex-shrink:0;}}
.r-foot-note{{font-size:9px;color:var(--gray);line-height:1.7;}}
.r-foot-total{{text-align:right;}}
.r-foot-total-lbl{{font-size:8.5px;letter-spacing:1.5px;color:var(--gray);display:block;margin-bottom:2px;}}
.r-foot-total-val{{font-size:19px;font-weight:900;color:var(--red);}}
.new-tag{{background:var(--red);color:white;font-size:8px;font-weight:900;letter-spacing:0.5px;
  padding:1px 5px;margin-right:5px;vertical-align:middle;position:relative;top:-1px;}}
</style></head><body>

<div class="header">
  <div class="ci-wrap">
    <div class="ci-box">
      <svg width="20" height="20" viewBox="0 0 24 24" fill="white">
        <path d="M2 20V4l10 9 10-9v16h-3V10l-7 6.5L5 10v10z"/>
      </svg>
    </div>
    <div>
      <span class="ci-main">meritz</span>
      <span class="ci-sub">메리츠화재</span>
    </div>
  </div>
  <div class="hdr-center">
    <div class="hdr-title">2월 3주차 파트너 시상 현황</div>
    <div class="hdr-sub">FEBRUARY · WEEK 3 · PARTNER INCENTIVE DASHBOARD</div>
  </div>
  <div class="hdr-date">기준일 <strong>2026.02.19</strong></div>
</div>

<div class="body">
  <div class="left">
    <div class="agent-block">
      <div>
        <div class="tier-tag {grade_css}">{grade_name}</div>
        <div class="agent-name">{d['agent_name']} <small>님</small></div>
        <div class="agent-unit">
          {d['agency_name']}<br>
          담당 매니저 · {d['manager_name']}
        </div>
      </div>
      <div class="rank-box">
        <div style="font-size:11px;font-weight:700;color:var(--red);letter-spacing:2px;">GA4본부</div>
        <div class="rank-label">{d['branch']}</div>
      </div>
    </div>

    <div class="perf-row">
      <div class="perf-cell">
        <span class="perf-label">현재 인정실적</span>
        <div class="perf-val">{fmt_comma(perf_k)}</div>
      </div>
      <div class="perf-cell">
        <span class="perf-label">예상 총 시상금</span>
        <div class="perf-val red">{fmt_comma(total_expected)}</div>
      </div>
      <div class="perf-cell">
        <span class="perf-label">전월 실적 / 등급</span>
        <div class="perf-val" style="font-size:14px;">{fmt_comma(perf_s)}</div>
        <div class="perf-grade">{grade_name}</div>
      </div>
    </div>

    <div class="sec-label">① 인보험 시상 &nbsp;·&nbsp; 2/19–2/28 &nbsp;·&nbsp; 13회차 익월지급</div>
    <div class="tier-table-wrap">
      <div class="tier-cols">{tier_cells(tier)}</div>
      <div class="gauge-info">
        <span class="g-cur">현재 {fmt_comma(perf_k)}</span>
        {gap_html}
      </div>
      <div class="gauge-track"><div class="gauge-fill" style="width:{gauge_w}%"></div></div>
    </div>

    <div class="sec-label" style="border-top:1px solid var(--border);">② 인보험 추가 시상 &nbsp;·&nbsp; 2/19–2/22 &nbsp;·&nbsp; 13회차 익월지급</div>
    <div class="tier-table-wrap">
      <div class="tier-cols">{tier_cells(tier)}</div>
      <div class="gauge-info">
        <span class="g-cur">기간 내 실적 기준</span>
        {gap_html}
      </div>
      <div class="gauge-track"><div class="gauge-fill" style="width:{gauge_w}%"></div></div>
    </div>

    <div class="sec-label" style="border-top:1px solid var(--border);">③ 2월 정규 시상 · 인보험 {mc_rate}% 지급 &nbsp;·&nbsp; 익월지급</div>
    <div class="regular-wrap">
      <div style="display:flex;justify-content:space-between;align-items:baseline;">
        <span style="font-size:11.5px;font-weight:700;">인보험 실적 기반 정규 시상</span>
        <span style="font-size:9px;color:var(--gray);">달성 실적 × {mc_rate}%</span>
      </div>
      <div class="regular-amount">{fmt_comma(regular_reward)}</div>
      <div class="regular-note">{reg_note}</div>
      <div style="margin-top:8px;">
        <div class="gauge-info">
          <span class="g-cur">현재 {fmt_comma(regular_reward)}</span>
          {'<span class="g-gap">다음 구간 달성 시 추가 지급</span>' if tier_gap > 0 else '<span class="g-ok">✓ 최고 구간 확보</span>'}
        </div>
        <div class="gauge-track"><div class="gauge-fill" style="width:{gauge_w}%"></div></div>
      </div>
    </div>

    <div class="total-box">
      <div>
        <div class="total-label">예상 총 시상금 &nbsp;/&nbsp; TOTAL EXPECTED</div>
        <div class="total-sub">인보험×2 + 연속가동×2 + 정규 + MC PLUS+</div>
      </div>
      <div class="total-val">{fmt_comma(total_expected)}</div>
    </div>
  </div>

  <div class="right">
    <div class="right-grid">

      <div class="r-sec">
        <div class="r-hdr">
          <div>
            <div class="r-title">④ 2월→3월 연속가동 시상</div>
            <div class="r-period">2/19–28 구간 달성 → 3/1–15 추가 10만 달성 시 지급 · 13회차</div>
          </div>
          <div style="text-align:right;">
            <div class="r-badge bd-progress">진행중</div>
            <div class="r-amount">{fmt(conn_reward)} 예상</div>
          </div>
        </div>
        <div class="conn-cols">{conn_cells(tier)}</div>
        <div class="flow-box">
          <div class="flow-step">
            <div class="flow-lbl">2월 달성 구간</div>
            <div class="flow-val ok">{WEEK_LABELS_KO[tier-1] if tier > 0 else '미달성'}</div>
          </div>
          <div class="flow-arrow">→</div>
          <div class="flow-step">
            <div class="flow-lbl">3월 추가 (마감 3/15)</div>
            <div class="flow-val need">+10만원 필요</div>
          </div>
          <div class="flow-arrow">→</div>
          <div class="flow-reward">
            <div class="flow-rlbl">달성 시 시상</div>
            <div class="flow-rval">{fmt(conn_reward)}</div>
          </div>
        </div>
        <div class="gap-note">
          <div class="gap-chip">{'구간 달성완료' if tier_gap == 0 else fmt(tier_gap)}</div>
          <div class="gap-desc">{conn_gap_text}</div>
        </div>
        <div class="gauge-track"><div class="gauge-fill" style="width:{gauge_w}%"></div></div>
      </div>

      <div class="r-sec">
        <div class="r-hdr">
          <div>
            <div class="r-title"><span class="new-tag">NEW</span>⑤ 2월→3월 연속가동 추가 시상</div>
            <div class="r-period">2/19–28 구간 달성 → 3/1–<strong style="color:var(--red);">8</strong> 추가 10만 달성 시 지급</div>
          </div>
          <div style="text-align:right;">
            <div class="r-badge bd-progress">진행중</div>
            <div class="r-amount">{fmt(conn_reward)} 예상</div>
          </div>
        </div>
        <div class="conn-cols">{conn_cells(tier)}</div>
        <div class="flow-box">
          <div class="flow-step">
            <div class="flow-lbl">2월 달성 구간</div>
            <div class="flow-val ok">{WEEK_LABELS_KO[tier-1] if tier > 0 else '미달성'}</div>
          </div>
          <div class="flow-arrow">→</div>
          <div class="flow-step">
            <div class="flow-lbl">3월 추가 <strong style="color:var(--red);">마감 3/8!</strong></div>
            <div class="flow-val need">+10만원 필요</div>
          </div>
          <div class="flow-arrow">→</div>
          <div class="flow-reward">
            <div class="flow-rlbl">달성 시 시상</div>
            <div class="flow-rval">{fmt(conn_reward)}</div>
          </div>
        </div>
        <div class="gap-note">
          <div class="gap-chip">{'구간 달성완료' if tier_gap == 0 else fmt(tier_gap)}</div>
          <div class="gap-desc">{conn_gap_text} · 3월 8일 마감!</div>
        </div>
        <div class="gauge-track"><div class="gauge-fill" style="width:{gauge_w}%"></div></div>
      </div>

      <div class="r-sec">
        <div class="r-hdr">
          <div>
            <div class="r-title">⑥ 메리츠클럽 PLUS+ (MC PLUS+)</div>
            <div class="r-period">1·2·3월 연속 구간 달성 조건 · 13회차 지급</div>
          </div>
          <div style="text-align:right;">
            <div class="r-badge {'bd-done' if mc_plus_achieved else ('bd-progress' if mc_tier_idx >= 0 else 'bd-pending')}">
              {'달성완료' if mc_plus_achieved else (MC_LABELS[mc_tier_idx] + ' 도전중' if mc_tier_idx >= 0 else '해당없음')}
            </div>
            <div class="r-amount {'locked' if mc_tier_idx < 0 else ''}">{fmt(mc_reward) if mc_tier_idx >= 0 else '—'}</div>
          </div>
        </div>
        <div class="plus-tiers">{mc_grid(mc_tier_idx)}</div>
        <div class="gauge-info">{mc_gl}</div>
        <div class="gauge-track" style="margin-bottom:4px;">
          <div class="gauge-fill" style="width:{mc_gauge_pct if mc_tier_idx >= 0 else 0}%"></div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(5,1fr);text-align:center;font-size:9px;margin-bottom:6px;">
          {''.join(['<div style="color:var(--red);font-weight:700;">' + MC_LABELS[i] + '</div>' if i <= mc_tier_idx else '<div style="color:var(--gray-light);">' + MC_LABELS[i] + '</div>' for i in range(5)])}
        </div>
        <div class="r-foot">
          <div class="r-foot-note">
            ※ 연속가동 시상금은 3월 실적 확정 후 지급<br>
            ※ MC PLUS+는 1·2·3월 월별 구간 연속 달성 조건
          </div>
          <div class="r-foot-total">
            <span class="r-foot-total-lbl">예상 총 시상금</span>
            <div class="r-foot-total-val">{fmt_comma(total_expected)}</div>
          </div>
        </div>
      </div>

    </div>
  </div>
</div>
</body></html>'''
    return html


# ─── 메인 ────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  메리츠화재 GA4본부 시상 대시보드 자동 생성기")
    print("=" * 55)

    print("\n[1/3] 엑셀 로딩 중...")
    df = pd.read_excel(XLSX_PATH, engine='openpyxl')
    print(f"      전체 행: {len(df):,}명")

    filtered = df[
        (df.iloc[:, 31] == 'GA4본부') &
        (df.iloc[:, 10] >= 100_000)
    ].copy().reset_index(drop=True)

    if TEST_MODE:
        filtered = filtered.head(TEST_COUNT)
        print(f"      ★ TEST MODE: {TEST_COUNT}명만 생성")

    total = len(filtered)
    print(f"      대상 인원: {total:,}명 (GA4본부 + 10만원 이상)")

    errors = []
    start_time = time.time()

    print(f"\n[2/3] 이미지 생성 시작...")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1600, 'height': 1100})

        for idx, row in filtered.iterrows():
            try:
                d = {
                    'agent_name':          safe_str(row.iloc[6]),
                    'manager_name':        safe_str(row.iloc[4]),
                    'agent_code':          safe_str(row.iloc[5]),
                    'perf_k':              safe_int(row.iloc[10]),
                    'perf_s':              safe_int(row.iloc[18]),
                    'branch':              safe_str(row.iloc[33]),  # AH
                    'agency_name':         safe_str(row.iloc[37]),  # AL
                    'mc_rate':             safe_int(row.iloc[46]) or 100,
                    'mc_plus_tier_consec': safe_int(row.iloc[51]),
                    'mc_plus_target':      safe_int(row.iloc[52]),
                    'mc_plus_gap':         safe_int(row.iloc[53]),
                }

                # 폴더 트리: AH / E / AL
                folder = (
                    OUTPUT_ROOT
                    / sanitize_path(d['branch'])        # AH
                    / sanitize_path(d['manager_name'])  # E
                    / sanitize_path(d['agency_name'])   # AL
                )
                folder.mkdir(parents=True, exist_ok=True)

                # 파일명: G열 이름 (중복 시 코드 추가)
                fname = sanitize_path(d['agent_name']) + '.png'
                out_path = folder / fname
                if out_path.exists():
                    fname = sanitize_path(d['agent_name']) + '_' + sanitize_path(d['agent_code']) + '.png'
                    out_path = folder / fname

                html = generate_html(d)
                page.set_content(html, wait_until='domcontentloaded')
                page.screenshot(path=str(out_path),
                                clip={'x': 0, 'y': 0, 'width': 1600, 'height': 1100})

                # 진행 표시
                done = idx + 1
                if done % 100 == 0 or done == total or done <= 5:
                    elapsed = time.time() - start_time
                    per_img = elapsed / done
                    remain = per_img * (total - done)
                    print(f"      [{done:4d}/{total}] {d['agent_name']:10s} | "
                          f"경과 {elapsed:.0f}s / 잔여 {remain:.0f}s")

            except Exception as e:
                errors.append((idx, safe_str(row.iloc[6]), str(e)))
                print(f"      ⚠ [{idx+1}] {safe_str(row.iloc[6])} 오류: {e}")

        browser.close()

    elapsed_total = time.time() - start_time
    ok = total - len(errors)

    print(f"\n[3/3] 완료!")
    print(f"      생성 성공: {ok:,}장")
    print(f"      오류:      {len(errors)}건")
    print(f"      총 소요:   {elapsed_total:.0f}초 ({elapsed_total/60:.1f}분)")
    print(f"      저장 위치: {OUTPUT_ROOT}")

    if errors:
        print(f"\n  오류 목록:")
        for i, name, err in errors:
            print(f"    행{i} {name}: {err}")

    print("\n  완료!")


if __name__ == '__main__':
    main()
