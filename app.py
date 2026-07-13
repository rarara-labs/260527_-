import streamlit as st
import os
import json
import pandas as pd
import re
import io
import requests
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go

# ── GitHub 데이터 파일 URL ──
DATA_URL = "https://raw.githubusercontent.com/ons-jinju/260527_-/main/26%EB%85%84_%EC%A7%84%EC%A3%BC%ED%92%88%EC%A7%88%EA%B0%9C%EC%84%A0%ED%8C%80_%EA%B3%A0%EC%9E%A5_RAW_DATA.xlsx"

st.set_page_config(
    page_title="진주품질개선팀 고장현황",
    page_icon="🚨",
    layout="wide",
    initial_sidebar_state="collapsed"
)


st.markdown("""
<style>
html,body,[class*="css"]{font-family:'Noto Sans KR',sans-serif;}

/* ── 탭 ── */
.stTabs [data-baseweb="tab-list"]{gap:1px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{padding:5px 8px;font-size:12px;white-space:nowrap;}

/* ── 알람 카드 ── */
.ac{
  border-left:3px solid rgba(128,128,128,.4);
  border-radius:6px;
  padding:8px 11px;
  margin-bottom:5px;
  background:rgba(128,128,128,.06);
  line-height:1.4;
}
.ac.new  {border-left-color:#f03e3e;}
.ac.bp   {border-left-color:#1971c2;}
.ac.done {border-left-color:#2f9e44;}
.ac.rep  {border-left-color:#f59f00;}
.ac.other{border-left-color:#7048e8;}

/* ── Port 태그 ── */
.ptag{
  font-family:monospace;
  font-size:11px;
  background:rgba(128,128,128,.18);
  padding:1px 5px;
  border-radius:3px;
  word-break:break-all;
}

/* ── 뱃지 ── */
.b{display:inline-block;padding:1px 7px;border-radius:10px;font-size:11px;font-weight:600;margin-right:2px;}
.br {background:rgba(240,62,62,.15); color:#f03e3e;}
.bg {background:rgba(47,158,68,.15);  color:#2f9e44;}
.bb {background:rgba(25,113,194,.15); color:#1971c2;}
.bo {background:rgba(230,119,0,.15);  color:#e67700;}
.by {background:rgba(245,159,0,.18);  color:#e67700;}
.bgr{background:rgba(128,128,128,.15);color:#868e96;}
.bpu{background:rgba(112,72,232,.15); color:#7048e8;}

/* ── 섹션 제목 ── */
.st{font-size:.88rem;font-weight:700;margin:12px 0 6px;padding-bottom:3px;border-bottom:2px solid rgba(128,128,128,.2);}

/* ── 다운로드 버튼 ── */
.stDownloadButton>button{
  background-color:#1971c2!important;
  color:#ffffff!important;
  border:none!important;
  border-radius:8px!important;
  font-weight:600!important;
  width:100%!important;
}

/* ── st.metric 커스텀 ── */
div[data-testid="metric-container"]{
  background:rgba(128,128,128,.06);
  border:1px solid rgba(128,128,128,.14);
  border-radius:8px;
  padding:10px 8px;
  text-align:center;
}
[data-testid="stMetricValue"]{font-size:1.55rem!important;line-height:1.15!important;}
[data-testid="stMetricLabel"]{font-size:.72rem!important;opacity:.75!important;}
[data-testid="stMetricDelta"]{display:none!important;}

/* ── Plotly 차트 ── */
.js-plotly-plot{border-radius:6px;}

/* ── 모바일 ── */
.block-container{padding-top:2.5rem!important;}
@media(max-width:768px){
  .block-container{padding:.4rem .5rem!important;padding-top:3rem!important;}
  [data-testid="stMetricValue"]{font-size:1.25rem!important;}
  div[data-testid="metric-container"]{padding:8px 5px;}
  .ac{padding:7px 9px;}
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════
#  데이터 로딩
# ════════════════════════════════════════════
@st.cache_data(ttl=60)
def load_data(url):
    response = requests.get(url)
    response.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(response.content))
    data = {}

    df_oos = pd.read_excel(xl, sheet_name='5G_LTE OOS_진주')
    df_oos.columns = df_oos.columns.str.strip()
    df_oos['본부'] = pd.to_datetime(df_oos['본부'], errors='coerce')
    cutoff = pd.Timestamp(date.today()) + pd.DateOffset(days=30)
    df_oos = df_oos[df_oos['본부'].isna() | (df_oos['본부'] <= cutoff)]
    data['oos'] = df_oos

    data['alarm_13'] = pd.read_excel(xl, sheet_name='13시 알람 공유')
    data['relay']    = pd.read_excel(xl, sheet_name='중계기 및 MIBOS 알람')
    data['grems']    = pd.read_excel(xl, sheet_name='gREMS')

    for key, sheet in [('rms_a','RMS_A망 미복구'),('rms_dacs','RMS_DACS 미복구'),('rms_rcu','RMS_통합RCU미복구')]:
        try:    data[key] = pd.read_excel(xl, sheet_name=sheet)
        except: data[key] = pd.DataFrame()

    data['5g_raw']  = pd.read_excel(xl, sheet_name='5G Raw')
    data['lte_raw'] = pd.read_excel(xl, sheet_name='LTE Raw')
    try:    data['3g_raw'] = pd.read_excel(xl, sheet_name='3G MOD ')
    except: data['3g_raw'] = pd.DataFrame()

    data['aau_port'] = pd.read_excel(xl, sheet_name='AAU PORT_0518')
    data['rru_port'] = pd.read_excel(xl, sheet_name='RRU PORT_0518')
    try:    data['3g_mod_port'] = pd.read_excel(xl, sheet_name='3G MOD_1027')
    except: data['3g_mod_port'] = pd.DataFrame()

    try:    data['lte_optical'] = pd.read_excel(xl, sheet_name='LTE광레벨불량')
    except: data['lte_optical'] = pd.DataFrame()
    try:    data['5g_optical']  = pd.read_excel(xl, sheet_name='5G광레벨불량')
    except: data['5g_optical']  = pd.DataFrame()

    return data


# ── 알람 히스토리 저장/로드 (날짜별 Port→알람 영구 보존) ──
ALARM_HIST_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'alarm_history.json')

def load_alarm_history():
    """날짜별 알람 히스토리 로드: {날짜str: {port: alarm}}"""
    if os.path.exists(ALARM_HIST_PATH):
        try:
            with open(ALARM_HIST_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_alarm_history(hist):
    """알람 히스토리 저장"""
    try:
        with open(ALARM_HIST_PATH, 'w', encoding='utf-8') as f:
            json.dump(hist, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

@st.cache_data(ttl=300)
def build_alarm_map(_data):
    EXCLUDE = {'Service Unavailable', 'nan', '', 'None'}
    PRIORITY = ['Input Power Failure', 'Link Failure', 'No Connection',
                'SW Error', 'FRU General Problem']

    def parse_fru_local(obj_ref):
        if not obj_ref or pd.isna(obj_ref): return ''
        m = re.search(r'FieldReplaceableUnit=([^,\s]+)', str(obj_ref))
        return m.group(1).strip() if m else ''

    def best_alarm(alarm_set):
        for p in PRIORITY:
            if p in alarm_set: return p
        return ', '.join(sorted(alarm_set)) if alarm_set else None

    rf_alarm = {}
    du_alarm  = {}

    for raw_key in ('5g_raw', 'lte_raw', '3g_raw'):
        df = _data.get(raw_key, pd.DataFrame())
        if df.empty: continue
        for _, r in df.iterrows():
            sp = str(r.get('specificProblem', '')).strip()
            if sp in EXCLUDE: continue
            du = str(r.get('DU', '')).strip()
            if not du or du == 'nan': continue
            du_alarm.setdefault(du, set()).add(sp)
            rf = parse_fru_local(r.get('objectOfReference', ''))
            if rf:
                rf_alarm.setdefault((du, rf), set()).add(sp)

    cell_port_map = {}

    aau = _data.get('aau_port', pd.DataFrame())
    for _, r in aau.iterrows():
        du      = str(r.iloc[3]).strip() if pd.notna(r.iloc[3]) else ''
        rf      = str(r.iloc[5]).strip() if pd.notna(r.iloc[5]) else ''
        cell_id = str(r.iloc[4]).strip() if pd.notna(r.iloc[4]) else ''
        if du and rf and cell_id and cell_id != 'nan':
            cell_port_map[cell_id] = (du, rf)

    rru = _data.get('rru_port', pd.DataFrame())
    for _, r in rru.iterrows():
        du      = str(r.iloc[3]).strip() if pd.notna(r.iloc[3]) else ''
        rf      = str(r.iloc[5]).strip() if pd.notna(r.iloc[5]) else ''
        cell_id = str(r.iloc[4]).strip() if pd.notna(r.iloc[4]) else ''
        if du and rf and cell_id and cell_id != 'nan':
            cell_port_map[cell_id] = (du, rf)

    return rf_alarm, du_alarm, cell_port_map


def get_real_alarm(port, alarm_map):
    rf_alarm, du_alarm, cell_port_map = alarm_map
    port = str(port).strip()
    PRIORITY = ['Input Power Failure', 'Link Failure', 'No Connection',
                'SW Error', 'FRU General Problem']

    def best(s):
        if not s: return None
        for p in PRIORITY:
            if p in s: return p
        return ', '.join(sorted(s))

    if port in cell_port_map:
        du, rf = cell_port_map[port]
        alarms = rf_alarm.get((du, rf), set())
        if alarms: return best(alarms)

    for cell_id, (du, rf) in cell_port_map.items():
        if cell_id in port or port in cell_id:
            alarms = rf_alarm.get((du, rf), set())
            if alarms: return best(alarms)

    parts = port.rsplit('_', 1)
    du = parts[0] if (len(parts) == 2 and parts[1].isdigit()) else port
    alarms = du_alarm.get(du, set())
    return best(alarms) if alarms else None


@st.cache_data(ttl=300)
def build_vswr(_data):
    def parse_fru(obj_ref):
        if not obj_ref or pd.isna(obj_ref): return None
        m = re.search(r'FieldReplaceableUnit=([^,]+)', str(obj_ref))
        return m.group(1).strip() if m else None

    results = []

    df5   = _data['5g_raw'].copy()
    if 'specificProblem' not in df5.columns:
        df5_v = pd.DataFrame()
    else:
        df5_v = df5[df5['specificProblem'] == 'VSWR Over Threshold'].copy()
    if not df5_v.empty and 'objectOfReference' in df5_v.columns:
        df5_v['FRU'] = df5_v['objectOfReference'].apply(parse_fru)
    else:
        df5_v = pd.DataFrame(columns=['DU','FRU','VSWR','기지국명','eventTime'])

    aau     = _data['aau_port']
    aau_map = {}
    for _, r in aau.iterrows():
        du      = str(r.iloc[3])  if pd.notna(r.iloc[3])  else ''
        rf      = str(r.iloc[5])  if pd.notna(r.iloc[5])  else ''
        cell_id = str(r.iloc[4])  if pd.notna(r.iloc[4])  else ''
        eqp_nm  = str(r.iloc[12]) if pd.notna(r.iloc[12]) else ''
        label   = str(r.iloc[11]) if pd.notna(r.iloc[11]) else ''
        loc = eqp_nm if eqp_nm not in ('nan','%%빈포트%%','') else label
        if du and rf: aau_map[(du, rf)] = (cell_id, loc)

    for _, r in df5_v.iterrows():
        du  = str(r.get('DU','')) if pd.notna(r.get('DU','')) else ''
        fru = str(r['FRU']) if r['FRU'] else ''
        vswr = str(r.get('VSWR','')).strip()
        cell_id, loc = aau_map.get((du, fru), ('',''))
        results.append({'시스템':'5G','기지국명':str(r.get('기지국명','')),'DU':du,
                        'Cell':cell_id,'국소명':loc,'RF_PORT':fru,
                        'VSWR':vswr,'발생시각':r.get('eventTime','')})

    df_lte   = _data['lte_raw'].copy()
    if 'specificProblem' not in df_lte.columns:
        df_lte_v = pd.DataFrame()
    else:
        df_lte_v = df_lte[df_lte['specificProblem'] == 'VSWR Over Threshold'].copy()
    if not df_lte_v.empty and 'objectOfReference' in df_lte_v.columns:
        df_lte_v['FRU'] = df_lte_v['objectOfReference'].apply(parse_fru)
    else:
        df_lte_v = pd.DataFrame(columns=['DU','FRU','VSWR','기지국명','eventTime'])

    rru     = _data['rru_port']
    rru_map = {}
    for _, r in rru.iterrows():
        du      = str(r.iloc[3])  if pd.notna(r.iloc[3])  else ''
        rf      = str(r.iloc[5])  if pd.notna(r.iloc[5])  else ''
        cell_id = str(r.iloc[4])  if pd.notna(r.iloc[4])  else ''
        ru_name = str(r.iloc[10]) if pd.notna(r.iloc[10]) else ''
        if du and rf: rru_map[(du, rf)] = (cell_id, ru_name)

    for _, r in df_lte_v.iterrows():
        du  = str(r.get('DU', r.iloc[1])) if pd.notna(r.get('DU', r.iloc[1])) else ''
        fru = str(r['FRU']) if r['FRU'] else ''
        vswr = str(r.get('VSWR','')).strip()
        cell_id, loc = rru_map.get((du, fru), ('',''))
        results.append({'시스템':'LTE','기지국명':str(r.get('기지국명','')),'DU':du,
                        'Cell':cell_id,'국소명':loc,'RF_PORT':fru,
                        'VSWR':vswr,'발생시각':r.get('eventTime','')})

    df_vswr = pd.DataFrame(results)
    df_vswr['VSWR_num'] = pd.to_numeric(df_vswr['VSWR'], errors='coerce')
    return df_vswr.sort_values('VSWR_num', ascending=False).reset_index(drop=True)


def process_oos(df, today_date):
    df = df.copy()
    df['발생시각_dt'] = pd.to_datetime(df['발생시각'], errors='coerce')
    df = df.drop_duplicates(subset=['Port', '발생시각'], keep='last').copy()

    today_ports = set(df[df['본부'].dt.date == today_date]['Port'].dropna().astype(str))
    port_latest = df.groupby('Port')['발생시각_dt'].max()
    port_count  = df.groupby('Port')['발생시각'].nunique()

    def get_status(row):
        port = str(row['Port']) if pd.notna(row['Port']) else ''
        if port not in today_ports: return '복구'
        latest = port_latest.get(port)
        if pd.notna(latest) and row['발생시각_dt'] == latest: return '미복구'
        return '복구'

    df['_복구여부'] = df.apply(get_status, axis=1)
    df['_다발횟수'] = df['Port'].map(port_count).fillna(1).astype(int)
    df['_다발']    = df['_다발횟수'].apply(lambda x: f'다발 {x}회' if x > 1 else '-')
    return df.sort_values('발생시각_dt', ascending=False).drop(columns=['발생시각_dt'])


# ── 헬퍼 ──────────────────────────────────────
def safe(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return ''
    return str(val).strip()

def sys_badge(sys):
    color = '#1971c2' if sys == '5G' else ('#e67700' if sys == 'LTE' else '#555')
    return f'<span class="b" style="color:{color};background:rgba(128,128,128,.12)">{sys}</span>'

def gj_badge(gj):
    s = safe(gj)
    if not s or s == 'nan': return ''
    su = s.upper()
    cls = 'bb' if su in ('BP','UNIT') else ('bg' if s == '복구' else ('by' if s == '점검중' else ('bo' if s == '재발생' else 'bgr')))
    return f'<span class="b {cls}">{s}</span>'

def render_generic_cards(df, title_col, sub_cols, max_cards=120):
    total = len(df)
    if total == 0:
        st.info("데이터 없음")
        return
    if total > max_cards:
        st.caption(f"⚠️ 상위 {max_cards}건 표시 중 (전체 {total}건)")
    for _, row in df.head(max_cards).iterrows():
        title = safe(row.get(title_col, '')) or '(이름 없음)'
        parts = []
        for col in sub_cols:
            v = safe(row.get(col, ''))
            if v and v != 'nan':
                parts.append(f'<span style="opacity:.5;font-size:11px">{col}</span>&nbsp;<span style="font-size:12px">{v}</span>')
        body = '&nbsp;&nbsp;'.join(parts)
        st.markdown(
            f'<div class="ac"><strong style="font-size:13px">{title}</strong>'
            f'<div style="margin-top:4px;line-height:1.8">{body}</div></div>',
            unsafe_allow_html=True
        )

def plotly_layout(fig, h=260):
    fig.update_layout(
        margin=dict(t=20, b=20, l=0, r=0),
        height=h,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
    )
    return fig

def overdue_note(df, date_col, name_col, days=5, extra_col=None):
    if df is None or df.empty: return
    if date_col not in df.columns or name_col not in df.columns: return

    today_ts = pd.Timestamp(date.today())
    sample = df[date_col].dropna().iloc[0] if not df[date_col].dropna().empty else None
    if sample is None: return

    try:
        if isinstance(sample, (int, float)) or str(sample).replace('.','').isdigit():
            df2 = df.copy()
            df2['_days'] = pd.to_numeric(df[date_col], errors='coerce')
        else:
            df2 = df.copy()
            df2['_dt'] = pd.to_datetime(df[date_col], errors='coerce')
            df2['_days'] = (today_ts - df2['_dt']).dt.days
    except Exception:
        return

    long = df2[df2['_days'] >= days].copy()
    if long.empty: return

    parts = []
    for _, r in long.head(15).iterrows():
        nm = safe(r.get(name_col, ''))
        d  = int(r['_days']) if pd.notna(r['_days']) else 0
        ex = f" {safe(r.get(extra_col,''))}" if extra_col and safe(r.get(extra_col,'')) else ''
        parts.append(f"{nm}{ex}({d}일)")

    note = "  ＊5일↑ 경과: " + " / ".join(parts)
    if len(long) > 15:
        note += f" 외 {len(long)-15}건"
    st.markdown(
        f'<div style="font-size:11px;opacity:.55;margin:-4px 0 8px;line-height:1.6">{note}</div>',
        unsafe_allow_html=True
    )


def kpi_grid(items):
    n = len(items)
    labels = ''.join(
        f'<div style="font-size:.68rem;opacity:.55;font-weight:500;'
        f'padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">'
        f'{lbl}</div>'
        for lbl, _, _ in items
    )
    values = ''.join(
        f'<div style="font-size:1.45rem;font-weight:700;color:{clr};padding-top:6px;line-height:1.1">'
        f'{val}</div>'
        for _, val, clr in items
    )
    return (
        f'<div style="background:rgba(128,128,128,.05);border:1px solid rgba(128,128,128,.15);'
        f'border-radius:10px;padding:12px 10px;margin-bottom:10px">'
        f'<div style="display:grid;grid-template-columns:repeat({n},1fr);text-align:center">'
        f'{labels}{values}</div></div>'
    )


def render_oos_table(df, col='알람등급'):
    """알람등급이 OOS인 행을 빨간색으로 강조한 HTML 테이블"""
    if df is None or df.empty:
        st.info("데이터 없음")
        return
    df = df.copy().reset_index(drop=True)
    th = ''.join(
        f'<th style="padding:5px 9px;text-align:left;font-size:11px;'
        f'opacity:.65;border-bottom:2px solid rgba(128,128,128,.2);'
        f'white-space:nowrap">{c}</th>'
        for c in df.columns
    )
    rows = []
    for _, row in df.iterrows():
        is_oos = col in df.columns and str(row.get(col, '')).strip() == 'OOS'
        rs = 'background:rgba(240,62,62,.13);' if is_oos else ''
        fc = 'color:#f03e3e;font-weight:600;' if is_oos else ''
        cells = ''.join(
            f'<td style="padding:4px 9px;font-size:12px;{fc}'
            f'border-bottom:1px solid rgba(128,128,128,.07);white-space:nowrap">'
            f'{("" if (v != v or v is None) else str(v))}</td>'
            for v in row.values
        )
        rows.append(f'<tr style="{rs}">{cells}</tr>')
    html = (
        '<div style="overflow-x:auto;margin-bottom:8px">'
        '<table style="width:100%;border-collapse:collapse;font-size:12px">'
        f'<thead><tr>{th}</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        '</table></div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ════════════════════════════════════════════
#  데이터 로드
# ════════════════════════════════════════════
ch, cb = st.columns([5, 1])
ch.markdown(
    '<div style="padding:14px 0 4px">'
    '<span style="font-size:.78rem;opacity:.6;font-weight:500">🚨 진주품질개선팀</span><br>'
    '<span style="font-size:1.2rem;font-weight:900;line-height:1.3">고장 현황 대시보드</span>'
    '</div>',
    unsafe_allow_html=True
)
if cb.button("🔄 새로고침", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

with st.spinner("📡 데이터 로딩 중..."):
    try:
        data       = load_data(DATA_URL)
        df_oos     = data['oos']
        df_vswr    = build_vswr(data)
        alarm_map  = build_alarm_map(data)
    except Exception as e:
        st.error(f"❌ 데이터 로드 실패: {e}")
        st.stop()

# 금일 날짜 = 데이터 내 최신 날짜
_valid_dates = df_oos['본부'].dropna().dt.date
today = _valid_dates.max() if not _valid_dates.empty else date.today()
df_today = df_oos[df_oos['본부'].dt.date == today].copy()
today_str = str(today)

# ── 화면·다운로드 공용: 각 Port의 실제 알람명 미리 계산 ──
def _calc_alarm(p):
    real = get_real_alarm(str(p), alarm_map)
    if real:
        return real
    _sub = df_today[df_today['Port'].astype(str) == str(p)]
    if '5G Link 알람' in df_today.columns and len(_sub) > 0:
        lv = safe(_sub['5G Link 알람'].iloc[0])
        if lv and lv not in ('nan', ''):
            return lv.split(':', 1)[-1].strip() if ':' in lv else lv
    if '세부알람' in df_today.columns and len(_sub) > 0:
        sv = safe(_sub['세부알람'].iloc[0])
        if sv and sv not in ('Service Unavailable', 'nan', ''):
            return sv
    return ''
df_today['_alarm'] = df_today['Port'].apply(_calc_alarm)

# 오늘 알람을 히스토리 파일에 저장
_alarm_hist = load_alarm_history()
_today_key  = today.strftime('%Y-%m-%d')
_today_alarm_dict = {
    str(row['Port']): row['_alarm']
    for _, row in df_today.iterrows()
    if row.get('_alarm')
}
if _today_alarm_dict:
    _alarm_hist[_today_key] = _today_alarm_dict
    # Port 단독 저장 (날짜 무관 영구 보존 - 전체 RAW 역대 알람 표시용)
    _ports_store = _alarm_hist.setdefault('_ports', {})
    _ports_store.update(_today_alarm_dict)
    save_alarm_history(_alarm_hist)

# 전체 RAW 중복제거 + 복구여부
df_processed = process_oos(df_oos, today)

# 전체 RAW에 알람 컬럼 추가
# 우선순위: ① 저장된 히스토리 ② OOS시트 5G Link 알람 ③ OOS시트 세부알람
_hist_updated = False

def _get_hist_alarm(row):
    global _hist_updated
    port = str(row.get('Port', ''))
    date_val = row.get('본부')
    if not port or pd.isna(date_val):
        return ''
    try:
        dk = pd.Timestamp(date_val).strftime('%Y-%m-%d')
        # ① 날짜별 히스토리
        cached = _alarm_hist.get(dk, {}).get(port, '')
        if cached:
            return cached
        # ② Port 단독 히스토리 (금일알람에서 쌓인 영구 데이터)
        port_cached = _alarm_hist.get('_ports', {}).get(port, '')
        if port_cached:
            return port_cached
        # ③ OOS 시트 5G Link 알람 컬럼
        lv = str(row.get('5G Link 알람', '')).strip()
        if lv and lv not in ('nan', '', 'None', 'Service Unavailable'):
            alarm = lv.split(':', 1)[-1].strip() if ':' in lv else lv
            if alarm:
                _alarm_hist.setdefault('_ports', {})[port] = alarm
                _alarm_hist.setdefault(dk, {})[port] = alarm
                _hist_updated = True
                return alarm
        # ④ 세부알람 컬럼
        sv = str(row.get('세부알람', '')).strip()
        if sv and sv not in ('nan', '', 'None', 'Service Unavailable'):
            _alarm_hist.setdefault('_ports', {})[port] = sv
            _alarm_hist.setdefault(dk, {})[port] = sv
            _hist_updated = True
            return sv
    except Exception:
        pass
    return ''

df_processed['알람'] = df_processed.apply(_get_hist_alarm, axis=1)
# 새로 찾은 알람이 있으면 히스토리 파일 저장
if _hist_updated:
    save_alarm_history(_alarm_hist)


tab1,tab2,tab3,tab4,tab5,tab6,tab7,tab8,tab9 = st.tabs([
    "🔔 금일 알람", "🕐 13시 알람", "📋 전체 RAW", "📊 대시보드",
    "📡 중계기·MIBOS", "🔔 기타 알람", "📶 VSWR 알람",
    "🔦 LTE광레벨", "🔦 5G광레벨"
])


# ══════════════════════════════════════════════════
#  TAB1  금일 알람
# ══════════════════════════════════════════════════
with tab1:
    st.caption(f"📅 기준일: **{today_str}**")

    if df_today.empty:
        st.warning("금일 알람 데이터가 없습니다.")
    else:
        total = len(df_today)

        # 직전 업로드 날짜 찾기 (두 번째로 최근 날짜)
        _all_dates = sorted(df_oos['본부'].dropna().dt.date.unique())
        _prev_date = _all_dates[-2] if len(_all_dates) >= 2 else None
        _df_prev   = df_oos[df_oos['본부'].dt.date == _prev_date] if _prev_date else pd.DataFrame()
        _prev_ports = set(_df_prev['Port'].dropna().astype(str)) if not _df_prev.empty else set()

        _today_ports = df_today['Port'].astype(str)

        # 신규: 오늘 있지만 직전 날짜에 없던 Port
        new_cnt = int((~_today_ports.isin(_prev_ports)).sum())

        # 점검중: 고장구분에 내용이 있는 국소
        if '고장구분' in df_today.columns:
            _has_gj = (df_today['고장구분'].notna() &
                       ~df_today['고장구분'].astype(str).str.strip().isin(['', 'nan']))
            chk_cnt = int(_has_gj.sum())
        else:
            _has_gj = pd.Series([False] * total, index=df_today.index)
            chk_cnt = 0

        # 미점검: 직전 날짜에 있었고 + 고장구분 비어있는 국소
        _from_prev = _today_ports.isin(_prev_ports)
        _no_gj     = ~_has_gj
        mij_cnt    = int((_from_prev & _no_gj).sum())

        # KPI 4칸 (기타 제거)
        st.markdown(f"""
<div style="background:rgba(128,128,128,.05);border:1px solid rgba(128,128,128,.15);
            border-radius:10px;padding:12px 10px;margin-bottom:10px">
  <div style="display:grid;grid-template-columns:repeat(4,1fr);text-align:center;gap:0">
    <div style="font-size:.7rem;opacity:.55;font-weight:500;padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">전체</div>
    <div style="font-size:.7rem;opacity:.55;font-weight:500;padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">신규</div>
    <div style="font-size:.7rem;opacity:.55;font-weight:500;padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">점검중</div>
    <div style="font-size:.7rem;opacity:.55;font-weight:500;padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">미점검</div>
    <div style="font-size:1.6rem;font-weight:700;color:#1971c2;padding-top:6px">{total}</div>
    <div style="font-size:1.6rem;font-weight:700;color:#f03e3e;padding-top:6px">{new_cnt}</div>
    <div style="font-size:1.6rem;font-weight:700;color:#e67700;padding-top:6px">{chk_cnt}</div>
    <div style="font-size:1.6rem;font-weight:700;color:#9c36b5;padding-top:6px">{mij_cnt}</div>
  </div>
</div>
""", unsafe_allow_html=True)

        # 다운로드 버튼 (화면과 동일: 국소명 / Cell / 알람)
        _dl = df_today.copy()
        _nm_col = next((c for c in ['장비명','국소명','사이트명'] if c in _dl.columns), None)
        _dl_out = pd.DataFrame({
            '국소명': _dl[_nm_col] if _nm_col else '',
            'Cell':   _dl['Port'],
            '알람':   _dl['_alarm'].replace('', '확인필요'),
        })
        st.download_button(
            label=f"⬇️ 금일 알람 다운로드 ({today_str})",
            data=_dl_out.to_csv(index=False).encode('utf-8-sig'),
            file_name=f"금일알람_{today_str}.csv",
            mime="text/csv",
            key="dl_today",
            use_container_width=True,
        )
        st.markdown("---")

        # 필터
        cs, csys = st.columns([3, 1])
        search  = cs.text_input("검색", key="t1_s", placeholder="🔍 장비명·Port·시군구", label_visibility="collapsed")
        sys_flt = csys.selectbox("시스템", ['전체','5G','LTE'], key="t1_sys", label_visibility="collapsed")

        df_show = df_today.copy()
        if search:
            df_show = df_show[df_show.apply(lambda r: search.lower() in str(r).lower(), axis=1)]
        if sys_flt != '전체':
            df_show = df_show[df_show['시스템'] == sys_flt]

        st.caption(f"표시 {len(df_show)}건")
        overdue_note(df_today, '본부', '장비명', days=5, extra_col='시군구')

        ALARM_COLOR = {
            'Link Failure':'#f03e3e','No Connection':'#e03131',
            'Input Power Failure':'#e67700','SW Error':'#1971c2',
            'FRU General Problem':'#7048e8',
        }
        cards_html = []
        for _, row in df_show.iterrows():
            gj      = safe(row.get('고장구분'))
            sys_nm  = safe(row.get('시스템'))
            device  = safe(row.get('장비명'))
            sigungu = safe(row.get('시군구'))
            port    = safe(row.get('Port'))
            detail  = safe(row.get('복구/미복구 상세내역'))
            occur   = safe(row.get('발생시각'))

            alarm_disp = safe(row.get('_alarm')) or safe(row.get('세부알람'))
            a_color    = ALARM_COLOR.get(alarm_disp, '#868e96')
            alarm_b    = (f'<span class="b" style="background:rgba(128,128,128,.1);color:{a_color}">'
                          f'{alarm_disp}</span>') if alarm_disp else ''

            gj_up    = gj.upper()
            card_cls = ('bp'   if gj_up in ('BP','UNIT') else
                        'done' if gj == '복구' else
                        'rep'  if gj == '점검중' else
                        'new'  if not gj else 'other')
            det_html = (f'<div style="font-size:11px;opacity:.6;margin-top:3px">📝 {detail}</div>'
                        ) if detail else ''

            cards_html.append(
                f'<div class="ac {card_cls}">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:2px">'
                f'<span>{sys_badge(sys_nm)}&nbsp;<strong style="font-size:13px">{device}</strong></span>'
                f'<span style="font-size:11px;opacity:.55">{sigungu}</span></div>'
                f'<div style="font-size:12px;margin-top:3px">🔌&nbsp;<span class="ptag">{port}</span></div>'
                f'<div style="margin-top:4px">{alarm_b}</div>'
                f'{det_html}'
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-top:4px">'
                f'<span>{gj_badge(gj)}</span>'
                f'<span style="font-size:10px;opacity:.4">🕐 {occur}</span></div>'
                f'</div>'
            )
        st.markdown(''.join(cards_html), unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  TAB2  13시 알람
# ══════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="st">🕐 13시 알람 공유</div>', unsafe_allow_html=True)
    df_13 = data['alarm_13'].dropna(how='all').copy()
    df_13.columns = [str(c).strip() for c in df_13.columns]
    st.caption(f"총 {len(df_13)}건")

    if df_13.empty:
        st.info("13시 알람 데이터가 없습니다.")
    else:
        cards13 = []
        for _, row in df_13.iterrows():
            # 장비명: A열(0) 또는 '장비명' 컬럼
            device13 = safe(row.get('장비명', '')) or safe(row.iloc[0] if len(row) > 0 else '')
            # 알람명: F열(5) 또는 '알람' 컬럼
            alarm13 = safe(row.get('알람', ''))
            if not alarm13 and len(row) > 5:
                alarm13 = safe(row.iloc[5])
            # 발생시각: '발생시각' 컬럼 또는 찾기
            occur13 = safe(row.get('발생시각', ''))
            if not occur13:
                for cn in row.index:
                    v = safe(row[cn])
                    if v and ('시각' in str(cn) or '시간' in str(cn) or '발생' in str(cn)):
                        occur13 = v
                        break
            # 지역
            region13 = safe(row.get('시군구', '')) or safe(row.get('지역', ''))

            alarm_b13 = (f'<span class="b br">{alarm13}</span>') if alarm13 else ''
            region_b  = (f'<span style="font-size:11px;opacity:.55">{region13}</span>') if region13 else ''
            occur_b   = (f'<span style="font-size:10px;opacity:.4">🕐 {occur13}</span>') if occur13 else ''

            cards13.append(
                f'<div class="ac new">'
                f'<div style="display:flex;justify-content:space-between;align-items:center">'
                f'<strong style="font-size:13px">{device13}</strong>{region_b}</div>'
                f'<div style="margin-top:5px">{alarm_b13}</div>'
                f'<div style="margin-top:4px;text-align:right">{occur_b}</div>'
                f'</div>'
            )
        st.markdown(''.join(cards13), unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  TAB3  전체 RAW
# ══════════════════════════════════════════════════
with tab3:
    st.caption("※ (Port+발생시각) 중복 제거 / 금일 알람 기준 복구여부 자동 계산")

    unfix_raw = int((df_processed['_복구여부'] == '미복구').sum())
    done_raw  = int((df_processed['_복구여부'] == '복구').sum())

    st.markdown(kpi_grid([
        ("전체",   len(df_processed), '#1971c2'),
        ("미복구", unfix_raw,          '#f03e3e'),
        ("복구",   done_raw,           '#2f9e44'),
    ]), unsafe_allow_html=True)
    overdue_note(df_processed[df_processed['_복구여부']=='미복구'], '본부', '장비명', days=5, extra_col='시군구')

    cs3, cf3 = st.columns([3, 1])
    search3 = cs3.text_input("검색", key="t3_s", placeholder="🔍 Port·장비명·시군구", label_visibility="collapsed")
    filter3 = cf3.selectbox("필터", ['전체','미복구만','5G','LTE'], key="t3_f", label_visibility="collapsed")

    base_cols = ['발생시각','장비명','Port','시스템','시군구','고장구분','고장구분(중분류)','_다발','_복구여부','복구/미복구 상세내역','알람']
    disp_cols = [c for c in base_cols if c in df_processed.columns]

    df_d3 = df_processed.copy()
    if search3:
        df_d3 = df_d3[df_d3.apply(lambda r: search3.lower() in str(r).lower(), axis=1)]
    if filter3 == '미복구만':
        df_d3 = df_d3[df_d3['_복구여부'] == '미복구']
    elif filter3 in ('5G','LTE'):
        df_d3 = df_d3[df_d3['시스템'] == filter3]

    st.caption(f"표시 {len(df_d3)}건")
    st.dataframe(df_d3[disp_cols], use_container_width=True, height=500)


# ══════════════════════════════════════════════════
#  TAB4  대시보드
# ══════════════════════════════════════════════════
with tab4:
    st.caption(f"📅 기준일: **{today_str}**")
    df_d = df_processed.copy()

    total_cnt = len(df_d)
    unfix_cnt = int((df_d['_복구여부'] == '미복구').sum())
    done_cnt  = int((df_d['_복구여부'] == '복구').sum())
    dabal_cnt = int((df_d['_다발횟수'] > 1).sum()) if '_다발횟수' in df_d.columns else 0
    g5_cnt    = int((df_d['시스템'] == '5G').sum())  if '시스템' in df_d.columns else 0
    lte_cnt   = int((df_d['시스템'] == 'LTE').sum()) if '시스템' in df_d.columns else 0

    st.markdown(kpi_grid([
        ("전체 고장", total_cnt, '#1971c2'),
        ("미복구",   unfix_cnt,  '#f03e3e'),
        ("복구",     done_cnt,   '#2f9e44'),
        ("다발알람", dabal_cnt,  '#e67700'),
        ("5G 고장",  g5_cnt,    '#7048e8'),
        ("LTE 고장", lte_cnt,   '#f59f00'),
    ]), unsafe_allow_html=True)
    overdue_note(df_d[df_d['_복구여부']=='미복구'], '본부', '장비명', days=5, extra_col='시군구')

    cmap = {'미복구':'#f03e3e','복구':'#2f9e44'}

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        st.markdown('<div class="st">복구 / 미복구 현황</div>', unsafe_allow_html=True)
        sc = df_d['_복구여부'].value_counts().reset_index()
        sc.columns = ['구분','건수']
        fig = px.pie(sc, values='건수', names='구분', hole=0.55,
                     color='구분', color_discrete_map=cmap)
        fig.update_traces(textfont_size=13)
        fig.update_layout(legend=dict(orientation='v', font=dict(size=11)))
        st.plotly_chart(plotly_layout(fig, 260), use_container_width=True)

    with r1c2:
        st.markdown('<div class="st">5G / LTE 복구여부 비교</div>', unsafe_allow_html=True)
        if '시스템' in df_d.columns:
            sg   = df_d.groupby(['시스템','_복구여부']).size().reset_index(name='건수')
            fig3 = px.bar(sg, x='시스템', y='건수', color='_복구여부',
                          barmode='group', color_discrete_map=cmap)
            fig3.update_layout(xaxis_title='', yaxis_title='건수', legend_title='')
            st.plotly_chart(plotly_layout(fig3, 260), use_container_width=True)

    # ── 작업 유형 분류 함수 ──
    WORK_KW = [
        ('자동복구',  '자동복구'),
        ('AAU',      'AAU 교체/수리'),
        ('정류기',   '정류기 교체/수리'),
        ('SFP',      'SFP 교체'),
        ('광점퍼',   '광점퍼 교체'),
        ('광케이블', '광케이블 교체'),
        ('PSU',      'PSU 교체/수리'),
        ('PRU',      'PRU 교체/수리'),
        ('공중선',   '공중선 작업'),
        ('차단기',   '차단기'),
        ('전원',     '전원 관련'),
        ('교체',     '기타 교체'),
        ('복구',     '기타 복구'),
    ]
    def classify_work(txt):
        s = str(txt).strip()
        if not s or s in ('nan', '-', 'None', '미복구', ''):
            return None
        for kw, label in WORK_KW:
            if kw in s:
                return label
        return '기타'

    detail_col = '복구/미복구 상세내역'
    if detail_col in df_d.columns:
        wk_series = df_d[detail_col].apply(classify_work).dropna()
        wk_cnt = wk_series.value_counts().reset_index()
        wk_cnt.columns = ['작업유형', '건수']
    else:
        wk_cnt = pd.DataFrame(columns=['작업유형','건수'])

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        st.markdown('<div class="st">작업내역별 현황</div>', unsafe_allow_html=True)
        if not wk_cnt.empty:
            fig_wk = px.bar(
                wk_cnt.sort_values('건수'),
                y='작업유형', x='건수', orientation='h',
                color='건수',
                color_continuous_scale=['#74c0fc','#1971c2'],
                text='건수',
            )
            fig_wk.update_traces(textposition='outside', textfont_size=11)
            fig_wk.update_layout(
                showlegend=False, coloraxis_showscale=False,
                yaxis_title='', xaxis_title='건수',
                margin=dict(l=10, r=40, t=10, b=10),
            )
            st.plotly_chart(plotly_layout(fig_wk, 320), use_container_width=True)
        else:
            st.info("작업내역 데이터 없음")

    with r2c2:
        st.markdown('<div class="st">장비 중분류별 고장</div>', unsafe_allow_html=True)
        mc = '고장구분(중분류)'
        if mc in df_d.columns:
            mid = df_d[mc].dropna()
            mid = mid[~mid.astype(str).str.strip().isin(['nan',''])].value_counts().reset_index()
            mid.columns = ['중분류','건수']
            fig4 = px.bar(mid, y='중분류', x='건수', orientation='h',
                          color_discrete_sequence=['#1971c2'],
                          text='건수')
            fig4.update_traces(textposition='outside', textfont_size=11)
            fig4.update_layout(yaxis_title='', xaxis_title='건수',
                               margin=dict(l=10, r=40, t=10, b=10))
            st.plotly_chart(plotly_layout(fig4, 320), use_container_width=True)


# ══════════════════════════════════════════════════
#  TAB5  중계기·MIBOS
# ══════════════════════════════════════════════════
with tab5:
    df_rel = data['relay']
    df_gr  = data['grems']
    legacy = int((df_rel['시스템 구분'].astype(str).str.contains('Legacy', na=False)).sum()) if '시스템 구분' in df_rel.columns else 0
    mibos  = int((df_rel['시스템 구분'].astype(str).str.contains('MIBOS',  na=False)).sum()) if '시스템 구분' in df_rel.columns else 0

    st.markdown(kpi_grid([
        ("LEGACY 중계기", legacy,      '#e67700'),
        ("MIBOS",         mibos,       '#1971c2'),
        ("gREMS",         len(df_gr),  '#7048e8'),
    ]), unsafe_allow_html=True)
    _rel_date = next((c for c in ['경과일','발생일','장애일','날짜'] if c in df_rel.columns), None)
    _rel_name = next((c for c in ['중계기명','장비명','국소명'] if c in df_rel.columns), None)
    if _rel_date and _rel_name:
        overdue_note(df_rel, _rel_date, _rel_name, days=5, extra_col='시군구' if '시군구' in df_rel.columns else None)

    st.markdown('<div class="st">중계기 및 MIBOS 알람</div>', unsafe_allow_html=True)
    render_oos_table(df_rel)

    st.markdown('<div class="st">gREMS 알람</div>', unsafe_allow_html=True)
    render_oos_table(df_gr)


# ══════════════════════════════════════════════════
#  TAB6  기타 알람
# ══════════════════════════════════════════════════
with tab6:
    for title, key in [('RMS A망 미복구','rms_a'),('RMS DACS 미복구','rms_dacs'),('RMS 통합RCU 미복구','rms_rcu')]:
        df_t = data[key]
        st.markdown(f'<div class="st">{title}</div>', unsafe_allow_html=True)
        if df_t is not None and not df_t.empty:
            st.caption(f"총 {len(df_t)}건")
            _t_date = next((c for c in ['경과일','발생일','장애일','알람발생일','날짜'] if c in df_t.columns), None)
            _t_name = next((c for c in ['장비명','중계기명','국소명','RCU명','사이트명'] if c in df_t.columns), None)
            if _t_date and _t_name:
                overdue_note(df_t, _t_date, _t_name, days=5)
            render_oos_table(df_t)
        else:
            st.info("데이터 없음")


# ══════════════════════════════════════════════════
#  TAB7  VSWR 알람
# ══════════════════════════════════════════════════
with tab7:
    if df_vswr.empty:
        st.warning("VSWR 알람 데이터가 없습니다.")
    else:
        g5v  = int((df_vswr['시스템'] == '5G').sum())
        ltev = int((df_vswr['시스템'] == 'LTE').sum())
        mtch = int((df_vswr['Cell'] != '').sum())

        st.markdown(kpi_grid([
            ("전체 VSWR", len(df_vswr), '#1971c2'),
            ("5G",         g5v,         '#7048e8'),
            ("LTE",        ltev,        '#f59f00'),
            ("국소 매칭",  mtch,        '#2f9e44'),
        ]), unsafe_allow_html=True)

        cv1, cv2, cv3 = st.columns([3, 1, 1])
        sv    = cv1.text_input("검색", key="t7_s", placeholder="🔍 기지국명·Cell·국소명", label_visibility="collapsed")
        sys_v = cv2.selectbox("시스템", ['전체','5G','LTE'], key="t7_sys", label_visibility="collapsed")
        vmin  = cv3.number_input("최솟값", min_value=1.0, max_value=5.0, value=1.0, step=0.1, key="t7_min")

        dfv = df_vswr.copy()
        if sv:
            dfv = dfv[dfv[['기지국명','Cell','국소명','DU']].apply(lambda r: sv.lower() in str(r).lower(), axis=1)]
        if sys_v != '전체':
            dfv = dfv[dfv['시스템'] == sys_v]
        if vmin > 1.0:
            dfv = dfv[dfv['VSWR_num'] >= vmin]

        st.caption(f"총 {len(dfv)}건")
        show_v = dfv[['시스템','기지국명','Cell','국소명','RF_PORT','VSWR','발생시각']].copy()

        def hi_vswr(val):
            try:
                v = float(str(val).strip())
                if v >= 2.0:   return 'background-color:rgba(240,62,62,.2);color:#f03e3e;font-weight:700'
                elif v >= 1.5: return 'background-color:rgba(245,159,0,.15);color:#e67700'
            except:
                pass
            return ''

        st.dataframe(show_v.style.map(hi_vswr, subset=['VSWR']), use_container_width=True, height=450)

        st.markdown('<div class="st">VSWR 값 분포</div>', unsafe_allow_html=True)
        fig_v = px.histogram(dfv.dropna(subset=['VSWR_num']), x='VSWR_num', color='시스템', nbins=20,
                             color_discrete_map={'5G':'#1971c2','LTE':'#f59f00'},
                             labels={'VSWR_num':'VSWR','count':'건수'})
        fig_v.update_layout(bargap=0.1)
        st.plotly_chart(plotly_layout(fig_v, 230), use_container_width=True)


# ══════════════════════════════════════════════════
#  TAB8  LTE 광레벨 불량
# ══════════════════════════════════════════════════
with tab8:
    df_lopt = data.get('lte_optical', pd.DataFrame())
    if df_lopt.empty:
        st.info("LTE광레벨불량 시트가 없거나 데이터가 없습니다.")
    else:
        df_lopt.columns = [str(c).strip() for c in df_lopt.columns]
        # 주요 컬럼
        LM = ['RU명','DU TX','DU RX','DU TXBS','RU TX','RU RX','RU TXBS']
        LD = ['DU 파장','du_sfp_vendor','RU 파장','ru_sfp_vendor']
        lm = [c for c in LM if c in df_lopt.columns]
        ld = [c for c in LD if c in df_lopt.columns]

        # 진단결과 필터
        if '진단결과' in df_lopt.columns:
            opts = ['전체'] + sorted(df_lopt['진단결과'].dropna().astype(str).unique().tolist())
            sel  = st.selectbox("진단결과 필터", opts, key="lopt_diag")
            df_lf = df_lopt if sel == '전체' else df_lopt[df_lopt['진단결과'].astype(str) == sel]
        else:
            df_lf = df_lopt

        st.caption(f"총 {len(df_lf)}건")

        for _, row in df_lf.iterrows():
            ru_nm = safe(row.get('RU명', '')) or '(이름없음)'
            vals  = '  '.join(
                f'{c}: {safe(row.get(c,""))}'
                for c in lm[1:] if safe(row.get(c,''))
            )
            label_l = f'{ru_nm}  {vals}' if vals else ru_nm

            with st.expander(label_l, expanded=False):
                # TX/RX 수치 (≤-19 빨강)
                num_l = []
                for c in lm[1:]:
                    v = safe(row.get(c,''))
                    try:
                        vc = '#f03e3e' if float(v) <= -19 else 'inherit'
                    except:
                        vc = 'inherit'
                    num_l.append(
                        f'<span style="display:inline-block;margin:2px 12px 2px 0;font-size:13px">'
                        f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;'
                        f'<b style="color:{vc}">{v}</b></span>'
                    )
                if num_l:
                    st.markdown('<div style="margin-bottom:8px">' + ''.join(num_l) + '</div>', unsafe_allow_html=True)
                # SFP / 파장 상세
                if ld:
                    parts_l = []
                    for c in ld:
                        v = safe(row.get(c,'—'))
                        parts_l.append(
                            f'<div style="margin:4px 0;font-size:13px">'
                            f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;&nbsp;'
                            f'<b>{v}</b></div>'
                        )
                    st.markdown(''.join(parts_l), unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  TAB9  5G 광레벨 불량
# ══════════════════════════════════════════════════
with tab9:
    df_5opt = data.get('5g_optical', pd.DataFrame())
    if df_5opt.empty:
        st.info("5G광레벨불량 시트가 없거나 데이터가 없습니다.")
    else:
        df_5opt.columns = [str(c).strip() for c in df_5opt.columns]
        # DUL명 컬럼 없으면 D열(iloc[3])로 대체
        if 'DUL명' not in df_5opt.columns and len(df_5opt.columns) > 3:
            df_5opt = df_5opt.copy()
            df_5opt['DUL명'] = df_5opt.iloc[:, 3].astype(str)

        FM = ['DUL명','DUH TX','DUH RX','DUL TX','DUL RX']
        FD = ['DUH SFP 제조사','DUH SFP 파장','DUL SFP 제조사','DUL SFP 파장']
        fm = [c for c in FM if c in df_5opt.columns]
        fd = [c for c in FD if c in df_5opt.columns]

        # 진단결과 필터
        if '진단결과' in df_5opt.columns:
            opts5 = ['전체'] + sorted(df_5opt['진단결과'].dropna().astype(str).unique().tolist())
            sel5  = st.selectbox("진단결과 필터", opts5, key="5opt_diag")
            df_5f = df_5opt if sel5 == '전체' else df_5opt[df_5opt['진단결과'].astype(str) == sel5]
        else:
            df_5f = df_5opt

        st.caption(f"총 {len(df_5f)}건")

        for _, row in df_5f.iterrows():
            dul_nm = safe(row.get('DUL명','')) or '(이름없음)'
            vals5  = '  '.join(
                f'{c}: {safe(row.get(c,""))}'
                for c in fm[1:] if safe(row.get(c,''))
            )
            label5 = f'{dul_nm}  {vals5}' if vals5 else dul_nm

            with st.expander(label5, expanded=False):
                # TX/RX 수치 (≤-19 빨강)
                num5 = []
                for c in fm[1:]:
                    v = safe(row.get(c,''))
                    try:
                        vc = '#f03e3e' if float(v) <= -19 else 'inherit'
                    except:
                        vc = 'inherit'
                    num5.append(
                        f'<span style="display:inline-block;margin:2px 12px 2px 0;font-size:13px">'
                        f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;'
                        f'<b style="color:{vc}">{v}</b></span>'
                    )
                if num5:
                    st.markdown('<div style="margin-bottom:8px">' + ''.join(num5) + '</div>', unsafe_allow_html=True)
                # SFP / 파장 상세
                if fd:
                    parts5 = []
                    for c in fd:
                        v = safe(row.get(c,'—'))
                        parts5.append(
                            f'<div style="margin:4px 0;font-size:13px">'
                            f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;&nbsp;'
                            f'<b>{v}</b></div>'
                        )
                    st.markdown(''.join(parts5), unsafe_allow_html=True)
