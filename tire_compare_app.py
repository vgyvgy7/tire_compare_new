"""
타이어 적재 비교 계산기 v2
규격 + 개수 → 평치(일반) / 벌집 3D 나란히 시각화
"""

import streamlit as st
import math, re, itertools
from pathlib import Path
from itertools import combinations
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ─── 컨테이너 치수 ────────────────────────────────────
CONT_W, CONT_L, CONT_H = 2352, 12032, 2698
N_SEG = 8  # 실린더 분할 수
DEFAULT_CONTAINER = "40ft HC"
CONTAINER_SPECS = {
    "40ft HC": {"w": 2352, "l": 12032, "h": 2698, "source": "엑셀 기준"},
    "40ft Dry": {"w": 2352, "l": 12032, "h": 2395, "source": "계산 기준"},
    "20ft Dry": {"w": 2352, "l": 5898, "h": 2395, "source": "계산 기준"},
    "45ft HC": {"w": 2352, "l": 13556, "h": 2698, "source": "계산 기준"},
    "1톤 카고": {"w": 1600, "l": 2850, "h": 1600, "source": "계산 기준"},
    "1톤 윙바디": {"w": 1600, "l": 2850, "h": 1800, "source": "계산 기준"},
    "2.5톤 카고": {"w": 1900, "l": 4300, "h": 1800, "source": "계산 기준"},
    "3.5톤 윙바디": {"w": 2100, "l": 4800, "h": 2100, "source": "계산 기준"},
    "5톤 카고": {"w": 2280, "l": 6200, "h": 2200, "source": "계산 기준"},
    "5톤 윙바디": {"w": 2280, "l": 6200, "h": 2300, "source": "계산 기준"},
    "11톤 윙바디": {"w": 2350, "l": 9100, "h": 2400, "source": "계산 기준"},
    "직접 입력": {"w": 2352, "l": 12032, "h": 2698, "source": "계산 기준"},
}
APP_DIR = Path(__file__).resolve().parent
LOAD_TABLE_CANDIDATES = [
    APP_DIR / "규격별 평치(일반,벌집) 적재 개수.xlsx",
    Path.home() / "Documents" / "KakaoTalk Downloads" / "규격별 평치(일반,벌집) 적재 개수.xlsx",
]

# ══════════════════════════════════════════════════════
# 1. 파싱
# ══════════════════════════════════════════════════════

def parse_tire_size(s):
    m = re.search(r'(\d+)[^\d]+(\d+)[^\d]+(\d+)', s)
    if not m:
        return None, None
    w, a, r = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return round((w * a / 100 * 2) + (r * 25.4)), w

@st.cache_data(show_spinner=False)
def load_capacity_table():
    for path in LOAD_TABLE_CANDIDATES:
        if path.exists():
            df = pd.read_excel(path)
            df = df.rename(columns=lambda c: str(c).strip())
            required = ["타이어규격", "일반평치", "벌집평치"]
            missing = set(required) - set(df.columns)
            if missing:
                raise ValueError(f"{path.name}에 필요한 컬럼이 없습니다: {', '.join(sorted(missing))}")
            df = df[required].copy()
            df["타이어규격"] = df["타이어규격"].astype(str).str.strip().str.upper()
            df["일반평치"] = pd.to_numeric(df["일반평치"], errors="coerce")
            df["벌집평치"] = pd.to_numeric(df["벌집평치"], errors="coerce")
            df = df.dropna(subset=["타이어규격", "일반평치", "벌집평치"])
            df["일반평치"] = df["일반평치"].astype(int)
            df["벌집평치"] = df["벌집평치"].astype(int)
            return df.sort_values("타이어규격").reset_index(drop=True), str(path)
    return pd.DataFrame(columns=["타이어규격", "일반평치", "벌집평치"]), None

# ══════════════════════════════════════════════════════
# 2. 적재 계산
# ══════════════════════════════════════════════════════

def calc_flat(od, h):
    flat_w  = CONT_W // od;  flat_h  = CONT_H // h;  main_l  = CONT_L // od
    w_rem   = CONT_W - flat_w * od
    stand_w = w_rem // h;    stand_h = CONT_H // od
    l_rem   = CONT_L - main_l * od
    extra_l = l_rem // h;    extra_w = CONT_W // od; extra_h = CONT_H // od
    flat_main  = flat_w  * flat_h  * main_l
    stand_main = stand_w * stand_h * main_l
    extra_main = extra_l * extra_w * extra_h
    return dict(
        flat_w=flat_w, flat_h=flat_h, main_l=main_l,
        stand_w=stand_w, stand_h=stand_h,
        w_rem=w_rem, l_rem=l_rem,
        extra_l=extra_l, extra_w=extra_w, extra_h=extra_h,
        flat_main=flat_main, stand_main=stand_main, extra_main=extra_main,
        total=flat_main + stand_main + extra_main
    )

def calc_honeycomb(od, h):
    STEP = math.sqrt(3) / 2 * od
    n    = max(1, 1 + int((CONT_W - od) / STEP))
    wu   = od + (n - 1) * STEP
    if wu > CONT_W:
        n -= 1; wu = od + (n - 1) * STEP
    wr = CONT_W - wu
    m  = CONT_L // od
    if od / 2 + m * od <= CONT_L:
        me = m;     lu = od / 2 + m * od
    else:
        me = m - 1; lu = float(m * od)
    lr = CONT_L - lu
    oc = math.ceil(n / 2);  ec = math.floor(n / 2)
    ol = oc * m + ec * me;  hc = CONT_H // h;  main = ol * hc
    ew, ewi = 0, {}
    if wr >= h:
        nw = int(wr // h); nl = int(CONT_L // od); nh = int(CONT_H // od)
        ew = nw * nl * nh;  ewi = dict(nw=nw, nl=nl, nh=nh)
    el, eli = 0, {}
    if lr >= h:
        nl2 = int(lr // h); nw2 = int(CONT_W // od); nh2 = int(CONT_H // od)
        el = nl2 * nw2 * nh2; eli = dict(nl=nl2, nw=nw2, nh=nh2)
    return dict(
        n=n, odd_col=oc, even_col=ec, m_odd=m, m_even=me, h_count=hc, one_layer=ol,
        width_used=wu, w_rem=wr, length_used=lu, l_rem=lr,
        main=main, extra_w=ew, extra_l=el, ew_info=ewi, el_info=eli,
        total=main + ew + el
    )

# ══════════════════════════════════════════════════════
# 3. 3D 위치 목록 생성
# ══════════════════════════════════════════════════════

def flat_positions(od, h, f):
    r = od / 2
    items = []
    # ① 메인 평치
    for layer in range(f['flat_h']):
        for col in range(f['flat_w']):
            for row in range(f['main_l']):
                items.append((col*od+r, row*od+r, layer*h, h, 'z'))
    # ② 폭 자투리 세움
    x0 = f['flat_w'] * od
    for layer in range(f['stand_h']):
        for col in range(f['stand_w']):
            for row in range(f['main_l']):
                items.append((x0+col*h, row*od+r, layer*od+r, h, 'x'))
    # ③ 길이 자투리 세움
    y0 = f['main_l'] * od
    for layer in range(f['extra_h']):
        for col in range(f['extra_w']):
            for row in range(f['extra_l']):
                items.append((col*od+r, y0+row*h, layer*od+r, h, 'y'))
    return items

def honeycomb_positions(od, h, hc):
    STEP = math.sqrt(3) / 2 * od
    r    = od / 2
    items = []
    # ① 메인 벌집
    for layer in range(hc['h_count']):
        cz = layer * h
        for col in range(hc['n']):
            cx    = col * STEP + r
            is_odd = (col % 2 == 1)
            oy    = od / 2 if is_odd else 0.0
            m_c   = hc['m_even'] if is_odd else hc['m_odd']
            for row in range(m_c):
                items.append((cx, oy + row*od + r, cz, h, 'z'))
    # ② 폭 자투리
    if hc['extra_w'] > 0:
        x0 = hc['width_used'];  d = hc['ew_info']
        for k in range(d['nh']):
            for i in range(d['nw']):
                for j in range(d['nl']):
                    items.append((x0+i*h, j*od+r, k*od+r, h, 'x'))
    # ③ 길이 자투리
    if hc['extra_l'] > 0:
        y0 = hc['length_used']; d = hc['el_info']
        for k in range(d['nh']):
            for i in range(d['nw']):
                for j in range(d['nl']):
                    items.append((i*od+r, y0+j*h, k*od+r, h, 'y'))
    return items

# ══════════════════════════════════════════════════════
# 4. 실린더 메시 배치
# ══════════════════════════════════════════════════════

def cylinder_batch(items, radius):
    n     = N_SEG
    theta = np.linspace(0, 2*np.pi, n, endpoint=False)
    ct, st_ = np.cos(theta), np.sin(theta)
    AX, AY, AZ, AI, AJ, AK = [], [], [], [], [], []
    off = 0
    for (cx, cy, cz, hl, axis) in items:
        if axis == 'z':
            xb = radius*ct+cx; yb = radius*st_+cy; zb = np.full(n, cz)
            xt = radius*ct+cx; yt = radius*st_+cy; zt = np.full(n, cz+hl)
            c0 = (cx, cy, cz);    c1 = (cx, cy, cz+hl)
        elif axis == 'x':
            xb = np.full(n, cx);    yb = radius*ct+cy; zb = radius*st_+cz
            xt = np.full(n, cx+hl); yt = radius*ct+cy; zt = radius*st_+cz
            c0 = (cx, cy, cz);    c1 = (cx+hl, cy, cz)
        else:
            xb = radius*ct+cx; yb = np.full(n, cy);    zb = radius*st_+cz
            xt = radius*ct+cx; yt = np.full(n, cy+hl); zt = radius*st_+cz
            c0 = (cx, cy, cz);    c1 = (cx, cy+hl, cz)
        xs = np.concatenate([xb, xt, [c0[0], c1[0]]])
        ys = np.concatenate([yb, yt, [c0[1], c1[1]]])
        zs = np.concatenate([zb, zt, [c0[2], c1[2]]])
        for v in range(n):
            vn = (v + 1) % n
            AI += [off+v, off+vn]; AJ += [off+vn, off+v+n]; AK += [off+v+n, off+vn+n]
        for v in range(n):
            AI.append(off+2*n); AJ.append(off+v);   AK.append(off+(v+1)%n)
        for v in range(n):
            AI.append(off+2*n+1); AJ.append(off+v+n); AK.append(off+(v+1)%n+n)
        AX.extend(xs); AY.extend(ys); AZ.extend(zs)
        off += 2*n + 2
    return np.array(AX), np.array(AY), np.array(AZ), AI, AJ, AK

# ══════════════════════════════════════════════════════
# 5. 컨테이너 와이어프레임 + 3D 빌드
# ══════════════════════════════════════════════════════

def container_wireframe():
    corners = np.array(list(itertools.product([0, CONT_W], [0, CONT_L], [0, CONT_H])))
    xw, yw, zw = [], [], []
    for s, e in combinations(corners, 2):
        if np.sum(np.abs(s - e)) in [CONT_W, CONT_L, CONT_H]:
            xw += [s[0], e[0], None]
            yw += [s[1], e[1], None]
            zw += [s[2], e[2], None]
    return go.Scatter3d(x=xw, y=yw, z=zw, mode='lines',
        line=dict(color='rgba(15,23,42,0.22)', width=1.5),
        name='컨테이너', showlegend=False, hoverinfo='none')

def build_3d(od, h, items, n_display, color_filled, color_empty,
             subtitle, cap, show_empty):
    r   = od / 2
    fig = go.Figure()
    fig.add_trace(container_wireframe())

    filled = items[:n_display]
    empty  = items[n_display:]

    if filled:
        x, y, z, i, j, k = cylinder_batch(filled, r)
        fig.add_trace(go.Mesh3d(
            x=x, y=y, z=z, i=i, j=j, k=k,
            color=color_filled, opacity=0.85,
            name=f'적재 {n_display:,}개', showlegend=True, flatshading=False,
            lighting=dict(ambient=0.45, diffuse=0.85, roughness=0.4, specular=0.4, fresnel=0.15),
            lightposition=dict(x=CONT_W*2, y=CONT_L, z=CONT_H*3)
        ))

    if show_empty and empty:
        x, y, z, i, j, k = cylinder_batch(empty, r)
        fig.add_trace(go.Mesh3d(
            x=x, y=y, z=z, i=i, j=j, k=k,
            color=color_empty, opacity=0.10,
            name=f'빈 자리 {len(empty):,}개', showlegend=True, flatshading=False,
            lighting=dict(ambient=0.3, diffuse=0.4)
        ))

    pct = n_display / cap * 100 if cap else 0
    fig.update_layout(
        title=dict(
            text=f'<b>{subtitle}</b>  ·  {n_display:,} / {cap:,}개  ({pct:.1f}%)',
            font=dict(size=12, family='Noto Sans KR, sans-serif', color='#0F172A')
        ),
        scene=dict(
            xaxis=dict(title='W (mm)', range=[0, CONT_W],
                       showbackground=True, backgroundcolor='rgba(241,245,249,0.6)',
                       gridcolor='rgba(148,163,184,0.4)', tickfont=dict(size=9)),
            yaxis=dict(title='L (mm)', range=[0, CONT_L],
                       showbackground=True, backgroundcolor='rgba(226,232,240,0.4)',
                       gridcolor='rgba(148,163,184,0.4)', tickfont=dict(size=9)),
            zaxis=dict(title='H (mm)', range=[0, CONT_H],
                       showbackground=True, backgroundcolor='rgba(248,250,252,0.6)',
                       gridcolor='rgba(148,163,184,0.4)', tickfont=dict(size=9)),
            aspectmode='data',
            camera=dict(eye=dict(x=1.5, y=-1.9, z=1.1)),
        ),
        margin=dict(l=0, r=0, t=46, b=0),
        height=560,
        legend=dict(x=0.01, y=0.98, bgcolor='rgba(255,255,255,0.9)',
                    bordercolor='#E2E8F0', borderwidth=1,
                    font=dict(size=11, family='Noto Sans KR')),
        paper_bgcolor='#FAFAFA',
    )
    return fig

# ══════════════════════════════════════════════════════
# 6. Streamlit UI
# ══════════════════════════════════════════════════════

st.set_page_config(page_title="타이어 적재 비교", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"], .stApp {
    font-family: 'Noto Sans KR', -apple-system, sans-serif;
    background: #F8FAFC;
}
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 2rem 2.5rem 3rem; max-width: 1440px; }

/* ── 헤더 */
.app-header {
    display: flex; align-items: baseline; gap: 1.5rem;
    padding-bottom: 1.2rem;
    border-bottom: 2px solid #0F172A;
    margin-bottom: 1.75rem;
}
.app-title { font-size: 1.3rem; font-weight: 700; color: #0F172A; letter-spacing: -0.02em; }
.app-meta  { font-size: 0.76rem; color: #64748B; font-family: 'IBM Plex Mono', monospace; }

/* ── 배지 */
.od-badge {
    display: inline-flex; align-items: center; gap: 0.9rem;
    background: #0F172A; color: #F8FAFC;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem; font-weight: 600;
    padding: 0.45rem 1rem; border-radius: 6px;
    letter-spacing: 0.02em;
}
.od-sep { opacity: 0.3; }

/* ── 메트릭 카드 */
.metrics-row {
    display: flex; gap: 0.85rem; margin-bottom: 1.75rem; flex-wrap: wrap;
}
.metric-card {
    flex: 1; min-width: 130px;
    background: #FFFFFF;
    border: 1px solid #E2E8F0; border-radius: 10px;
    padding: 1.1rem 1.3rem;
    border-top: 3px solid #CBD5E1;
}
.metric-card.blue   { border-top-color: #2563EB; }
.metric-card.amber  { border-top-color: #D97706; }
.metric-card.green  { border-top-color: #16A34A; }
.metric-card.red    { border-top-color: #DC2626; }
.metric-label {
    font-size: 0.65rem; font-weight: 600; color: #94A3B8;
    text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 0.45rem;
}
.metric-value {
    font-size: 1.85rem; font-weight: 700; color: #0F172A;
    font-family: 'IBM Plex Mono', monospace; line-height: 1.05;
}
.metric-sub {
    font-size: 0.72rem; color: #64748B;
    margin-top: 0.3rem; font-family: 'IBM Plex Mono', monospace;
}
.metric-sub.good { color: #16A34A; font-weight: 600; }
.metric-sub.bad  { color: #DC2626; font-weight: 600; }

/* ── 시각화 구분선 */
.viz-header {
    display: flex; align-items: center; gap: 0.6rem;
    margin-bottom: 0.4rem;
}
.viz-dot {
    width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0;
}
.viz-title {
    font-size: 0.82rem; font-weight: 700; color: #0F172A;
    text-transform: uppercase; letter-spacing: 0.05em;
}
.viz-sub {
    font-size: 0.72rem; color: #64748B;
    font-family: 'IBM Plex Mono', monospace;
    margin-bottom: 0.6rem;
}

/* ── 라디오 / 토글 */
.stRadio label { font-size: 0.82rem !important; font-weight: 500 !important; }
.stButton > button {
    background: #0F172A; color: #F8FAFC;
    border: none; border-radius: 6px;
    font-size: 0.82rem; font-weight: 500;
    padding: 0.45rem 1.2rem;
}
.stButton > button:hover { background: #1E293B; }
</style>
""", unsafe_allow_html=True)

# ── 입력 ──────────────────────────────────────────────
try:
    capacity_df, capacity_source = load_capacity_table()
except Exception as exc:
    st.error(f"적재 기준표를 읽을 수 없습니다: {exc}")
    st.stop()

if capacity_df.empty:
    st.error("적재 기준표 엑셀을 찾을 수 없습니다. 앱 폴더 또는 KakaoTalk Downloads 폴더에 파일을 두세요.")
    st.stop()

sizes = capacity_df["타이어규격"].tolist()
default_index = sizes.index("205/55R16") if "205/55R16" in sizes else 0
container_names = list(CONTAINER_SPECS)
selected_container = st.session_state.get("container_type", DEFAULT_CONTAINER)
if selected_container not in CONTAINER_SPECS:
    selected_container = DEFAULT_CONTAINER
if selected_container == "직접 입력":
    container_spec = {
        "w": st.session_state.get("custom_cont_w", CONTAINER_SPECS["직접 입력"]["w"]),
        "l": st.session_state.get("custom_cont_l", CONTAINER_SPECS["직접 입력"]["l"]),
        "h": st.session_state.get("custom_cont_h", CONTAINER_SPECS["직접 입력"]["h"]),
        "source": "계산 기준",
    }
else:
    container_spec = CONTAINER_SPECS[selected_container]
CONT_W, CONT_L, CONT_H = container_spec["w"], container_spec["l"], container_spec["h"]

# ── 헤더 ──────────────────────────────────────────────
st.markdown(f"""
<div class="app-header">
  <span class="app-title">타이어 적재 비교 계산기</span>
  <span class="app-meta">{selected_container} · W {CONT_W} × L {CONT_L} × H {CONT_H} mm</span>
</div>
""", unsafe_allow_html=True)

col_a, col_b, col_c, col_d, col_pad = st.columns([1.7, 1.4, 1.8, 1.4, 3.2])
with col_a:
    raw = st.selectbox(
        "타이어 규격",
        options=sizes,
        index=default_index,
        label_visibility="visible",
    )
with col_b:
    qty = st.number_input("타이어 개수 (개)",
                          min_value=1, max_value=500000,
                          value=500, step=1,
                          label_visibility="visible")
with col_c:
    selected_container = st.selectbox(
        "컨테이너/차량 규격",
        options=container_names,
        index=container_names.index(selected_container),
        key="container_type",
        label_visibility="visible",
    )
    if selected_container == "직접 입력":
        container_spec = {
            "w": st.session_state.get("custom_cont_w", CONTAINER_SPECS["직접 입력"]["w"]),
            "l": st.session_state.get("custom_cont_l", CONTAINER_SPECS["직접 입력"]["l"]),
            "h": st.session_state.get("custom_cont_h", CONTAINER_SPECS["직접 입력"]["h"]),
            "source": "계산 기준",
        }
    else:
        container_spec = CONTAINER_SPECS[selected_container]
    CONT_W, CONT_L, CONT_H = container_spec["w"], container_spec["l"], container_spec["h"]
with col_d:
    st.text_input("기준표", value=f"{len(capacity_df):,}개 규격", disabled=True)

if selected_container == "직접 입력":
    dim_c1, dim_c2, dim_c3, dim_pad = st.columns([1.3, 1.3, 1.3, 6.1])
    with dim_c1:
        CONT_W = st.number_input("폭 W (mm)", min_value=500, max_value=4000,
                                 value=int(CONT_W), step=10, key="custom_cont_w")
    with dim_c2:
        CONT_L = st.number_input("길이 L (mm)", min_value=1000, max_value=20000,
                                 value=int(CONT_L), step=10, key="custom_cont_l")
    with dim_c3:
        CONT_H = st.number_input("높이 H (mm)", min_value=500, max_value=4000,
                                 value=int(CONT_H), step=10, key="custom_cont_h")

if not raw.strip():
    st.stop()
od, h = parse_tire_size(raw.strip())
if od is None:
    st.error("규격을 인식할 수 없습니다. 예: 205/55R16")
    st.stop()

st.markdown(f"""
<div style="margin: -0.4rem 0 1.6rem;">
  <span class="od-badge">
    OD <span class="od-sep">|</span> {od} mm
    &nbsp;&nbsp; 폭 <span class="od-sep">|</span> {h} mm
    &nbsp;&nbsp; 규격 <span class="od-sep">|</span> {selected_container}
    &nbsp;&nbsp; W×L×H <span class="od-sep">|</span> {CONT_W}×{CONT_L}×{CONT_H}
    &nbsp;&nbsp; 수량 <span class="od-sep">|</span> {qty:,}개
  </span>
</div>
""", unsafe_allow_html=True)

# ── 계산 ──────────────────────────────────────────────
flat  = calc_flat(od, h)
honey = calc_honeycomb(od, h)
capacity_row = capacity_df.loc[capacity_df["타이어규격"] == raw.strip().upper()].iloc[0]
calc_flat_total = flat["total"]
calc_honey_total = honey["total"]
uses_excel_capacity = selected_container == DEFAULT_CONTAINER
if uses_excel_capacity:
    flat["total"] = int(capacity_row["일반평치"])
    honey["total"] = int(capacity_row["벌집평치"])

if flat['total'] == 0 or honey['total'] == 0:
    st.error("타이어가 너무 커서 컨테이너에 적재할 수 없습니다.")
    st.stop()

if uses_excel_capacity and (flat["total"] != calc_flat_total or honey["total"] != calc_honey_total):
    st.info(
        f"적재수량은 엑셀 기준표를 적용했습니다. "
        f"시각화 배치 계산값: 평치 {calc_flat_total:,}개, 벌집 {calc_honey_total:,}개"
    )
elif not uses_excel_capacity:
    st.info(
        f"{selected_container}는 엑셀 기준표가 없어 선택한 컨테이너 치수로 계산한 적재수량을 적용했습니다."
    )

flat_n_cont  = math.ceil(qty / flat['total'])
honey_n_cont = math.ceil(qty / honey['total'])

flat_last    = qty % flat['total']  or flat['total']   # 마지막 컨테이너 적재 수
honey_last   = qty % honey['total'] or honey['total']

flat_util    = flat_last  / flat['total']  * 100
honey_util   = honey_last / honey['total'] * 100

cont_saved   = flat_n_cont - honey_n_cont
best_label   = "벌집" if honey_n_cont < flat_n_cont else ("평치" if flat_n_cont < honey_n_cont else "동일")
best_color   = "amber" if honey_n_cont < flat_n_cont else ("blue" if flat_n_cont < honey_n_cont else "green")

if cont_saved > 0:
    saved_txt = f"벌집이 {cont_saved}개 절약"
    saved_cls = "good"
elif cont_saved < 0:
    saved_txt = f"평치가 {abs(cont_saved)}개 절약"
    saved_cls = "good"
else:
    saved_txt  = "두 방식 동일"
    saved_cls  = ""

# ── 메트릭 카드 ───────────────────────────────────────
st.markdown(f"""
<div class="metrics-row">
  <div class="metric-card blue">
    <div class="metric-label">평치 — 컨테이너 수</div>
    <div class="metric-value">{flat_n_cont}</div>
    <div class="metric-sub">1개당 {flat['total']:,}개 적재</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">벌집 — 컨테이너 수</div>
    <div class="metric-value">{honey_n_cont}</div>
    <div class="metric-sub">1개당 {honey['total']:,}개 적재</div>
  </div>
  <div class="metric-card {best_color}">
    <div class="metric-label">효율 우위</div>
    <div class="metric-value" style="font-size:1.5rem">{best_label}</div>
    <div class="metric-sub {saved_cls}">{saved_txt}</div>
  </div>
  <div class="metric-card blue">
    <div class="metric-label">평치 마지막 적재율</div>
    <div class="metric-value">{flat_util:.1f}%</div>
    <div class="metric-sub">{flat_last:,} / {flat['total']:,}개</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">벌집 마지막 적재율</div>
    <div class="metric-value">{honey_util:.1f}%</div>
    <div class="metric-sub">{honey_last:,} / {honey['total']:,}개</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── 3D 옵션 ───────────────────────────────────────────
opt_c1, opt_c2, opt_c3 = st.columns([2.2, 2.5, 4])
with opt_c1:
    show_empty = st.checkbox("빈 자리 표시", value=False,
                             help="마지막 컨테이너의 남은 공간을 반투명하게 표시합니다.")
with opt_c2:
    view_mode = st.radio(
        "표시 기준",
        ["마지막 컨테이너", "첫 컨테이너 (만재)"],
        horizontal=True,
        help="마지막 = 실제 채워지는 최종 컨테이너 / 첫 = 가득 채워진 기준 컨테이너"
    )

# ── 표시 개수 결정 ────────────────────────────────────
if view_mode == "첫 컨테이너 (만재)":
    flat_disp  = min(qty, flat['total'])
    honey_disp = min(qty, honey['total'])
    view_label_f = f"컨테이너 1/{flat_n_cont}"
    view_label_h = f"컨테이너 1/{honey_n_cont}"
else:
    flat_disp  = flat_last
    honey_disp = honey_last
    view_label_f = f"컨테이너 {flat_n_cont}/{flat_n_cont}"
    view_label_h = f"컨테이너 {honey_n_cont}/{honey_n_cont}"

# ── 3D 시각화 나란히 ──────────────────────────────────
flat_items  = flat_positions(od, h, flat)
honey_items = honeycomb_positions(od, h, honey)

col_f, col_h = st.columns(2)

with col_f:
    st.markdown(f"""
<div class="viz-header">
  <div class="viz-dot" style="background:#1D4ED8"></div>
  <span class="viz-title">일반 적재 (평치)</span>
</div>
<div class="viz-sub">
  {view_label_f} &nbsp;·&nbsp; {flat_disp:,} / {flat['total']:,}개
  &nbsp;·&nbsp; 적재율 {flat_disp/flat['total']*100:.1f}%
  &nbsp;·&nbsp; 전체 {flat_n_cont}컨테이너 필요
</div>
""", unsafe_allow_html=True)
    with st.spinner("평치 3D 생성 중…"):
        fig_flat = build_3d(
            od, h, flat_items, flat_disp,
            '#1D4ED8', '#CBD5E1',
            '평치 적재', flat['total'], show_empty
        )
    st.plotly_chart(fig_flat, use_container_width=True)

with col_h:
    st.markdown(f"""
<div class="viz-header">
  <div class="viz-dot" style="background:#D97706"></div>
  <span class="viz-title">벌집 적재 (육각)</span>
</div>
<div class="viz-sub">
  {view_label_h} &nbsp;·&nbsp; {honey_disp:,} / {honey['total']:,}개
  &nbsp;·&nbsp; 적재율 {honey_disp/honey['total']*100:.1f}%
  &nbsp;·&nbsp; 전체 {honey_n_cont}컨테이너 필요
</div>
""", unsafe_allow_html=True)
    with st.spinner("벌집 3D 생성 중…"):
        fig_honey = build_3d(
            od, h, honey_items, honey_disp,
            '#D97706', '#FDE68A',
            '벌집 적재', honey['total'], show_empty
        )
    st.plotly_chart(fig_honey, use_container_width=True)

# ── 컨테이너별 상세 ───────────────────────────────────
st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
with st.expander("컨테이너별 적재 상세"):
    flat_rows, honey_rows = [], []
    for i in range(1, flat_n_cont + 1):
        cnt = flat['total'] if i < flat_n_cont else flat_last
        flat_rows.append({
            "컨테이너": i,
            "적재 수": cnt,
            "적재율": f"{cnt/flat['total']*100:.1f}%",
            "상태": "만재" if cnt == flat['total'] else "부분",
        })
    for i in range(1, honey_n_cont + 1):
        cnt = honey['total'] if i < honey_n_cont else honey_last
        honey_rows.append({
            "컨테이너": i,
            "적재 수": cnt,
            "적재율": f"{cnt/honey['total']*100:.1f}%",
            "상태": "만재" if cnt == honey['total'] else "부분",
        })

    dc1, dc2 = st.columns(2)
    with dc1:
        st.caption(f"평치 — {flat_n_cont}컨테이너")
        st.dataframe(pd.DataFrame(flat_rows), use_container_width=True, hide_index=True)
    with dc2:
        st.caption(f"벌집 — {honey_n_cont}컨테이너")
        st.dataframe(pd.DataFrame(honey_rows), use_container_width=True, hide_index=True)
