"""
QA Product Dashboard · Telkomsel — Streamlit Python version
============================================================
Source: data/*.xlsx (cleaned files)
Run    : streamlit run app.py
Author : Generated for Panca

Features:
- Auto-load all XLSX files from data/ folder on startup
- Executive Summary with bullets, glossary, KPIs, action items
- Per-file views: KIP April, Cohort LIS, QA Product
- Upload tambahan with smart header detection (sama dengan HTML)
- Mobile-responsive (Streamlit's native responsive layout)
"""

from pathlib import Path
import re
import io

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# =============================================================
# CONFIG
# =============================================================
st.set_page_config(
    page_title="QA Product Dashboard · Telkomsel",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = Path(__file__).parent / "data"

PALETTE = ['#ED1C24', '#FF8A3D', '#FFC93C', '#5DD39E', '#3B82F6',
           '#A78BFA', '#EC4899', '#22D3EE', '#F472B6', '#84CC16']
ACCENT = '#ED1C24'
C_GREEN = '#5DD39E'
C_ORANGE = '#FF8A3D'
C_BLUE = '#3B82F6'

# Plotly layout defaults matching dark theme
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(family='Inter, system-ui, sans-serif', color='#a3a3a3', size=11),
    margin=dict(l=10, r=10, t=40, b=10),
    xaxis=dict(gridcolor='#1f1f1f', linecolor='#2a2a2a', zeroline=False, color='#737373'),
    yaxis=dict(gridcolor='#1f1f1f', linecolor='#2a2a2a', zeroline=False, color='#737373'),
    hoverlabel=dict(bgcolor='#1c1c1c', bordercolor='#3a3a3a', font_family='Inter, sans-serif'),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(size=10)),
    colorway=PALETTE,
)


# =============================================================
# CSS THEME — make Streamlit look like our dashboard
# =============================================================
def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Inter+Tight:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

    .stApp {
        background:
            radial-gradient(ellipse 80% 50% at 50% -10%, rgba(237,28,36,.06), transparent 70%),
            #0a0a0a !important;
        color: #f5f5f5;
        font-family: 'Inter Tight', sans-serif;
    }
    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #141414;
        border-right: 1px solid #2a2a2a;
    }
    [data-testid="stSidebar"] .stRadio label, [data-testid="stSidebar"] p {
        color: #d4d4d4 !important;
        font-size: 13px;
    }
    /* Main content max width */
    .main .block-container { padding-top: 1.5rem; max-width: 1400px; }
    /* Headings */
    h1, h2, h3, h4 { font-family: 'Instrument Serif', serif !important; font-style: italic; font-weight: 400 !important; letter-spacing: -.01em !important; }
    h1 { font-size: 2.4rem !important; }
    h2 { font-size: 1.6rem !important; margin-top: 1.6rem !important; }
    h3 { font-size: 1.25rem !important; }
    /* Metric (st.metric) styling */
    [data-testid="stMetric"] {
        background: #141414;
        border: 1px solid #2a2a2a;
        border-radius: 12px;
        padding: 14px 16px;
        position: relative;
        overflow: hidden;
    }
    [data-testid="stMetric"]::before {
        content: ''; position: absolute; left:0; top:0; height:100%; width:3px;
        background: var(--accent, #ED1C24);
    }
    [data-testid="stMetricLabel"] {
        font-size: 11px !important;
        text-transform: uppercase;
        letter-spacing: .08em;
        color: #737373 !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 24px !important;
        font-weight: 600 !important;
        color: #f5f5f5 !important;
    }
    [data-testid="stMetricDelta"] {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 12px !important;
    }
    /* Style for our custom HTML cards */
    .bullet-card {
        background: #141414;
        border: 1px solid #2a2a2a;
        border-radius: 14px;
        padding: 18px;
        position: relative;
        overflow: hidden;
        height: 100%;
    }
    .bullet-card::before {
        content: ''; position: absolute; left:0; top:0; width: 100%; height: 3px;
    }
    .bullet-card.perf::before { background: linear-gradient(90deg, #5DD39E, transparent); }
    .bullet-card.risk::before { background: linear-gradient(90deg, #ED1C24, transparent); }
    .bullet-card.cx::before   { background: linear-gradient(90deg, #3B82F6, transparent); }
    .bullet-card .tag {
        display: inline-block; font-family: 'JetBrains Mono', monospace;
        font-size: 11px; font-weight: 700; letter-spacing: .1em;
        padding: 3px 8px; border-radius: 5px; margin-bottom: 10px;
    }
    .bullet-card.perf .tag { background: rgba(93,211,158,.14); color: #5DD39E; }
    .bullet-card.risk .tag { background: rgba(237,28,36,.14); color: #ED1C24; }
    .bullet-card.cx .tag   { background: rgba(59,130,246,.14); color: #3B82F6; }
    .bullet-card h4 {
        font-family: 'Instrument Serif', serif !important;
        font-style: italic; font-weight: 400;
        font-size: 18px !important; margin: 0 0 12px !important; color: #f5f5f5;
    }
    .bullet-card ul { margin: 0; padding: 0; list-style: none; }
    .bullet-card li {
        position: relative; padding: 8px 0 8px 20px;
        font-size: 13px; color: #a3a3a3; line-height: 1.55;
        border-bottom: 1px dashed #2a2a2a;
    }
    .bullet-card li:last-child { border-bottom: none; }
    .bullet-card li::before {
        content: ''; position: absolute; left: 4px; top: 15px;
        width: 6px; height: 6px; border-radius: 50%;
    }
    .bullet-card.perf li::before { background: #5DD39E; }
    .bullet-card.risk li::before { background: #ED1C24; }
    .bullet-card.cx li::before { background: #3B82F6; }
    .bullet-card li b { color: #f5f5f5; font-weight: 600; }
    .bullet-card li code {
        font-family: 'JetBrains Mono', monospace; font-size: 11.5px;
        background: #1c1c1c; padding: 1px 6px; border-radius: 4px;
        color: #f5f5f5; border: 1px solid #2a2a2a;
    }
    /* Glossary cards */
    .gloss-card {
        background: #141414;
        border: 1px solid #2a2a2a;
        border-radius: 11px;
        padding: 14px 16px;
        height: 100%;
    }
    .gloss-term {
        display: flex; justify-content: space-between; align-items: center;
        gap: 10px; margin-bottom: 6px;
    }
    .gloss-term b { font-size: 13.5px; color: #f5f5f5; font-weight: 600; }
    .gloss-val {
        font-family: 'JetBrains Mono', monospace; font-size: 13px;
        font-weight: 600; color: #ED1C24; white-space: nowrap;
    }
    .gloss-desc { font-size: 12.5px; color: #a3a3a3; line-height: 1.55; margin: 0; }
    .gloss-formula {
        display: inline-block; font-family: 'JetBrains Mono', monospace; font-size: 11px;
        background: #1c1c1c; padding: 3px 8px; border-radius: 5px;
        color: #a3a3a3; border: 1px solid #2a2a2a; margin-top: 6px;
    }
    /* Action cards */
    .action-card {
        background: #141414; border: 1px solid #2a2a2a;
        border-radius: 13px; padding: 16px; height: 100%;
        display: flex; flex-direction: column; gap: 8px;
    }
    .action-card .num {
        font-family: 'Instrument Serif', serif; font-style: italic;
        font-size: 34px; line-height: 1; color: #ED1C24;
    }
    .action-card .title { font-weight: 600; font-size: 14px; color: #f5f5f5; }
    .action-card .body { font-size: 12.5px; color: #a3a3a3; line-height: 1.55; }
    .action-card .tag {
        display: inline-block; font-size: 10px; font-weight: 700;
        letter-spacing: .1em; text-transform: uppercase;
        padding: 3px 8px; border-radius: 5px; align-self: flex-start;
    }
    .action-card .tag.urgent { background: rgba(237,28,36,.14); color: #ED1C24; }
    .action-card .tag.eff { background: rgba(93,211,158,.14); color: #5DD39E; }
    .action-card .tag.gov { background: rgba(255,138,61,.14); color: #FF8A3D; }
    /* Hero banner */
    .hero-banner {
        background: linear-gradient(135deg, rgba(237,28,36,.06) 0%, rgba(59,130,246,.04) 100%);
        border: 1px solid #2a2a2a; border-radius: 16px; padding: 22px;
        margin-bottom: 22px;
    }
    .hero-banner .label {
        font-size: 11px; color: #ED1C24; text-transform: uppercase;
        letter-spacing: .14em; font-weight: 700;
    }
    .hero-banner h2 {
        font-family: 'Instrument Serif', serif !important; font-style: italic; font-weight: 400;
        font-size: 1.5rem !important; line-height: 1.3; margin: 8px 0 0 !important;
        color: #f5f5f5;
    }
    .hero-banner h2 .h { color: #ED1C24; font-style: italic; }
    .hero-banner h2 .g { color: #5DD39E; font-style: italic; }
    /* Latency cells */
    .lat-cell {
        background: #141414; border: 1px solid #2a2a2a;
        border-radius: 11px; padding: 12px 14px;
        position: relative;
    }
    .lat-cell::after {
        content: ''; position: absolute; right: 14px; top: 14px;
        width: 8px; height: 8px; border-radius: 50%;
    }
    .lat-cell.ok::after { background: #5DD39E; }
    .lat-cell.warn::after { background: #FF8A3D; }
    .lat-cell.bad::after { background: #ED1C24; animation: pulse 1.6s ease-in-out infinite; }
    @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }
    .lat-cell .prod { font-size: 12px; color: #a3a3a3; }
    .lat-cell .val {
        font-family: 'JetBrains Mono', monospace; font-size: 20px;
        font-weight: 600; color: #f5f5f5;
    }
    .lat-cell .meta {
        font-size: 10px; color: #737373; font-family: 'JetBrains Mono', monospace;
    }
    /* Section number prefix */
    .sec-num {
        font-family: 'JetBrains Mono', monospace; font-size: 13px;
        color: #737373; margin-right: 10px; font-weight: 500;
        font-style: normal; vertical-align: middle;
    }
    /* Hide Streamlit chrome */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    [data-testid="stHeader"] { background: transparent; }
    </style>
    """, unsafe_allow_html=True)


# =============================================================
# HELPERS
# =============================================================
def fmt_num(n):
    if n is None or pd.isna(n): return "–"
    a = abs(n)
    if a >= 1e12: return f"{n/1e12:.2f} T"
    if a >= 1e9:  return f"{n/1e9:.2f} B"
    if a >= 1e6:  return f"{n/1e6:.2f} M"
    if a >= 1e3:  return f"{n/1e3:.1f} K"
    return f"{round(n):,}".replace(",", ".")

def fmt_idr(n):
    return "Rp " + fmt_num(n)

def section(num, text):
    st.markdown(f'<h2><span class="sec-num">{num}</span><i>{text}</i></h2>', unsafe_allow_html=True)


# =============================================================
# DATA LOADERS (cached)
# =============================================================
@st.cache_data(show_spinner=False)
def load_kip():
    df = pd.read_excel(DATA_DIR / "KIP_April_2026_CLEANED.xlsx")
    df['Date'] = pd.to_datetime(df['Date'])
    return df

@st.cache_data(show_spinner=False)
def load_cohort():
    df = pd.read_excel(DATA_DIR / "cohort_oct25_mar26_lis_CLEANED.xlsx")
    df['report_month'] = df['report_month'].astype(str)
    return df

@st.cache_data(show_spinner=False)
def load_qa():
    return pd.read_excel(DATA_DIR / "Dashboard_QA_Product_Q1_2026_CLEANED.xlsx", sheet_name=None)


# =============================================================
# SMART HEADER DETECTION (for uploaded files)
# =============================================================
def score_header_row(row, prev_row=None):
    if not row or len(row) == 0: return -np.inf
    total = len(row)
    filled = sum(1 for v in row if pd.notna(v) and v != '')
    if filled == 0: return -np.inf
    fill_ratio = filled / total
    strings = sum(1 for v in row if isinstance(v, str) and v.strip())
    nums = sum(1 for v in row if isinstance(v, (int, float)) and not isinstance(v, bool) and pd.notna(v))
    dates = sum(1 for v in row if hasattr(v, 'strftime'))
    string_ratio = strings / filled
    uniq = len(set(str(v).strip() for v in row if pd.notna(v) and v != ''))
    uniq_ratio = uniq / filled
    score = fill_ratio*10 + string_ratio*6 + uniq_ratio*8 - (nums/total)*8 - (dates/total)*4
    if prev_row is not None:
        prev_filled = sum(1 for v in prev_row if pd.notna(v) and v != '')
        if prev_filled/total < 0.3: score += 2.5
    if filled < 3: score -= 5
    return score

def detect_header_row(rows):
    if not rows: return 0
    best, best_score = 0, -np.inf
    for i in range(min(10, len(rows))):
        s = score_header_row(rows[i], rows[i-1] if i > 0 else None)
        if s > best_score:
            best_score, best = s, i
    return best

def read_sheet_smart(file_bytes, sheet_name, header_row):
    """Read a sheet using a specific header row index."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df = df.dropna(how='all').reset_index(drop=True)
    # De-dupe column names
    seen = {}
    new_cols = []
    for c in df.columns:
        name = str(c) if pd.notna(c) and str(c).strip() else f"col_{len(new_cols)+1}"
        name = name.strip()
        if name in seen:
            seen[name] += 1
            new_cols.append(f"{name} ({seen[name]})")
        else:
            seen[name] = 1
            new_cols.append(name)
    df.columns = new_cols
    return df

def classify_column(series):
    s = series.dropna()
    if len(s) == 0: return 'unknown'
    if pd.api.types.is_datetime64_any_dtype(series): return 'date'
    if pd.api.types.is_numeric_dtype(series): return 'number'
    try:
        pd.to_datetime(s.head(50), errors='raise')
        return 'date'
    except Exception:
        pass
    uniq = s.nunique()
    return 'category' if (uniq < len(s)*0.5 and uniq < 60) else 'text'

def assess_quality(df):
    types = [classify_column(df[c]) for c in df.columns]
    has_date = 'date' in types
    has_num = 'number' in types
    has_cat = 'category' in types
    if len(df) < 2: return 'poor'
    if has_date and has_num: return 'great'
    if (has_num and has_cat) or sum(t in ('date','number','category') for t in types) >= 3:
        return 'great'
    if has_cat or has_num: return 'ok'
    return 'poor'


# =============================================================
# EXECUTIVE SUMMARY
# =============================================================
def render_executive(kip, coh, qa_sheets):
    qa_raw = qa_sheets['Raw Data']
    qa_monthly = qa_sheets['Monthly Trend']
    qa_regional = qa_sheets['Regional Analysis']
    qa_top5 = qa_sheets['Top 5 Issues']
    qa_ceto = qa_sheets['Ceto Latency']

    total_lis = int(coh['lis'].sum())
    total_with_ticket = int(coh['lis_with_ticket'].sum())
    total_revenue = float(coh['rev_bill'].sum())
    ticket_rate = total_with_ticket / total_lis * 100
    rev_per_lis = total_revenue / total_lis

    # ---- HERO ----
    st.markdown(f"""
    <div class="hero-banner">
      <div class="label">★ Executive Summary · 6 Bulan terakhir</div>
      <h2>{fmt_num(total_lis)} LIS dengan ticket rate <span class="h">{ticket_rate:.2f}%</span>. Tren komplain lifestyle products
      <span class="g">menurun -37% s.d. -54%</span>, tapi <span class="h">FTTR Jakarta naik +131%</span> &mdash; perlu deep-dive sebelum ekspansi nasional.</h2>
    </div>
    """, unsafe_allow_html=True)

    # ---- KPI strip ----
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total LIS", fmt_num(total_lis), "pelanggan aktif · 6 bln")
    k2.metric("Ticket Rate", f"{ticket_rate:.2f}%", f"{fmt_num(total_with_ticket)} tickets")
    k3.metric("Revenue 6M", fmt_idr(total_revenue), f"Rp {fmt_num(rev_per_lis)} / LIS")
    k4.metric("Q1 Records", fmt_num(len(qa_raw)), f"{qa_raw['product'].nunique()} products")
    k5.metric("Top Pain Point", "K26", "subscription gagal")

    # ---- BULLETS ----
    section("01", "Ringkasan Eksekutif")
    cc_pct = (kip['unit_type'] == 'callcenter').mean() * 100
    ws_pct = (kip['unit_type'] == 'webservice').mean() * 100
    gp_pct = (kip['unit_type'] == 'grapari').mean() * 100
    reg_counts = qa_raw['regional'].value_counts()
    top3_conc = reg_counts.head(3).sum() / reg_counts.sum() * 100

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(f"""
        <div class="bullet-card perf">
          <span class="tag">A · WIN</span>
          <h4>Performance Highlights</h4>
          <ul>
            <li><b>Lifestyle products membaik konsisten</b> — komplain Q1→April turun di 7 dari 9 produk</li>
            <li><b>CHATGPT Go</b>: 897 → 564 cases (<code>-37%</code>) — fix campaign dan UX terlihat berdampak</li>
            <li><b>Travel Assistant</b>: 436 → 200 (<code>-54%</code>) — penurunan tertinggi di portfolio</li>
            <li><b>Jaringan Prioritas</b>: 617 → 295 (<code>-52%</code>), VNSP <code>-51%</code></li>
            <li><b>Revenue 6 bulan</b>: Rp {fmt_num(total_revenue)} dari kohort, Rp {fmt_num(rev_per_lis)} per LIS</li>
            <li><b>Speed 100Mbps mendominasi</b> — 97,5% dari base LIS, indikasi product-market fit kuat</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div class="bullet-card risk">
          <span class="tag">B · RISK</span>
          <h4>Yang Perlu Diwaspadai</h4>
          <ul>
            <li><b>FTTR melonjak</b>: 140 → 324 cases (<code>+131%</code>) — satu-satunya produk dengan tren naik signifikan</li>
            <li><b>Konsentrasi Jakarta-Banten</b>: 90% komplain FTTR (939 dari 1.040) — masalah onboarding di market launch</li>
            <li><b>K26 "Subscription gagal"</b>: 537 cases lintas data (430 ChatGPT + 107 KIP April)</li>
            <li><b>ProtekSi Kecil latency 108s</b> — 2× rata-rata produk lain</li>
            <li><b>Bill Shock</b>: 8 cases/bulan dari multi-OTT subscription stack (hingga Rp 978K)</li>
            <li><b>Beban geografis</b>: Jateng+Jabar+HQ = {top3_conc:.0f}% dari semua komplain</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)
    with col_c:
        st.markdown(f"""
        <div class="bullet-card cx">
          <span class="tag">C · CX</span>
          <h4>Customer Experience</h4>
          <ul>
            <li><b>Ticket Rate {ticket_rate:.2f}%</b> — dari 100 pelanggan, ~7 sampai membuka tiket. <b>92,85% silent</b></li>
            <li><b>Channel mix</b>: Call Center {cc_pct:.0f}%, Webservice {ws_pct:.0f}%, GraPARI {gp_pct:.0f}%</li>
            <li><b>Top 3 issue codes</b>: K26 (subscription), K57 (SMS spam), K54 (modify offer/VAS)</li>
            <li><b>Latency aktivasi</b>: 7 dari 8 produk sehat (44-66s), kecuali ProtekSi (108s) dan V-NSP (70s warn)</li>
            <li><b>Konsentrasi region</b>: Jateng 35%, Jabar 28%, HQ 14%</li>
            <li><b>Activation transactions</b>: 251 trans/5 hari di sample Ceto, status mostly OK</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    # ---- GLOSSARY ----
    section("02", "Apa Arti Parameter Ini?")
    gloss = [
        ("Ticket Rate", f"{ticket_rate:.2f}%",
         "Persentase pelanggan aktif (LIS) yang sampai membuka tiket keluhan dalam periode pengukuran. <b>7,15%</b> artinya dari 100 pelanggan, sekitar 7 yang komplain — sisanya silent.",
         "LIS with ticket ÷ Total LIS × 100%"),
        ("LIS (Live Subscriber)", fmt_num(total_lis),
         "Jumlah pelanggan dengan layanan aktif di periode pengukuran. Base untuk semua perhitungan rasio.",
         "COUNT DISTINCT subscribers per month"),
        ("Revenue per LIS", f"Rp {fmt_num(rev_per_lis)}",
         "Rata-rata pendapatan dari billing per pelanggan dalam 6 bulan kohort. Indikator monetisasi & ARPU konversi.",
         "Total Revenue ÷ Total LIS"),
        ("Decline % (Jan→Apr)", "-37% s.d. -54%",
         "Perubahan jumlah komplain antara Januari dan April 2026. <b>Negatif (hijau) = membaik</b>, positif (merah) = memburuk.",
         "(Apr − Jan) ÷ Jan × 100%"),
        ("K-Code", "K21·K26·K54·K57",
         "Kode klasifikasi keluhan internal. <b>K21</b>=produk umum, <b>K26</b>=subscription, <b>K41</b>=billing, <b>K54</b>=modify offer, <b>K57</b>=SMS spam.",
         "KCS coding system (internal)"),
        ("Concentration %", f"{top3_conc:.0f}%",
         f"Persentase beban komplain dari top-3 region. <b>{top3_conc:.0f}%</b> berarti ¾ beban support dari 3 region — sinyal untuk re-allocate resource.",
         "Σ Top-3 region tickets ÷ Total tickets"),
        ("Activation Latency (s)", "44 – 108s",
         "Waktu rata-rata transaksi aktivasi produk. <b>&lt;60s = OK 🟢</b>, <b>60-70s = Warn 🟠</b>, <b>&gt;70s = Lambat 🔴</b>.",
         "AVG(end_time − start_time) per product"),
        ("Bill Shock", "8 / bulan",
         "Tagihan tiba-tiba membengkak akibat <i>subscription stacking</i> (ChatGPT + Adobe + Netflix di nomor sama, semua auto-renew). Hingga Rp 978K/case.",
         "K41 type_3 cases / period"),
        ("Channel Mix", f"{cc_pct:.0f}% CC",
         f"Distribusi kanal komplain. <b>Call Center dominan ({cc_pct:.0f}%)</b> = kanal mahal. Migrasi ke chatbot bisa potong cost CC ±30-40%.",
         "Tickets per channel ÷ Total tickets"),
    ]
    # 3 columns, 3 cards each row
    for i in range(0, len(gloss), 3):
        cols = st.columns(3)
        for j, c in enumerate(cols):
            if i+j < len(gloss):
                term, val, desc, formula = gloss[i+j]
                c.markdown(f"""
                <div class="gloss-card">
                  <div class="gloss-term"><b>{term}</b><span class="gloss-val">{val}</span></div>
                  <p class="gloss-desc">{desc}</p>
                  <span class="gloss-formula">{formula}</span>
                </div>
                """, unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ---- TREND CHART ----
    section("03", "Visual Tren Q1 → April 2026")

    months = ['2026-01', '2026-02', '2026-03', '2026-04']
    fig = go.Figure()
    for i, row in qa_monthly.iterrows():
        product = row['Produk-KIP']
        values = [row[m] for m in months]
        is_fttr = product == 'FTTR'
        fig.add_trace(go.Scatter(
            x=months, y=values, name=product, mode='lines+markers',
            line=dict(color=ACCENT if is_fttr else PALETTE[i % len(PALETTE)],
                      width=3 if is_fttr else 1.5),
            marker=dict(size=8 if is_fttr else 4),
            fill='tozeroy' if is_fttr else None,
            fillcolor='rgba(237,28,36,.15)' if is_fttr else None,
        ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Monthly Tickets per Product · FTTR highlighted",
                   font=dict(family='Instrument Serif, serif', size=16, color='#f5f5f5')),
        height=400,
    )
    st.plotly_chart(fig, use_container_width=True)

    # Two side-by-side
    cc1, cc2 = st.columns(2)
    with cc1:
        # Regional donut
        top3 = reg_counts.head(3)
        rest = reg_counts.iloc[3:].sum()
        labels = list(top3.index) + ['Lainnya']
        values = list(top3.values) + [rest]
        pcts = [v/sum(values)*100 for v in values]
        labels = [f"{l} ({p:.0f}%)" for l, p in zip(labels, pcts)]
        fig2 = go.Figure(go.Pie(labels=labels, values=values, hole=0.55,
                                marker=dict(colors=[PALETTE[0], PALETTE[1], PALETTE[2], '#3a3a3a'],
                                            line=dict(color='#141414', width=2))))
        fig2.update_layout(
            **PLOTLY_LAYOUT,
            title=dict(text=f"Konsentrasi Regional · Top 3 = {top3_conc:.0f}%",
                       font=dict(family='Instrument Serif, serif', size=16, color='#f5f5f5')),
            height=360, showlegend=True,
        )
        st.plotly_chart(fig2, use_container_width=True)

    with cc2:
        # Top pain points
        pain = qa_raw['type_3'].value_counts().head(6)
        pain_labels = [l[:40] + '…' if len(l) > 40 else l for l in pain.index]
        colors = [ACCENT] + [PALETTE[(i+1) % len(PALETTE)] for i in range(len(pain)-1)]
        fig3 = go.Figure(go.Bar(
            x=pain.values, y=pain_labels, orientation='h',
            marker=dict(color=colors), text=pain.values, textposition='outside',
        ))
        fig3.update_layout(
            **PLOTLY_LAYOUT,
            title=dict(text="Top 6 Pain Points (K-codes)",
                       font=dict(family='Instrument Serif, serif', size=16, color='#f5f5f5')),
            height=360, yaxis=dict(autorange='reversed', gridcolor='#1f1f1f', color='#a3a3a3'),
        )
        st.plotly_chart(fig3, use_container_width=True)

    # ---- LATENCY ----
    section("04", "System Health · Activation Latency")
    cols = st.columns(4)
    for i, (_, row) in enumerate(qa_ceto.iterrows()):
        lat = row['Avg. Latency (s)']
        status = 'bad' if lat > 70 else 'warn' if lat > 60 else 'ok'
        cols[i % 4].markdown(f"""
        <div class="lat-cell {status}">
          <div class="prod">{row['Product Periode 20sd24Apr']}</div>
          <div class="val">{lat:.1f}s</div>
          <div class="meta">{row['Qty (Trans)']} trans · {row['Status']}</div>
        </div>
        """, unsafe_allow_html=True)

    # ---- ACTIONS ----
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
    section("05", "Recommended Executive Actions")
    actions = [
        ('urgent', 'Urgent · Engineering', 'Fix ChatGPT Go redemption flow',
         'K26 "Pembelian berhasil namun subscription tidak didapatkan" = 430 cases di CHATGPT Go saja, plus 107 di KIP April. Engineering deep-dive ke voucher claim API.'),
        ('urgent', 'Urgent · Product', 'Audit FTTR Jakarta launch',
         'FTTR komplain naik +131% sejak Januari, 90% terkonsentrasi di Jakarta-Banten. Sebelum ekspansi nasional, resolve root cause di Jakarta market.'),
        ('eff', 'Efficiency · UX', 'Optimize ProtekSi Kecil journey',
         'Avg latency aktivasi 108 detik (2× rata-rata). High-friction journey = drop-off pembelian. Target turunkan ke <60s.'),
        ('gov', 'Governance · Billing', 'Bill Shock prevention untuk Halo multi-OTT',
         'Customer Halo dengan multiple OTT subscriptions = high-risk segment untuk billing dispute. Hingga Rp 978K/case. Implement opt-in confirmation.'),
        ('eff', 'Resource · Operations', 'Reallocate CC ke Jateng & Jabar',
         f'Top 3 region = {top3_conc:.0f}% beban support. Allocate tim CC mengikuti konsentrasi geografis, bukan distribusi merata.'),
        ('eff', 'Efficiency · Channel', 'Self-service untuk top-3 K-codes',
         f'Call Center menyumbang {cc_pct:.0f}% complaints. Top issues (K26, K57, K54) = repetitive. Chatbot/in-app FAQ berpotensi reduce 30-40% beban CC.'),
    ]
    for i in range(0, len(actions), 3):
        cols = st.columns(3)
        for j, c in enumerate(cols):
            if i+j < len(actions):
                tag, tag_label, title, body = actions[i+j]
                c.markdown(f"""
                <div class="action-card">
                  <span class="tag {tag}">{tag_label}</span>
                  <div class="num">{i+j+1:02d}</div>
                  <div class="title">{title}</div>
                  <div class="body">{body}</div>
                </div>
                """, unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)


# =============================================================
# KIP RENDERER
# =============================================================
def render_kip(df):
    st.markdown(f'<h1>KIP April 2026 <span style="color:#737373;font-size:.5em;font-family:JetBrains Mono;font-style:normal">{len(df)} rows</span></h1>', unsafe_allow_html=True)
    by_product = df['product_name'].value_counts()
    top1 = by_product.index[0]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Tickets", fmt_num(len(df)))
    k2.metric("Products", df['product_name'].nunique())
    k3.metric("Top Product", top1, f"{fmt_num(by_product.iloc[0])} tickets")
    k4.metric("Periode", df['Date'].min().strftime("%d %b"), f"s.d. {df['Date'].max().strftime('%d %b')}")

    section("01", "Distribusi Tiket")
    c1, c2 = st.columns([3, 2])
    with c1:
        fig = go.Figure(go.Bar(
            x=by_product.values, y=by_product.index, orientation='h',
            marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(by_product))]),
            text=by_product.values, textposition='outside',
        ))
        fig.update_layout(**PLOTLY_LAYOUT, title="Tickets per Product",
                          yaxis=dict(autorange='reversed', gridcolor='#1f1f1f'), height=420)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        by_unit = df['unit_type'].value_counts()
        fig2 = go.Figure(go.Pie(labels=by_unit.index, values=by_unit.values, hole=0.55,
                                marker=dict(colors=[PALETTE[0], PALETTE[4], PALETTE[3]], line=dict(color='#141414', width=2))))
        fig2.update_layout(**PLOTLY_LAYOUT, title="Channel Mix", height=420)
        st.plotly_chart(fig2, use_container_width=True)

    section("02", "Daily Trend & Top Issues")
    daily = df.groupby(df['Date'].dt.date).size()
    fig3 = go.Figure(go.Scatter(x=daily.index, y=daily.values, mode='lines+markers',
                                 line=dict(color=ACCENT, width=2), fill='tozeroy',
                                 fillcolor='rgba(237,28,36,.12)'))
    fig3.update_layout(**PLOTLY_LAYOUT, title="Daily Ticket Volume", height=320)
    st.plotly_chart(fig3, use_container_width=True)

    top_issues = df['type_3'].value_counts().head(10)
    top_labels = [l[:60] + '…' if len(l) > 60 else l for l in top_issues.index]
    fig4 = go.Figure(go.Bar(x=top_issues.values, y=top_labels, orientation='h',
                            marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(top_issues))]),
                            text=top_issues.values, textposition='outside'))
    fig4.update_layout(**PLOTLY_LAYOUT, title="Top 10 Issues (K-code · type_3)",
                       yaxis=dict(autorange='reversed', gridcolor='#1f1f1f'), height=440)
    st.plotly_chart(fig4, use_container_width=True)


# =============================================================
# COHORT RENDERER
# =============================================================
def render_cohort(df):
    st.markdown(f'<h1>Cohort LIS · Oct 2025 – Mar 2026 <span style="color:#737373;font-size:.5em;font-family:JetBrains Mono;font-style:normal">{fmt_num(len(df))} rows</span></h1>', unsafe_allow_html=True)
    total_lis = int(df['lis'].sum())
    total_ticket = int(df['lis_with_ticket'].sum())
    rate = total_ticket/total_lis*100

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total LIS", fmt_num(total_lis))
    k2.metric("With Ticket", fmt_num(total_ticket), f"{rate:.2f}%")
    k3.metric("Revenue", fmt_idr(df['rev_bill'].sum()))
    k4.metric("Usage", f"{fmt_num(df['usage_gb'].sum())} GB")
    k5.metric("Months", df['report_month'].nunique())

    section("01", "Tren Bulanan")
    monthly = df.groupby('report_month')[['lis', 'lis_with_ticket']].sum().reset_index()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=monthly['report_month'], y=monthly['lis_with_ticket'],
                             mode='lines+markers', name='With Ticket',
                             line=dict(color=ACCENT, width=2.5), fill='tozeroy',
                             fillcolor='rgba(237,28,36,.15)'))
    fig.update_layout(**PLOTLY_LAYOUT, title="LIS with Ticket per Bulan", height=360)
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        by_area = df.groupby('area')['lis'].sum().sort_values(ascending=True)
        fig2 = go.Figure(go.Bar(x=by_area.values, y=by_area.index, orientation='h',
                                marker=dict(color=PALETTE[:len(by_area)]),
                                text=by_area.values.astype(int), textposition='outside'))
        fig2.update_layout(**PLOTLY_LAYOUT, title="LIS by Area", height=380)
        st.plotly_chart(fig2, use_container_width=True)
    with c2:
        by_region = df.groupby('region')['lis'].sum().sort_values(ascending=False).head(10)
        fig3 = go.Figure(go.Bar(x=by_region.index, y=by_region.values,
                                marker=dict(color=PALETTE[:len(by_region)])))
        fig3.update_layout(**PLOTLY_LAYOUT, title="Top 10 Regions by LIS", height=380, xaxis=dict(tickangle=-30))
        st.plotly_chart(fig3, use_container_width=True)

    section("02", "Segmentasi")
    c3, c4 = st.columns(2)
    with c3:
        by_arpu = df.groupby('arpu_group')['lis'].sum().sort_index()
        fig4 = go.Figure(go.Bar(x=by_arpu.index, y=by_arpu.values,
                                marker=dict(color=PALETTE[:len(by_arpu)])))
        fig4.update_layout(**PLOTLY_LAYOUT, title="LIS by ARPU Group", height=340)
        st.plotly_chart(fig4, use_container_width=True)
    with c4:
        by_class = df.dropna(subset=['class_cat']).groupby('class_cat')['lis_with_ticket'].sum()
        fig5 = go.Figure(go.Pie(labels=by_class.index, values=by_class.values, hole=0.55,
                                marker=dict(colors=[PALETTE[0], PALETTE[1], PALETTE[4]], line=dict(color='#141414', width=2))))
        fig5.update_layout(**PLOTLY_LAYOUT, title="Tickets by Class Category", height=340)
        st.plotly_chart(fig5, use_container_width=True)


# =============================================================
# QA RENDERER
# =============================================================
def render_qa(sheets):
    raw = sheets['Raw Data']
    monthly = sheets['Monthly Trend']
    regional = sheets['Regional Analysis']
    top5 = sheets['Top 5 Issues']
    ceto = sheets['Ceto Latency']

    st.markdown(f'<h1>QA Product Q1 2026 <span style="color:#737373;font-size:.5em;font-family:JetBrains Mono;font-style:normal">{fmt_num(len(raw))} rows</span></h1>', unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Records", fmt_num(len(raw)))
    k2.metric("Products", raw['product'].nunique())
    k3.metric("Regions", raw['regional'].nunique())
    k4.metric("Top Product", raw['product'].value_counts().index[0])

    section("01", "Tren Bulanan per Produk")
    months = ['2026-01', '2026-02', '2026-03', '2026-04']
    fig = go.Figure()
    for i, row in monthly.iterrows():
        fig.add_trace(go.Scatter(x=months, y=[row[m] for m in months],
                                 mode='lines+markers', name=row['Produk-KIP'],
                                 line=dict(color=PALETTE[i % len(PALETTE)])))
    fig.update_layout(**PLOTLY_LAYOUT, title="Monthly Trend (Multi-Product)", height=400)
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        by_prod = raw['product'].value_counts()
        fig2 = go.Figure(go.Bar(x=by_prod.values, y=by_prod.index, orientation='h',
                                marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(by_prod))]),
                                text=by_prod.values, textposition='outside'))
        fig2.update_layout(**PLOTLY_LAYOUT, title="Records per Product",
                           yaxis=dict(autorange='reversed', gridcolor='#1f1f1f'), height=420)
        st.plotly_chart(fig2, use_container_width=True)
    with c2:
        by_reg = raw['regional'].value_counts()
        fig3 = go.Figure(go.Bar(x=by_reg.index, y=by_reg.values,
                                marker=dict(color=[PALETTE[(i+1) % len(PALETTE)] for i in range(len(by_reg))])))
        fig3.update_layout(**PLOTLY_LAYOUT, title="By Regional", height=420, xaxis=dict(tickangle=-30))
        st.plotly_chart(fig3, use_container_width=True)

    section("02", "Top Issues & Latency")
    c3, c4 = st.columns(2)
    with c3:
        top_iss = raw['type_3'].value_counts().head(10)
        top_lbl = [l[:55] + '…' if len(l) > 55 else l for l in top_iss.index]
        fig4 = go.Figure(go.Bar(x=top_iss.values, y=top_lbl, orientation='h',
                                marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(top_iss))]),
                                text=top_iss.values, textposition='outside'))
        fig4.update_layout(**PLOTLY_LAYOUT, title="Top 10 Issues",
                           yaxis=dict(autorange='reversed', gridcolor='#1f1f1f'), height=440)
        st.plotly_chart(fig4, use_container_width=True)
    with c4:
        ceto_sorted = ceto.sort_values('Avg. Latency (s)')
        colors = [ACCENT if v > 70 else C_ORANGE if v > 60 else C_GREEN for v in ceto_sorted['Avg. Latency (s)']]
        fig5 = go.Figure(go.Bar(x=ceto_sorted['Avg. Latency (s)'], y=ceto_sorted['Product Periode 20sd24Apr'],
                                orientation='h', marker=dict(color=colors),
                                text=[f"{v:.1f}s" for v in ceto_sorted['Avg. Latency (s)']],
                                textposition='outside'))
        fig5.update_layout(**PLOTLY_LAYOUT, title="Avg Latency per Product (s)",
                           yaxis=dict(gridcolor='#1f1f1f'), height=440)
        st.plotly_chart(fig5, use_container_width=True)

    section("03", "Top 5 Issues by Product")
    st.dataframe(top5, use_container_width=True, hide_index=True)


# =============================================================
# CUSTOM FILE RENDERER (auto-detected)
# =============================================================
def render_custom(label, df, col_types):
    st.markdown(f'<h1>{label} <span style="color:#737373;font-size:.5em;font-family:JetBrains Mono;font-style:normal">{fmt_num(len(df))} rows</span></h1>', unsafe_allow_html=True)

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Rows", fmt_num(len(df)))
    k2.metric("Columns", len(df.columns))
    k3.metric("Detected Types", f"{sum(1 for v in col_types.values() if v in ('date','number','category'))}/{len(col_types)} typed")

    date_cols = [c for c, t in col_types.items() if t == 'date']
    num_cols = [c for c, t in col_types.items() if t == 'number']
    cat_cols = [c for c, t in col_types.items() if t == 'category']

    section("01", "Auto-Detected Charts")

    if date_cols and num_cols:
        date_col = st.selectbox("Date column", date_cols, key=f"date_{label}")
        num_col = st.selectbox("Numeric column", num_cols, key=f"num_{label}")
        ts = df[[date_col, num_col]].dropna()
        ts[date_col] = pd.to_datetime(ts[date_col], errors='coerce')
        ts = ts.dropna().sort_values(date_col)
        fig = go.Figure(go.Scatter(x=ts[date_col], y=ts[num_col], mode='lines+markers',
                                    line=dict(color=ACCENT, width=2), fill='tozeroy',
                                    fillcolor='rgba(237,28,36,.12)'))
        fig.update_layout(**PLOTLY_LAYOUT, title=f"{num_col} over {date_col}", height=360)
        st.plotly_chart(fig, use_container_width=True)

    if cat_cols:
        c1, c2 = st.columns([3, 2])
        with c1:
            cat = st.selectbox("Category breakdown", cat_cols, key=f"cat_{label}")
            counts = df[cat].value_counts().head(15)
            fig2 = go.Figure(go.Bar(x=counts.values, y=[str(x)[:40] for x in counts.index],
                                    orientation='h',
                                    marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(counts))]),
                                    text=counts.values, textposition='outside'))
            fig2.update_layout(**PLOTLY_LAYOUT, title=f"Distribution: {cat}",
                               yaxis=dict(autorange='reversed', gridcolor='#1f1f1f'), height=420)
            st.plotly_chart(fig2, use_container_width=True)
        with c2:
            if num_cols:
                nc = num_cols[0]
                vals = pd.to_numeric(df[nc], errors='coerce').dropna()
                fig3 = go.Figure(go.Histogram(x=vals, nbinsx=20, marker=dict(color=C_GREEN)))
                fig3.update_layout(**PLOTLY_LAYOUT, title=f"Histogram: {nc}", height=420)
                st.plotly_chart(fig3, use_container_width=True)

    section("02", "Data Preview")
    st.dataframe(df.head(200), use_container_width=True, hide_index=True)
    if len(df) > 200:
        st.caption(f"Showing first 200 rows of {fmt_num(len(df))} total")


# =============================================================
# UPLOAD HANDLER
# =============================================================
def render_upload_sidebar():
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📁 Add File")

    uploaded = st.sidebar.file_uploader("Upload XLSX/XLS", type=["xlsx", "xls"], key="uploader")
    if not uploaded:
        return

    file_bytes = uploaded.getvalue()
    file_key = f"upload_{uploaded.name}"

    if file_key not in st.session_state:
        # Parse workbook to discover sheets
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets_info = {}
        for sn in xl.sheet_names:
            try:
                raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sn, header=None)
                rows = raw_df.values.tolist()[:10]
                auto_row = detect_header_row(rows)
                sheets_info[sn] = {'auto_row': auto_row, 'header_row': auto_row, 'included': False}
            except Exception as e:
                continue
        st.session_state[file_key] = sheets_info

    sheets_info = st.session_state[file_key]
    st.sidebar.markdown(f"**{uploaded.name}**  \n<span style='color:#737373;font-size:11px;font-family:JetBrains Mono'>{len(sheets_info)} sheets detected</span>", unsafe_allow_html=True)

    # Pick sheet
    sheet_name = st.sidebar.selectbox("Pilih sheet:", list(sheets_info.keys()), key=f"sel_{uploaded.name}")
    info = sheets_info[sheet_name]

    # Header row input
    header_row = st.sidebar.number_input(
        f"Header row (auto-detected: {info['auto_row']})",
        min_value=0, max_value=9, value=info['header_row'],
        key=f"hr_{uploaded.name}_{sheet_name}",
    )
    info['header_row'] = header_row

    # Preview + import
    try:
        df = read_sheet_smart(file_bytes, sheet_name, header_row)
        col_types = {c: classify_column(df[c]) for c in df.columns}
        quality = assess_quality(df)
        qcolor = {'great': '#5DD39E', 'ok': '#FFC93C', 'poor': '#ED1C24'}[quality]
        qlabel = {'great': '🟢 Great', 'ok': '🟡 OK', 'poor': '🔴 Poor'}[quality]

        st.sidebar.markdown(f"<div style='font-size:12px;margin:6px 0'>Quality: <b style='color:{qcolor}'>{qlabel}</b> · {len(df)} rows · {len(df.columns)} cols</div>", unsafe_allow_html=True)

        # Preview
        with st.sidebar.expander("Preview"):
            st.dataframe(df.head(5), hide_index=True)

        if st.sidebar.button(f"Import sheet '{sheet_name}'", type="primary", use_container_width=True):
            # Add to active sheets
            if 'imported' not in st.session_state:
                st.session_state['imported'] = {}
            label = f"{uploaded.name.rsplit('.',1)[0]} · {sheet_name}"
            st.session_state['imported'][label] = {'df': df, 'col_types': col_types}
            st.sidebar.success(f"✓ Imported as '{label}'")
            st.rerun()

    except Exception as e:
        st.sidebar.error(f"Error parsing: {e}")


# =============================================================
# MAIN
# =============================================================
def main():
    inject_css()

    # Sidebar nav
    st.sidebar.markdown('<h2 style="margin-top:0">◆ QA Dashboard</h2>', unsafe_allow_html=True)
    st.sidebar.markdown('<p style="color:#737373;font-size:11px;letter-spacing:.1em;text-transform:uppercase;margin-top:-10px;font-weight:600">TELKOMSEL · 2026</p>', unsafe_allow_html=True)

    base_views = ["★ Executive Summary", "KIP April 2026", "Cohort LIS", "QA Product Q1"]
    imported_views = list(st.session_state.get('imported', {}).keys())
    all_views = base_views + imported_views

    view = st.sidebar.radio("Pilih view:", all_views, label_visibility="collapsed")

    if imported_views:
        st.sidebar.markdown("**Imported files:**")
        for v in imported_views:
            cc1, cc2 = st.sidebar.columns([5, 1])
            cc1.markdown(f"<span style='font-size:12px;color:#a3a3a3'>· {v}</span>", unsafe_allow_html=True)
            if cc2.button("×", key=f"del_{v}", help="Hapus"):
                del st.session_state['imported'][v]
                st.rerun()

    render_upload_sidebar()

    # Load data
    try:
        kip = load_kip()
        coh = load_cohort()
        qa = load_qa()
    except FileNotFoundError as e:
        st.error(f"File data tidak ditemukan: {e}. Pastikan folder `data/` berisi 3 file CLEANED.xlsx.")
        return

    # Route
    if view == "★ Executive Summary":
        render_executive(kip, coh, qa)
    elif view == "KIP April 2026":
        render_kip(kip)
    elif view == "Cohort LIS":
        render_cohort(coh)
    elif view == "QA Product Q1":
        render_qa(qa)
    elif view in st.session_state.get('imported', {}):
        item = st.session_state['imported'][view]
        render_custom(view, item['df'], item['col_types'])

    # Footer
    st.markdown("""
    <div style="margin-top:60px;padding-top:18px;border-top:1px solid #2a2a2a;text-align:center;color:#737373;font-size:12px">
      QA Product Dashboard · Streamlit · Telkomsel 2026
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
