"""
app.py — VAULT: Personal Finance Tracker
Run: streamlit run app.py
"""

import streamlit as st
from datetime import datetime, date
import calendar
from collections import defaultdict

from data_manager import (init_data, load_data, save_data,
                          add_entry, mark_paid, award_points, update_streak, delete_entry)
from logic import (calculate_balance, calculate_pending, estimate_days_for_reward,
                   monthly_income_projection, actual_monthly_income, compute_daily_saving)
from export_excel import export_to_excel

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="VAULT — Finance Tracker",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — all variables in one place
# BG_DEEP  = #0a0b0f   (almost black)
# BG_CARD  = #12141a   (dark card)
# BG_RAISE = #1a1d27   (slightly raised surface)
# BG_EDGE  = #22263a   (border / hover)
# ACCENT   = #6c63ff   (electric violet — primary)
# ACCENT2  = #00d4aa   (teal — secondary)
# GREEN    = #22c55e
# RED      = #ef4444
# AMBER    = #f59e0b
# TEXT_HI  = #f4f4f5   (high contrast white)
# TEXT_MID = #a1a1aa   (mid grey — still readable)
# TEXT_LOW = #71717a   (lowest readable — only for hints)
# ═══════════════════════════════════════════════════════════════════════════════

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=JetBrains+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');

/* ── RESET & BASE ─────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    background: #0a0b0f !important;
    color: #f4f4f5 !important;
}

.stApp {
    background: #0a0b0f !important;
}

/* ── SIDEBAR ──────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #0d0e14 !important;
    border-right: 1px solid #22263a !important;
}
/* The toggle arrow button — make it always visible */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"] {
    background: #1a1d27 !important;
    border: 1px solid #6c63ff !important;
    border-radius: 50% !important;
    color: #f4f4f5 !important;
}
[data-testid="stSidebarCollapseButton"]:hover,
[data-testid="collapsedControl"]:hover {
    background: #6c63ff !important;
}
[data-testid="stSidebarCollapseButton"] svg,
[data-testid="collapsedControl"] svg {
    fill: #f4f4f5 !important;
    stroke: #f4f4f5 !important;
}

/* ALL sidebar text — 100% explicit, no inheritance hell */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] li,
[data-testid="stSidebar"] a,
[data-testid="stSidebar"] strong,
[data-testid="stSidebar"] em,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] .stMarkdown *,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {
    color: #f4f4f5 !important;
}

/* ── MAIN AREA GLOBAL TEXT ────────────────────────────────────────────── */
.main .block-container p,
.main .block-container span,
.main .block-container label,
.main .block-container div,
.main .block-container h1,
.main .block-container h2,
.main .block-container h3,
.main .block-container h4,
.main .block-container li,
.main .block-container strong {
    color: #f4f4f5 !important;
}

/* ── STREAMLIT WIDGET LABELS ─────────────────────────────────────────── */
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] *,
.stSelectbox label,
.stNumberInput label,
.stTextInput label,
.stRadio label,
.stDateInput label,
.stToggle label,
.stRadio [data-testid="stMarkdownContainer"] p,
.stCheckbox label {
    color: #a1a1aa !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    letter-spacing: .04em !important;
    text-transform: uppercase !important;
}

/* ── TABS ─────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: #12141a !important;
    border-radius: 10px !important;
    padding: 4px !important;
    gap: 4px !important;
    border: 1px solid #22263a !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 7px !important;
    color: #a1a1aa !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    padding: 8px 18px !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: #6c63ff !important;
    color: #ffffff !important;
}

/* ── INPUTS ───────────────────────────────────────────────────────────── */
.stTextInput input,
.stNumberInput input,
.stDateInput input {
    background: #1a1d27 !important;
    border: 1.5px solid #22263a !important;
    border-radius: 8px !important;
    color: #f4f4f5 !important;
    font-family: 'Inter', sans-serif !important;
}
.stTextInput input:focus,
.stNumberInput input:focus,
.stDateInput input:focus {
    border-color: #6c63ff !important;
    box-shadow: 0 0 0 3px rgba(108,99,255,.2) !important;
}
.stSelectbox > div > div {
    background: #1a1d27 !important;
    border: 1.5px solid #22263a !important;
    border-radius: 8px !important;
    color: #f4f4f5 !important;
}

/* ── RADIO ────────────────────────────────────────────────────────────── */
.stRadio > div {
    gap: 10px !important;
}
.stRadio [data-testid="stMarkdownContainer"] p {
    color: #f4f4f5 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}

/* ── TOGGLE ───────────────────────────────────────────────────────────── */
.stToggle [data-testid="stMarkdownContainer"] p {
    color: #f4f4f5 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}

/* ── BUTTONS ──────────────────────────────────────────────────────────── */
.stButton > button {
    background: #6c63ff !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    padding: 10px 20px !important;
    width: 100% !important;
    transition: all .18s ease !important;
    letter-spacing: .02em !important;
}
.stButton > button:hover {
    background: #5b52e8 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(108,99,255,.4) !important;
}
.stButton > button:active {
    transform: translateY(0) !important;
}

/* ── EXPANDER ─────────────────────────────────────────────────────────── */
details {
    background: #12141a !important;
    border: 1px solid #22263a !important;
    border-radius: 10px !important;
}
details summary {
    color: #f4f4f5 !important;
    font-weight: 600 !important;
    padding: 12px 16px !important;
}
details[open] summary { border-bottom: 1px solid #22263a !important; }

/* ── DATAFRAME ────────────────────────────────────────────────────────── */
.stDataFrame,
.stDataFrame *,
[data-testid="stDataFrame"],
[data-testid="stDataFrame"] * {
    color: #f4f4f5 !important;
    background: #12141a !important;
}
.stDataFrame th {
    background: #1a1d27 !important;
    color: #a1a1aa !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: .07em !important;
}

/* ── CAPTIONS ─────────────────────────────────────────────────────────── */
[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] * {
    color: #71717a !important;
}

/* ── SUCCESS / ERROR ──────────────────────────────────────────────────── */
[data-testid="stAlert"] {
    background: #12141a !important;
    border-radius: 8px !important;
    border: 1px solid #22263a !important;
}
[data-testid="stAlert"] p { color: #f4f4f5 !important; }
.stSuccess [data-testid="stAlert"] { border-left: 4px solid #22c55e !important; }
.stError   [data-testid="stAlert"] { border-left: 4px solid #ef4444 !important; }
.stWarning [data-testid="stAlert"] { border-left: 4px solid #f59e0b !important; }

/* ── PROGRESS ─────────────────────────────────────────────────────────── */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #6c63ff, #00d4aa) !important;
}

/* ── MARKDOWN (main) ──────────────────────────────────────────────────── */
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {
    color: #f4f4f5 !important;
}
.stMarkdown p { color: #f4f4f5 !important; }
.stMarkdown strong { color: #f4f4f5 !important; }

/* ── CUSTOM COMPONENTS ────────────────────────────────────────────────── */

/* Page title */
.vault-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.8rem;
    font-weight: 800;
    background: linear-gradient(135deg, #6c63ff 0%, #00d4aa 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height: 1;
    margin-bottom: 4px;
    display: inline-block;
}
.vault-sub {
    font-size: 13px;
    color: #71717a !important;
    margin-bottom: 28px;
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: .06em;
}

/* Dashboard stat card */
.stat-card {
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 14px;
    padding: 20px 22px 16px;
    border-top: 3px solid #6c63ff;
    position: relative;
    overflow: hidden;
}
.stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 60px;
    background: linear-gradient(180deg, rgba(108,99,255,.06), transparent);
    pointer-events: none;
}
.stat-lbl {
    font-size: 10px;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: #71717a;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 500;
    margin-bottom: 8px;
}
.stat-val {
    font-family: 'Syne', sans-serif;
    font-size: 1.7rem;
    font-weight: 700;
    color: #f4f4f5;
    line-height: 1.1;
}
.stat-sub {
    font-size: 11px;
    color: #71717a;
    margin-top: 6px;
    font-family: 'JetBrains Mono', monospace;
}

/* Section header */
.sec-hdr {
    font-family: 'Syne', sans-serif;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .22em;
    text-transform: uppercase;
    color: #6c63ff;
    border-bottom: 1px solid #22263a;
    padding-bottom: 8px;
    margin: 36px 0 18px;
}

/* Form card */
.form-card {
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 14px;
    padding: 24px 26px 22px;
}
.form-card-title {
    font-family: 'Syne', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    color: #f4f4f5;
    margin-bottom: 18px;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* Sidebar cards */
.sb-block {
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 10px;
    padding: 12px 14px;
    margin-bottom: 8px;
}
.sb-lbl {
    font-size: 9px;
    letter-spacing: .15em;
    text-transform: uppercase;
    color: #71717a !important;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 500;
    margin-bottom: 4px;
}
.sb-val {
    font-family: 'JetBrains Mono', monospace;
    font-size: 19px;
    font-weight: 600;
    color: #f4f4f5 !important;
}
.c-green  { color: #22c55e !important; }
.c-red    { color: #ef4444 !important; }
.c-amber  { color: #f59e0b !important; }
.c-violet { color: #6c63ff !important; }
.c-teal   { color: #00d4aa !important; }
.c-mid    { color: #a1a1aa !important; }

/* Streak */
.streak-card {
    background: linear-gradient(135deg, #1a0a3a 0%, #0d1a3a 100%);
    border: 1px solid #6c63ff;
    border-radius: 12px;
    padding: 16px 14px 12px;
    text-align: center;
    margin-bottom: 8px;
    box-shadow: 0 0 20px rgba(108,99,255,.15), inset 0 1px 0 rgba(255,255,255,.05);
}
.streak-num {
    font-family: 'Syne', sans-serif;
    font-size: 40px;
    font-weight: 800;
    background: linear-gradient(135deg, #a78bfa, #6c63ff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height: 1;
}
.streak-lbl {
    font-size: 9px;
    letter-spacing: .15em;
    text-transform: uppercase;
    color: #a78bfa !important;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 600;
    margin-top: 4px;
}

/* Points */
.pts-card {
    background: linear-gradient(135deg, #0a1f12 0%, #0a1a0d 100%);
    border: 1px solid #22c55e;
    border-radius: 10px;
    padding: 10px 14px;
    text-align: center;
    margin-bottom: 8px;
    box-shadow: 0 0 14px rgba(34,197,94,.1);
}
.pts-lbl { color: #4ade80 !important; font-size: 9px; letter-spacing:.15em; text-transform:uppercase; font-family:'JetBrains Mono',monospace; }
.pts-val  { color: #22c55e !important; font-family:'JetBrains Mono',monospace; font-size:18px; font-weight:700; }

/* Progress bar */
.prog-bg {
    background: #22263a;
    border-radius: 99px;
    height: 6px;
    overflow: hidden;
    margin: 6px 0 3px;
}
.prog-fill {
    height: 100%;
    border-radius: 99px;
    background: linear-gradient(90deg, #6c63ff, #00d4aa);
}

/* Reward card sidebar */
.rw-block {
    background: #0f1018;
    border: 1px solid #22263a;
    border-radius: 9px;
    padding: 10px 12px;
    margin-bottom: 7px;
}
.rw-name { font-size: 12px; font-weight: 600; color: #f4f4f5 !important; }
.rw-cost { font-size: 10px; color: #a1a1aa !important; font-family: 'JetBrains Mono', monospace; }
.rw-pct  { font-size: 10px; color: #71717a !important; font-family: 'JetBrains Mono', monospace; }
.rw-block a { color: #6c63ff !important; }

/* Bonus pill */
.bonus-pill {
    display: inline-block;
    background: rgba(245,158,11,.12);
    border: 1px solid rgba(245,158,11,.3);
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 11px;
    color: #f59e0b !important;
    margin: 2px;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 600;
}

/* Category tag */
.cat-tag {
    display: inline-block;
    background: #1a1d27;
    border: 1px solid #22263a;
    border-radius: 6px;
    padding: 4px 12px;
    font-size: 11px;
    color: #a1a1aa !important;
    margin: 3px;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 500;
}

/* Finance tip */
.tip-card {
    background: #0d1218;
    border: 1px solid #1e3a2e;
    border-left: 3px solid #00d4aa;
    border-radius: 10px;
    padding: 14px 18px;
    font-size: 13px;
    color: #a1a1aa !important;
    margin-bottom: 20px;
}
.tip-term {
    font-weight: 700;
    color: #00d4aa !important;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    letter-spacing: .04em;
}
.tip-body { color: #d4d4d8 !important; }

/* Analytics table row */
.ana-row {
    display: grid;
    grid-template-columns: 80px 1fr 1fr 1fr 1fr;
    gap: 0;
    padding: 10px 14px;
    border-bottom: 1px solid #22263a;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    align-items: center;
}
.ana-row:last-child { border-bottom: none; }
.ana-row.hdr {
    background: #1a1d27;
    border-radius: 8px 8px 0 0;
    color: #71717a !important;
    font-size: 10px;
    letter-spacing: .1em;
    text-transform: uppercase;
    font-weight: 600;
    border-bottom: 1px solid #6c63ff;
}
.ana-row.hdr span { color: #71717a !important; }
.ana-row span { color: #f4f4f5 !important; }
.ana-row .pos  { color: #22c55e !important; }
.ana-row .neg  { color: #ef4444 !important; }
.ana-row .ambi { color: #6c63ff !important; }
.ana-wrap {
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 10px;
    overflow: hidden;
}

/* Entry line */
.entry-line {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 9px 14px;
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 8px;
    margin-bottom: 6px;
    font-size: 13px;
}
.entry-amt  { font-family: 'JetBrains Mono',monospace; font-weight:600; color:#f4f4f5 !important; min-width:100px; }
.entry-meta { color: #71717a !important; font-size:11px; flex:1; font-family:'JetBrains Mono',monospace; }
.entry-badge-g { background:#0a2010; border:1px solid #22c55e; border-radius:4px; padding:2px 7px; font-size:10px; color:#22c55e !important; font-weight:700; }
.entry-badge-r { background:#1f0a0a; border:1px solid #ef4444; border-radius:4px; padding:2px 7px; font-size:10px; color:#ef4444 !important; font-weight:700; }
.entry-badge-y { background:#1f1500; border:1px solid #f59e0b; border-radius:4px; padding:2px 7px; font-size:10px; color:#f59e0b !important; font-weight:700; }

/* Delete button — small red variant */
.stButton.del > button,
button[kind="secondary"].del {
    background: #1f0a0a !important;
    border: 1px solid #ef4444 !important;
    color: #ef4444 !important;
    font-size: 11px !important;
    padding: 5px 10px !important;
    width: auto !important;
}
.stButton.del > button:hover {
    background: #ef4444 !important;
    color: #fff !important;
    box-shadow: 0 2px 8px rgba(239,68,68,.35) !important;
}

/* Export info box */
.export-info {
    background: #12141a;
    border: 1px solid #22263a;
    border-radius: 10px;
    padding: 16px 20px;
    font-size: 13px;
    color: #a1a1aa !important;
    line-height: 1.7;
}
.export-info strong { color: #f4f4f5 !important; }

/* Divider */
.vdivider {
    border: none;
    border-top: 1px solid #22263a;
    margin: 14px 0;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# DATA + COMPUTED VALUES
# ═══════════════════════════════════════════════════════════════════════════════

data           = init_data()
balance_info   = calculate_balance(data)
to_pay, to_rcv = calculate_pending(data)
net            = balance_info["net_balance"]
streak         = data.get("streak", 0)
points         = data.get("points", 0)
daily_saving   = compute_daily_saving(data)

# ═══════════════════════════════════════════════════════════════════════════════
# FINANCE TIPS
# ═══════════════════════════════════════════════════════════════════════════════

TIPS = [
    ("Cash Flow",       "Net movement of money in/out over a period. Positive = earning more than spending."),
    ("Liquidity",       "How quickly an asset converts to cash. Your bank balance is perfectly liquid."),
    ("Receivable",      "Money owed TO you — an asset. Until collected, it's not spendable."),
    ("Payable",         "Money you owe others — a liability. Clearing payables improves net worth."),
    ("Net Worth",       "Total assets minus total liabilities. The single truest snapshot of financial health."),
    ("Opportunity Cost","The value of the next-best option you give up. Every ₹ spent has one."),
    ("Amortisation",    "Spreading a cost over time. Saving ₹X/day toward a goal amortises the purchase."),
    ("Compound Growth", "Returns earned on previous returns. Even small daily savings compound massively."),
    ("50/30/20 Rule",   "Budget guideline: 50% needs · 30% wants · 20% savings. Start here."),
    ("Emergency Fund",  "3–6 months of expenses in liquid savings. Your financial safety net before investing."),
    ("Sunk Cost",       "Money already spent that can't be recovered. Don't let it drive future decisions."),
    ("Dollar-Cost Avg", "Investing a fixed amount regularly regardless of price — reduces timing risk."),
]
tip_term, tip_body = TIPS[date.today().timetuple().tm_yday % len(TIPS)]

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

def sb(label, value, cls=""):
    st.markdown(f"""
    <div class="sb-block">
        <div class="sb-lbl">{label}</div>
        <div class="sb-val {cls}">{value}</div>
    </div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("""
    <div style="padding:4px 0 12px">
        <div style="font-family:'Syne',sans-serif;font-size:1.4rem;font-weight:800;
             background:linear-gradient(135deg,#6c63ff,#00d4aa);
             -webkit-background-clip:text;-webkit-text-fill-color:transparent;
             background-clip:text;">🔐 VAULT</div>
        <div style="font-size:10px;color:#71717a;font-family:'JetBrains Mono',monospace;
             letter-spacing:.1em;margin-top:2px">PERSONAL FINANCE OS</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)

    # Streak
    flames = "🔥" * min(streak, 5) if streak > 0 else "—"
    st.markdown(f"""
    <div class="streak-card">
        <div class="streak-num">{streak}</div>
        <div class="streak-lbl">{flames} Day Streak</div>
    </div>""", unsafe_allow_html=True)

    # Points
    st.markdown(f"""
    <div class="pts-card">
        <div class="pts-lbl">⭐ Total Points</div>
        <div class="pts-val">{points:,} pts</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)

    sb("Net Balance",   f"₹{net:,.2f}",         "c-green" if net >= 0 else "c-red")
    sb("To Pay",        f"₹{to_pay:,.2f}",       "c-amber" if to_pay > 0 else "c-mid")
    sb("To Receive",    f"₹{to_rcv:,.2f}",       "c-teal"  if to_rcv > 0 else "c-mid")
    sb("Daily Saving",  f"₹{daily_saving:,.0f}", "c-violet")

    # Rewards progress
    rewards = data.get("rewards", [])
    if rewards:
        st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)
        st.markdown("<div class='sb-lbl' style='margin-bottom:8px;padding:0 2px'>🎯 REWARD GOALS</div>",
                    unsafe_allow_html=True)
        for rw in rewards:
            cost = float(rw.get("cost", 0) or 0)
            pct  = max(0.0, min(100.0, (net / cost * 100) if cost > 0 else 100))
            link = (f"<a href='{rw['link']}' target='_blank' style='color:#6c63ff'>↗</a> "
                    if rw.get("link") else "")
            tgt  = rw.get("target_date", "")
            tgt_html = f"<div class='rw-pct'>🗓 {tgt}</div>" if tgt else ""
            st.markdown(f"""
            <div class="rw-block">
                <div class="rw-name">{rw['name']} {link}</div>
                <div class="rw-cost">₹{cost:,.0f}</div>
                {tgt_html}
                <div class="prog-bg"><div class="prog-fill" style="width:{pct:.1f}%"></div></div>
                <div class="rw-pct">{pct:.0f}% saved</div>
            </div>""", unsafe_allow_html=True)

    # Open bonuses
    open_b = [b for b in data.get("bonuses", []) if not b.get("hidden")]
    if open_b:
        st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)
        st.markdown("<div class='sb-lbl' style='margin-bottom:6px;padding:0 2px'>🎁 OPEN BONUSES</div>",
                    unsafe_allow_html=True)
        pills = "".join(
            f"<span class='bonus-pill'>+₹{float(b.get('amount',0)):,.0f}"
            f"{' · '+b['note'] if b.get('note') else ''}</span>"
            for b in open_b
        )
        st.markdown(f"<div style='line-height:2'>{pills}</div>", unsafe_allow_html=True)

    hidden_ct = sum(1 for b in data.get("bonuses", []) if b.get("hidden"))
    if hidden_ct:
        st.markdown(f"<p style='font-size:11px;color:#71717a;font-family:JetBrains Mono,monospace;"
                    f"margin-top:6px'>🔒 {hidden_ct} hidden bonus stored.</p>",
                    unsafe_allow_html=True)

    st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:10px;color:#71717a;font-family:JetBrains Mono,monospace'>"
                f"{datetime.now().strftime('%d %b %Y · %H:%M')}</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN — HEADER
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='vault-title'>VAULT</div>", unsafe_allow_html=True)
st.markdown("<div class='vault-sub'>personal finance os · track · save · conquer</div>",
            unsafe_allow_html=True)

st.markdown(f"""
<div class="tip-card">
    <span class="tip-term">💡 {tip_term}</span>
    <span class="tip-body"> — {tip_body}</span>
</div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Dashboard</div>", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5, gap="small")

def stat(col, label, value, sub, color, accent_color="#6c63ff"):
    col.markdown(f"""
    <div class="stat-card" style="border-top-color:{accent_color}">
        <div class="stat-lbl">{label}</div>
        <div class="stat-val" style="color:{color}">{value}</div>
        <div class="stat-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

stat(c1, "Net Balance",    f"₹{net:,.0f}",
     "income + bonuses − expenses",
     "#22c55e" if net >= 0 else "#ef4444",
     "#22c55e" if net >= 0 else "#ef4444")
stat(c2, "Total Income",   f"₹{balance_info['total_income']:,.0f}",
     f"+ ₹{balance_info['total_bonuses']:,.0f} in bonuses", "#6c63ff", "#6c63ff")
stat(c3, "Total Expenses", f"₹{balance_info['total_expenses']:,.0f}",
     "all recorded outflows", "#ef4444", "#ef4444")
stat(c4, "Pending",        f"₹{to_pay:,.0f}",
     f"₹{to_rcv:,.0f} due to me",
     "#f59e0b" if to_pay > 0 else "#a1a1aa", "#f59e0b")
stat(c5, "Daily Saving",   f"₹{daily_saving:,.0f}",
     "avg surplus per day", "#00d4aa", "#00d4aa")

# ═══════════════════════════════════════════════════════════════════════════════
# ANALYTICS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Analytics</div>", unsafe_allow_html=True)

view_mode = st.radio("View by", ["Month", "Week", "Year"],
                     horizontal=True, key="ana_mode")

# ── helpers ───────────────────────────────────────────────────────────────────

def _parse_dt(entry):
    raw = entry.get("date") or entry.get("created_at","")
    try:
        return datetime.fromisoformat(str(raw)[:19])
    except:
        return None

def _safe_amt(e):
    try: return float(e.get("amount", 0) or 0)
    except: return 0.0

def _build_daily(entries_inc, entries_exp, date_filter_fn):
    """Return {date_str: {income, expense, saved}} for entries matching filter."""
    rows = defaultdict(lambda: {"income": 0.0, "expense": 0.0})
    for e in entries_inc:
        dt = _parse_dt(e)
        if dt and date_filter_fn(dt):
            rows[dt.date().isoformat()]["income"] += _safe_amt(e)
    for e in entries_exp:
        dt = _parse_dt(e)
        if dt and date_filter_fn(dt):
            rows[dt.date().isoformat()]["expense"] += _safe_amt(e)
    result = {}
    for d in sorted(rows):
        inc = rows[d]["income"]
        exp = rows[d]["expense"]
        result[d] = {"income": inc, "expense": exp, "saved": inc - exp}
    return result

def _show_ana_table(daily_data, period_label):
    if not daily_data:
        st.markdown("<div style='color:#71717a;padding:20px;font-family:JetBrains Mono,monospace;"
                    "font-size:13px'>No entries for this period.</div>", unsafe_allow_html=True)
        return

    total_inc = sum(v["income"]  for v in daily_data.values())
    total_exp = sum(v["expense"] for v in daily_data.values())
    total_sav = total_inc - total_exp

    # Period summary strip
    s1, s2, s3 = st.columns(3, gap="small")
    def mini(col, lbl, val, cls):
        col.markdown(f"""
        <div style="background:#12141a;border:1px solid #22263a;border-radius:10px;
                    padding:14px 16px;margin-bottom:4px">
            <div style="font-size:9px;letter-spacing:.14em;text-transform:uppercase;
                        color:#71717a;font-family:'JetBrains Mono',monospace;margin-bottom:5px">{lbl}</div>
            <div style="font-family:'JetBrains Mono',monospace;font-size:1.3rem;
                        font-weight:700" class="{cls}">₹{val:,.0f}</div>
        </div>""", unsafe_allow_html=True)
    mini(s1, f"{period_label} income",   total_inc, "c-green")
    mini(s2, f"{period_label} expenses", total_exp, "c-red")
    mini(s3, f"{period_label} saved",    total_sav, "c-violet" if total_sav >= 0 else "c-red")

    # Date-wise table
    header = """
    <div class="ana-wrap">
        <div class="ana-row hdr">
            <span>Date</span>
            <span>Income (₹)</span>
            <span>Expense (₹)</span>
            <span>Saved (₹)</span>
            <span>Status</span>
        </div>"""
    rows_html = ""
    for d, v in daily_data.items():
        inc = v["income"]; exp = v["expense"]; sav = v["saved"]
        status = "surplus" if sav > 0 else ("break-even" if sav == 0 else "deficit")
        s_cls  = "pos" if sav > 0 else ("c-mid" if sav == 0 else "neg")
        rows_html += f"""
        <div class="ana-row">
            <span>{d[5:]}</span>
            <span class="pos">₹{inc:,.0f}</span>
            <span class="neg">₹{exp:,.0f}</span>
            <span class="{s_cls}">₹{sav:,.0f}</span>
            <span class="{s_cls}">{status}</span>
        </div>"""
    st.markdown(header + rows_html + "</div>", unsafe_allow_html=True)

# ── Mode: Month ───────────────────────────────────────────────────────────────

all_inc = data.get("income", [])
all_exp = data.get("expenses", [])

if view_mode == "Month":
    today = date.today()
    months_available = sorted(set(
        _parse_dt(e).strftime("%Y-%m")
        for e in all_inc + all_exp
        if _parse_dt(e)
    ), reverse=True) or [today.strftime("%Y-%m")]

    month_labels = {m: datetime.strptime(m, "%Y-%m").strftime("%B %Y") for m in months_available}
    chosen = st.selectbox("Select month", months_available,
                          format_func=lambda m: month_labels[m], key="ana_month")
    y, mo = int(chosen[:4]), int(chosen[5:])
    daily = _build_daily(all_inc, all_exp,
                         lambda dt: dt.year == y and dt.month == mo)
    _show_ana_table(daily, month_labels[chosen])

elif view_mode == "Week":
    today = date.today()
    # collect all ISO weeks present in data
    weeks = sorted(set(
        _parse_dt(e).strftime("%Y-W%W")
        for e in all_inc + all_exp
        if _parse_dt(e)
    ), reverse=True) or [today.strftime("%Y-W%W")]

    chosen_w = st.selectbox("Select week", weeks, key="ana_week")
    y_w, wk  = int(chosen_w[:4]), int(chosen_w[6:])
    daily = _build_daily(all_inc, all_exp,
                         lambda dt: dt.year == y_w and int(dt.strftime("%W")) == wk)
    _show_ana_table(daily, f"Week {wk}, {y_w}")

elif view_mode == "Year":
    years = sorted(set(
        _parse_dt(e).year
        for e in all_inc + all_exp
        if _parse_dt(e)
    ), reverse=True) or [date.today().year]

    chosen_y = st.selectbox("Select year", years, key="ana_year")
    daily = _build_daily(all_inc, all_exp,
                         lambda dt: dt.year == chosen_y)
    _show_ana_table(daily, str(chosen_y))

# ═══════════════════════════════════════════════════════════════════════════════
# INCOME SOURCE PROFILE
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Income Sources</div>", unsafe_allow_html=True)

with st.expander("➕ Add income source", expanded=False):
    sc1, sc2 = st.columns(2, gap="large")
    with sc1:
        src_name = st.text_input("Source name", placeholder="Salary, Freelance, Rental…",
                                 key="src_name")
        src_mode = st.radio("Income type",
                            ["Monthly (fixed — same every month)",
                             "Daily (per day — more in 31-day months)"],
                            key="src_mode")
        src_amount  = st.number_input("Amount per period (₹)", min_value=0.0,
                                      step=100.0, key="src_amount")
        src_status  = st.radio("Status", ["Already credited ✅", "Yet to be paid ⏳"],
                               horizontal=True, key="src_status")
        if st.button("Save Source", key="btn_src"):
            if not src_name.strip() or src_amount <= 0:
                st.error("Fill in source name and amount.")
            else:
                mode_key = "monthly" if "Monthly" in src_mode else "daily"
                profile  = {
                    "source": src_name.strip(), "mode": mode_key,
                    "amount": src_amount,
                    "status": "received" if "credited" in src_status else "pending",
                    "date":   datetime.now().isoformat(timespec="seconds"),
                }
                d = load_data()
                d["income_sources"].append(profile)
                add_entry("income", dict(profile))
                save_data(d)
                st.success(f"✓ '{src_name.strip()}' added! +10 pts")
                st.rerun()

    with sc2:
        sources = data.get("income_sources", [])
        if not sources:
            st.caption("No sources yet.")
        else:
            today = date.today()
            st.markdown("**Projected income this month:**")
            for si, s in enumerate(sources):
                mode = s.get("mode", "monthly")
                amt  = float(s.get("amount", 0))
                proj = actual_monthly_income(amt, mode, today.month, today.year)
                days = calendar.monthrange(today.year, today.month)[1]
                note = (f"₹{amt:,.0f}/day × {days} days" if mode == "daily"
                        else f"₹{amt:,.0f} fixed/month")
                c1, c2 = st.columns([9, 1])
                with c1:
                    st.markdown(
                        f"<div class='entry-line'>"
                        f"<span class='entry-amt c-teal'>₹{proj:,.0f}</span>"
                        f"<span class='entry-meta'>{s.get('source','—')} · {note}</span>"
                        f"</div>", unsafe_allow_html=True)
                with c2:
                    if st.button("🗑", key=f"del_src_{si}", help="Remove source"):
                        d = load_data()
                        d["income_sources"].pop(si)
                        save_data(d)
                        st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# LOG ENTRY
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Log Entry</div>", unsafe_allow_html=True)

col_inc, col_exp = st.columns(2, gap="large")

with col_inc:
    st.markdown("<div class='form-card'>"
                "<div class='form-card-title'>💵 Add Income</div>", unsafe_allow_html=True)
    inc_source = st.text_input("Source", placeholder="Salary, Freelance…", key="inc_source")
    inc_amount = st.number_input("Amount (₹)", min_value=0.0, step=100.0,
                                 format="%.2f", key="inc_amount")
    inc_mode   = st.selectbox("Frequency",
                              ["monthly","daily","weekly","biweekly","yearly"],
                              key="inc_mode")
    inc_date   = st.date_input("Date", value=date.today(), key="inc_date")
    inc_note   = st.text_input("Note", placeholder="Invoice #, project…", key="inc_note")
    inc_status = st.radio("Status", ["Received ✅", "Pending ⏳"],
                          horizontal=True, key="inc_status")
    if st.button("Add Income", key="btn_income"):
        if inc_amount <= 0:
            st.error("Amount must be > 0")
        else:
            add_entry("income", {
                "source": inc_source.strip() or "—",
                "amount": inc_amount,
                "mode":   inc_mode,
                "monthly_equiv": monthly_income_projection(inc_amount, inc_mode)["monthly"],
                "note":   inc_note.strip() or None,
                "status": "received" if "Received" in inc_status else "pending",
                "date":   inc_date.isoformat(),
            })
            bi2 = calculate_balance(load_data())
            if bi2["total_income"] > bi2["total_expenses"]:
                award_points(20)
                st.success("✓ Income saved! +10 pts · 🎉 Under-budget bonus +20 pts")
            else:
                st.success("✓ Income saved! +10 pts")
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

with col_exp:
    st.markdown("<div class='form-card'>"
                "<div class='form-card-title'>💸 Add Expense</div>", unsafe_allow_html=True)
    exp_amount   = st.number_input("Amount (₹)", min_value=0.0, step=50.0,
                                   format="%.2f", key="exp_amount")
    cats         = data.get("categories", ["Other"])
    exp_category = st.selectbox("Category", cats, key="exp_category")
    exp_date     = st.date_input("Date", value=date.today(), key="exp_date")
    exp_note     = st.text_input("Note", placeholder="Merchant, reason…", key="exp_note")
    exp_hidden   = st.toggle("Hidden entry 🔒", key="exp_hidden",
                             help="Stored privately — excluded from shared/export views.")
    exp_status   = st.radio("Status", ["Paid ✅", "Pending ⏳"],
                            horizontal=True, key="exp_status")
    if st.button("Add Expense", key="btn_expense"):
        if exp_amount <= 0:
            st.error("Amount must be > 0")
        else:
            add_entry("expenses", {
                "amount":   exp_amount,
                "category": exp_category,
                "note":     exp_note.strip() or None,
                "hidden":   exp_hidden,
                "status":   "paid" if "Paid" in exp_status else "pending",
                "date":     exp_date.isoformat(),
            })
            st.success("✓ Expense recorded! +10 pts")
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PENDING PAYMENTS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Pending Payments</div>", unsafe_allow_html=True)

pt1, pt2 = st.tabs(["💸 I owe  (Payables)", "💵 Owed to me  (Receivables)"])

with pt1:
    pend_exp = [(i, e) for i, e in enumerate(data.get("expenses", []))
                if str(e.get("status","")).lower() == "pending"]
    if not pend_exp:
        st.caption("No pending expenses ✅")
    else:
        for idx, (i, e) in enumerate(pend_exp):
            r1, r2, r3 = st.columns([5, 1, 1])
            with r1:
                st.markdown(
                    f"<div class='entry-line'>"
                    f"<span class='entry-badge-y'>PENDING</span>"
                    f"<span class='entry-amt'>₹{float(e.get('amount',0)):,.2f}</span>"
                    f"<span class='entry-meta'>{e.get('category','—')} · {e.get('note') or '—'} "
                    f"· {(e.get('date',''))[:10]}</span>"
                    f"</div>", unsafe_allow_html=True)
            with r2:
                if st.button("✅ Paid", key=f"pe_{i}_{idx}"):
                    mark_paid("expenses", i); award_points(5); st.rerun()
            with r3:
                if st.button("🗑", key=f"del_pe_{i}_{idx}", help="Delete entry"):
                    delete_entry("expenses", i); st.rerun()

with pt2:
    pend_inc = [(i, e) for i, e in enumerate(data.get("income", []))
                if str(e.get("status","")).lower() == "pending"]
    if not pend_inc:
        st.caption("Nothing pending ✅")
    else:
        for idx, (i, e) in enumerate(pend_inc):
            r1, r2, r3 = st.columns([5, 1, 1])
            with r1:
                st.markdown(
                    f"<div class='entry-line'>"
                    f"<span class='entry-badge-y'>PENDING</span>"
                    f"<span class='entry-amt'>₹{float(e.get('amount',0)):,.2f}</span>"
                    f"<span class='entry-meta'>{e.get('source','—')} · {e.get('note') or '—'} "
                    f"· {(e.get('date',''))[:10]}</span>"
                    f"</div>", unsafe_allow_html=True)
            with r2:
                if st.button("✅ Got it", key=f"pi_{i}_{idx}"):
                    mark_paid("income", i); award_points(5); st.rerun()
            with r3:
                if st.button("🗑", key=f"del_pi_{i}_{idx}", help="Delete entry"):
                    delete_entry("income", i); st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# RECENT ENTRIES
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Recent Entries</div>", unsafe_allow_html=True)

ri, re_ = st.tabs(["Income (last 10)", "Expenses (last 10)"])

with ri:
    all_inc_sorted = sorted(
        enumerate(data.get("income", [])),
        key=lambda x: x[1].get("date", ""), reverse=True
    )[:10]
    if not all_inc_sorted:
        st.caption("No income entries yet.")
    for orig_i, r in all_inc_sorted:
        badge = "entry-badge-g" if r.get("status") == "received" else "entry-badge-y"
        label = "RECEIVED" if r.get("status") == "received" else "PENDING"
        c1, c2 = st.columns([9, 1])
        with c1:
            st.markdown(
                f"<div class='entry-line'>"
                f"<span class='{badge}'>{label}</span>"
                f"<span class='entry-amt'>₹{float(r.get('amount',0)):,.2f}</span>"
                f"<span class='entry-meta'>{r.get('source') or r.get('mode','—')} · {r.get('date','')[:10]}</span>"
                f"</div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑", key=f"del_inc_{orig_i}", help="Delete this entry"):
                delete_entry("income", orig_i)
                st.rerun()

with re_:
    all_exp_sorted = sorted(
        enumerate(data.get("expenses", [])),
        key=lambda x: x[1].get("date", ""), reverse=True
    )[:10]
    if not all_exp_sorted:
        st.caption("No expense entries yet.")
    for orig_i, r in all_exp_sorted:
        badge = "entry-badge-r" if r.get("status") == "paid" else "entry-badge-y"
        label = "PAID" if r.get("status") == "paid" else "PENDING"
        lock  = " 🔒" if r.get("hidden") else ""
        c1, c2 = st.columns([9, 1])
        with c1:
            st.markdown(
                f"<div class='entry-line'>"
                f"<span class='{badge}'>{label}</span>"
                f"<span class='entry-amt'>₹{float(r.get('amount',0)):,.2f}</span>"
                f"<span class='entry-meta'>{r.get('category','Other')}{lock} · {r.get('note') or '—'} · {r.get('date','')[:10]}</span>"
                f"</div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑", key=f"del_exp_{orig_i}", help="Delete this entry"):
                delete_entry("expenses", orig_i)
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# MANAGE: Goals · Bonuses · Categories
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Manage</div>", unsafe_allow_html=True)

tab_rw, tab_bn, tab_cat = st.tabs(["🎯 Goals", "🎁 Bonuses", "🏷️ Categories"])

# ── Goals (formerly Rewards) ──────────────────────────────────────────────────
with tab_rw:
    ra, rb = st.columns([1, 1], gap="large")
    with ra:
        st.markdown("**Add a goal**")
        rw_name       = st.text_input("Goal name", placeholder="Mechanical keyboard, Goa trip…", key="rw_name")
        rw_cost       = st.number_input("Cost (₹)", min_value=0.0, step=100.0, key="rw_cost")
        rw_target_date= st.date_input("Target date", value=date.today(), min_value=date.today(), key="rw_target")
        rw_link       = st.text_input("Product / reference link", placeholder="https://…", key="rw_link")
        if st.button("Save Goal", key="btn_rw"):
            if not rw_name.strip() or rw_cost <= 0:
                st.error("Name and cost required.")
            else:
                add_entry("rewards", {
                    "name":        rw_name.strip(),
                    "cost":        rw_cost,
                    "target_date": rw_target_date.isoformat(),
                    "link":        rw_link.strip() or None,
                    "date":        datetime.now().isoformat(timespec="seconds"),
                })
                st.success("✓ Goal saved!")
                st.rerun()

    with rb:
        rewards = data.get("rewards", [])
        if not rewards:
            st.caption("Add a goal to start tracking.")
        else:
            today_d = date.today()
            for gi, rw in enumerate(rewards):
                cost        = float(rw.get("cost", 0) or 0)
                pct         = max(0.0, min(100.0, (net / cost * 100) if cost > 0 else 100))
                est         = estimate_days_for_reward(net, daily_saving, cost)

                # ── Date-based calculations ──
                target_str  = rw.get("target_date")
                days_left   = None
                need_per_day= None
                if target_str:
                    try:
                        target_d  = date.fromisoformat(target_str)
                        days_left = max(0, (target_d - today_d).days)
                        shortfall = max(0.0, cost - net)
                        need_per_day = (shortfall / days_left) if days_left > 0 else shortfall
                    except:
                        days_left = None

                # ETA with current saving rate
                if est["affordable"]:
                    eta_cur = "✅ Affordable now!"
                elif est["reachable"]:
                    eta_cur = f"~{est['days']} days at current rate"
                else:
                    eta_cur = "⚠️ Increase daily savings"

                with st.expander(f"🎯 {rw['name']}  —  ₹{cost:,.0f}"):
                    st.progress(pct / 100, text=f"{pct:.0f}% of goal saved")

                    g1, g2 = st.columns(2)
                    with g1:
                        st.markdown(
                            f"<div class='sb-block' style='margin-bottom:6px'>"
                            f"<div class='sb-lbl'>Current saving rate</div>"
                            f"<div class='sb-val c-violet'>₹{daily_saving:,.0f}/day</div>"
                            f"</div>"
                            f"<div class='sb-block'>"
                            f"<div class='sb-lbl'>Days at this rate</div>"
                            f"<div class='sb-val c-teal'>{eta_cur}</div>"
                            f"</div>", unsafe_allow_html=True)
                    with g2:
                        if days_left is not None:
                            dpd_color = "c-green" if (need_per_day or 0) <= daily_saving else "c-amber"
                            st.markdown(
                                f"<div class='sb-block' style='margin-bottom:6px'>"
                                f"<div class='sb-lbl'>Target date</div>"
                                f"<div class='sb-val'>{target_str}  ({days_left}d left)</div>"
                                f"</div>"
                                f"<div class='sb-block'>"
                                f"<div class='sb-lbl'>Must save per day to hit target</div>"
                                f"<div class='sb-val {dpd_color}'>₹{need_per_day:,.0f}/day</div>"
                                f"</div>", unsafe_allow_html=True)
                        else:
                            st.caption("No target date set.")

                    if rw.get("link"):
                        st.markdown(f"[View product / link ↗]({rw['link']})")

                    # Delete button inside expander
                    if st.button(f"🗑 Remove this goal", key=f"del_goal_{gi}"):
                        delete_entry("rewards", gi)
                        st.rerun()

            # Summary table
            st.markdown("<hr class='vdivider'>", unsafe_allow_html=True)
            rows_est = []
            today_d2 = date.today()
            for rw in rewards:
                cost     = float(rw.get("cost", 0) or 0)
                est      = estimate_days_for_reward(net, daily_saving, cost)
                shortfall= max(0.0, cost - net)
                tgt      = rw.get("target_date", "—")
                try:
                    dl   = max(0, (date.fromisoformat(tgt) - today_d2).days)
                    npd  = f"₹{shortfall/dl:,.0f}/day" if dl > 0 else "Past due"
                except:
                    dl, npd = "—", "—"
                rows_est.append({
                    "Goal":              rw["name"],
                    "Cost":              f"₹{cost:,.0f}",
                    "Shortfall":         f"₹{shortfall:,.0f}",
                    "Target Date":       tgt,
                    "Days Left":         dl,
                    "Need/Day for Target": npd,
                    "Days @ Current Rate": ("Ready ✅" if est["affordable"]
                                            else (f"~{est['days']}d" if est["reachable"]
                                                  else "⚠️")),
                })
            st.dataframe(rows_est, use_container_width=True, hide_index=True)

# ── Bonuses ───────────────────────────────────────────────────────────────────
with tab_bn:
    ba, bb = st.columns([1, 1], gap="large")
    with ba:
        st.markdown("**Log a bonus**")
        bn_amt    = st.number_input("Amount (₹)", min_value=0.0, step=100.0, key="bn_amt")
        bn_note   = st.text_input("Note", placeholder="Referral, cashback…", key="bn_note")
        bn_hidden = st.toggle("Hidden 🔒",
                              help="Stored privately. Not in open lists or Excel export.",
                              key="bn_hidden")
        if st.button("Add Bonus", key="btn_bn"):
            if bn_amt <= 0:
                st.error("Amount must be > 0")
            else:
                add_entry("bonuses", {
                    "amount": bn_amt, "note": bn_note.strip() or None,
                    "hidden": bn_hidden,
                    "date":   datetime.now().isoformat(timespec="seconds"),
                })
                st.success("✓ Bonus logged!" + (" (hidden)" if bn_hidden else " +10 pts"))
                st.rerun()

    with bb:
        st.markdown("**Open bonuses**")
        open_b_indexed = [(i, b) for i, b in enumerate(data.get("bonuses", []))
                          if not b.get("hidden")]
        if not open_b_indexed:
            st.caption("No open bonuses yet.")
        for bi_, (orig_i, b) in enumerate(open_b_indexed):
            c1, c2 = st.columns([9, 1])
            with c1:
                st.markdown(
                    f"<div class='entry-line'>"
                    f"<span class='entry-badge-g'>BONUS</span>"
                    f"<span class='entry-amt c-amber'>₹{float(b.get('amount',0)):,.2f}</span>"
                    f"<span class='entry-meta'>{b.get('note') or '—'} · {(b.get('date') or b.get('created_at',''))[:10]}</span>"
                    f"</div>", unsafe_allow_html=True)
            with c2:
                if st.button("🗑", key=f"del_bon_{orig_i}", help="Delete bonus"):
                    delete_entry("bonuses", orig_i)
                    st.rerun()
        hc = sum(1 for b in data.get("bonuses", []) if b.get("hidden"))
        if hc:
            st.caption(f"🔒 {hc} hidden bonus{'es' if hc > 1 else ''} stored.")

# ── Categories ────────────────────────────────────────────────────────────────
with tab_cat:
    ca, cb = st.columns([1, 1], gap="large")
    with ca:
        st.markdown("**Add category**")
        new_cat = st.text_input("Name", placeholder="Subscriptions, Pet care…", key="new_cat")
        if st.button("Add Category", key="btn_cat"):
            if not new_cat.strip():
                st.error("Enter a name.")
            elif new_cat.strip() in data.get("categories", []):
                st.warning("Already exists.")
            else:
                add_entry("categories", new_cat.strip())
                st.success(f"✓ '{new_cat.strip()}' added!")
                st.rerun()
    with cb:
        st.markdown("**All categories**")
        cats = data.get("categories", [])
        if cats:
            for ci, cat in enumerate(cats):
                c1, c2 = st.columns([8, 1])
                with c1:
                    st.markdown(f"<span class='cat-tag'>{cat}</span>",
                                unsafe_allow_html=True)
                with c2:
                    if st.button("🗑", key=f"del_cat_{ci}", help=f"Remove {cat}"):
                        delete_entry("categories", ci)
                        st.rerun()
        else:
            st.caption("No categories yet.")

# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='sec-hdr'>Export</div>", unsafe_allow_html=True)

ex1, ex2 = st.columns([1, 3], gap="medium")

with ex1:
    if st.button("📥 Export to Excel", key="btn_export"):
        try:
            with st.spinner("Building report…"):
                path = export_to_excel()
            with open(path, "rb") as fh:
                st.download_button(
                    "⬇️ Download expense_report.xlsx", data=fh,
                    file_name="expense_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl",
                )
        except Exception as e:
            st.error(f"Export failed: {e}")

with ex2:
    st.markdown("""
    <div class="export-info">
        Report contains <strong>4 sheets</strong>:
        <em>Overview · Expenses · Income &amp; Open Bonuses · Category Summary</em>
        (monthly + yearly pivot).<br>
        🔒 Hidden entries &amp; bonuses are excluded from the export.
    </div>""", unsafe_allow_html=True)
