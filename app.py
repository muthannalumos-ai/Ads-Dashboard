"""
Ads Performance Dashboard — 2025-2026 | CEO Report
Fiscal Year: July → June
"""
import re, warnings
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Ads Performance Dashboard", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# ── CSS: Bright Theme ─────────────────────────────────────────────────────────
st.markdown("""<style>
.stApp { background-color: #f0f4f9 !important; }
.main .block-container { padding-top: 1rem; }

/* Sidebar — deep blue */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a237e 0%, #1565c0 100%) !important;
}
[data-testid="stSidebar"] * { color: #ffffff !important; }
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.18) !important;
    border: 1px solid rgba(255,255,255,0.35) !important;
    border-radius: 8px; font-weight: 600; width: 100%;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.28) !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.2); }

/* KPI cards */
.kpi-card {
    background: #ffffff;
    border: 1.5px solid #d8e2f0;
    border-radius: 12px;
    padding: 13px 8px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(26,35,126,0.08);
}
.kpi-label {
    font-size: 9px; font-weight: 700; letter-spacing: 0.6px;
    text-transform: uppercase; color: #5c6780 !important;
    margin-bottom: 4px; white-space: nowrap;
}
.kpi-value {
    font-size: 19px; font-weight: 800; color: #1a237e !important;
    line-height: 1.1; white-space: nowrap; overflow: hidden;
}
.kpi-sub { font-size: 9px; color: #90a4ae !important; margin-top: 3px; }

/* Section headers */
.section-header {
    font-size: 12px; font-weight: 700; letter-spacing: 0.6px;
    color: #37474f !important; text-transform: uppercase;
    border-bottom: 2px solid #d8e2f0;
    padding-bottom: 6px; margin-bottom: 12px; margin-top: 6px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff; border-radius: 10px; padding: 4px;
    border: 1px solid #d8e2f0; gap: 2px;
}
.stTabs [data-baseweb="tab"] {
    color: #37474f !important; font-weight: 600;
    border-radius: 7px; padding: 9px 24px;
}
.stTabs [aria-selected="true"] {
    background: #1a237e !important; color: #ffffff !important;
}

/* Tables */
[data-testid="stDataFrame"] {
    border: 1px solid #d8e2f0; border-radius: 10px; overflow: hidden;
}
</style>""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
GADS_PATH = BASE_DIR / "data" / "Google ads spend .xlsx"
FB_PATH   = BASE_DIR / "data" / "fb ad spend .xlsx"
CONV_PATH = BASE_DIR / "data" / "Meta _ Ads leads generated .xlsx"

FISCAL_MONTHS = ["July","August","September","October","November","December",
                 "January","February","March","April","May","June"]
MONTH_ORDER   = {m: i for i, m in enumerate(FISCAL_MONTHS)}
PARENT_KW     = ["parent","parents","worksheet","worksheets","amazon","amz","mdt"]
BLUE_MAP      = {"Google Ads": "#1a237e", "Facebook Ads": "#1877f2"}

# ── Chart base layout (bright) ────────────────────────────────────────────────
_BASE = dict(
    paper_bgcolor="#ffffff", plot_bgcolor="#f7f9fc",
    font=dict(family="Arial", size=12, color="#1a1d23"),
    margin=dict(t=55, b=45, l=40, r=20),
    xaxis=dict(gridcolor="#eaeff6", zeroline=False, linecolor="#d8e2f0",
               tickfont=dict(color="#344563", size=11)),
    yaxis=dict(gridcolor="#eaeff6", zeroline=False, linecolor="#d8e2f0",
               tickfont=dict(color="#344563", size=11)),
    legend=dict(bgcolor="rgba(255,255,255,0.95)", bordercolor="#d8e2f0",
                borderwidth=1, font=dict(color="#1a1d23", size=11)),
    title_font=dict(color="#1a237e", size=14, family="Arial"),
)

def al(fig, **kw):
    """Apply layout safely — merges nested dicts, no duplicate key errors."""
    lay = {k: dict(v) if isinstance(v, dict) else v for k, v in _BASE.items()}
    for k, v in kw.items():
        if k in lay and isinstance(lay[k], dict) and isinstance(v, dict):
            lay[k].update(v)
        else:
            lay[k] = v
    fig.update_layout(**lay)

# ── Categorisation helpers ────────────────────────────────────────────────────
def _tok(name): return re.split(r"[\s\-_:.,()#/]+", str(name).lower())

def is_parent(name):
    nl = str(name).lower()
    return any(kw in _tok(name) or kw in nl for kw in PARENT_KW)

def has_form(name):
    return bool(re.search(r"leads?\s+form|\bform\b", str(name).lower()))

def school_type(name):
    nl = str(name).lower()
    if has_form(nl):                                                   return "Lead Form"
    if "placement" in nl or ("awareness" in nl and "school" in nl):   return "Awareness"
    if "type" in nl:                                                   return "Lead Gen (Type)"
    if "pmax" in nl or "performance max" in nl:                        return "PMax"
    if "retarget" in nl or "remarketing" in nl:                        return "Retargeting"
    if re.search(r"conf[e:]|conference|\bncsc\b", nl):                 return "Conference"
    if "librar" in nl:                                                  return "Library"
    return "Other"

def parent_type(name):
    nl = str(name).lower()
    if has_form(nl):                             return "Lead Form"
    if "amazon" in nl or "amz" in nl:            return "Amazon/AMZ"
    if "mdt" in nl:                              return "MDT"
    if "worksheet" in nl:                        return "Worksheet"
    if "retarget" in nl or "remarketing" in nl:  return "Retargeting"
    return "Other"

def other_sub(name):
    nl = str(name).lower()
    if "brand" in nl:                         return "Brand"
    if "competitor" in nl or "compet" in nl:  return "Competitor"
    if "district" in nl:                      return "District"
    if "blog" in nl:                          return "Blog"
    if "raise" in nl:                         return "RAISE Act"
    if "video" in nl:                         return "Video"
    if "display" in nl:                       return "Display"
    if "youtube" in nl or " yt " in nl:       return "YouTube"
    if "search" in nl:                        return "Search"
    return "Misc"

def classify(name):
    if is_parent(name):
        aud, ct = "Parent", parent_type(name)
    else:
        aud, ct = "School", school_type(name)
    return aud, ct, (other_sub(name) if ct == "Other" else "")

# ── Data loaders ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_gads():
    df = pd.read_excel(GADS_PATH, sheet_name="Sheet0", header=2)
    df.columns = df.columns.str.strip()
    df = df[~df["Campaign status"].astype(str).str.startswith("Total", na=False)]
    df = df.dropna(subset=["Campaign"])
    df = df[df["Campaign"].astype(str).str.strip() != ""]
    df["Cost"]     = pd.to_numeric(df["Cost"],     errors="coerce").fillna(0)
    df["Impr."]    = pd.to_numeric(df["Impr."],    errors="coerce").fillna(0)
    df["Clicks"]   = pd.to_numeric(df["Clicks"],   errors="coerce").fillna(0)
    df["Avg. CPC"] = pd.to_numeric(df["Avg. CPC"], errors="coerce").fillna(0)
    df["CTR"]      = pd.to_numeric(df["CTR"],      errors="coerce").fillna(0) * 100
    df = df[df["Cost"] > 0].copy()
    rows = [classify(r["Campaign"]) for _, r in df.iterrows()]
    df["Audience"], df["Campaign Type"], df["Other Sub"] = zip(*rows)
    df["Platform"] = "Google Ads"
    return df.reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_fb():
    df = pd.read_excel(FB_PATH, sheet_name="Worksheet")
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Campaign name"])
    df = df[df["Campaign name"].astype(str).str.strip() != ""]
    df["Amount spent (USD)"] = pd.to_numeric(df["Amount spent (USD)"], errors="coerce").fillna(0)
    df = df[df["Amount spent (USD)"] > 0].copy()
    df["Impressions"]      = pd.to_numeric(df["Impressions"],      errors="coerce").fillna(0)
    df["Results"]          = pd.to_numeric(df["Results"],          errors="coerce").fillna(0)
    df["Cost per results"] = pd.to_numeric(df["Cost per results"], errors="coerce").fillna(0)
    rows = [classify(r["Campaign name"]) for _, r in df.iterrows()]
    df["Audience"], df["Campaign Type"], df["Other Sub"] = zip(*rows)
    df["Platform"] = "Facebook Ads"
    return df.reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_conv():
    raw = pd.read_excel(CONV_PATH, sheet_name="QL LEADS", header=None)

    def _c(df, mc="Month"):
        df[mc] = df[mc].astype(str).str.strip()
        df["Month_Order"] = df[mc].map(MONTH_ORDER)
        return df.sort_values("Month_Order").reset_index(drop=True)

    ov = raw.iloc[2:13,[1,2,3,4,5,6]].copy()
    ov.columns = ["Month","FB_HL","GA_HL","Total_HL","QL","HL_QL_pct"]
    ov = ov.dropna(subset=["Month"])
    for c in ov.columns[1:]: ov[c] = pd.to_numeric(ov[c], errors="coerce").fillna(0)
    if ov["HL_QL_pct"].max() < 2: ov["HL_QL_pct"] *= 100
    ov = _c(ov)

    aw = raw.iloc[20:30,[0,1,2,3]].copy(); aw.columns = ["Month","FB","GA","Total"]
    aw = aw.dropna(subset=["Month"])
    for c in ["FB","GA","Total"]: aw[c] = pd.to_numeric(aw[c], errors="coerce").fillna(0)
    aw = _c(aw)

    rt = raw.iloc[20:30,[4,5,6,7]].copy(); rt.columns = ["Month","FB","GA","Total"]
    rt = rt.dropna(subset=["Month"])
    for c in ["FB","GA","Total"]: rt[c] = pd.to_numeric(rt[c], errors="coerce").fillna(0)
    rt = _c(rt)

    fbt = raw.iloc[35:45,[0,1,2,3]].copy(); fbt.columns = ["Month","HL","QL","HL_QL_pct"]
    fbt = fbt.dropna(subset=["Month"])
    for c in ["HL","QL","HL_QL_pct"]: fbt[c] = pd.to_numeric(fbt[c], errors="coerce").fillna(0)
    if fbt["HL_QL_pct"].max() < 2: fbt["HL_QL_pct"] *= 100
    fbt = _c(fbt)

    fbf = raw.iloc[35:45,[4,5,6,7]].copy(); fbf.columns = ["Month","HL","QL","HL_QL_pct"]
    fbf = fbf.dropna(subset=["Month"])
    for c in ["HL","QL","HL_QL_pct"]: fbf[c] = pd.to_numeric(fbf[c], errors="coerce").fillna(0)
    if not fbf.empty and fbf["HL_QL_pct"].max() < 2: fbf["HL_QL_pct"] *= 100
    fbf = _c(fbf)

    return dict(overall=ov, awareness=aw, retarget=rt, fb_type=fbt, fb_form=fbf)

def mf(df, months, col="Month"):
    return df[df[col].str.strip().isin(months)]

def kpi(label, value, sub=""):
    return (f'<div class="kpi-card">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "") +
            '</div>')

# ══════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📊 Ads Dashboard")
    st.markdown("<p style='font-size:11px;opacity:0.7;margin-top:-8px;'>FY 2025–2026 | Jul → Jun</p>",
                unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    st.markdown("### Filters")
    platform_opt = st.multiselect("Platform",["Google Ads","Facebook Ads"],
                                  default=["Google Ads","Facebook Ads"])
    audience_opt = st.multiselect("Audience",["School","Parent"],
                                  default=["School","Parent"])
    month_opt    = st.multiselect("Month", FISCAL_MONTHS, default=FISCAL_MONTHS,
                                  help="Filters conversion (HL/QL) metrics. Spend = full period.")
    st.markdown("---")
    st.caption("Spend: full fiscal period\nConversions: Meta Ads Leads sheet")

# ── Load ──────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    gads = load_gads(); fb = load_fb(); conv = load_conv()

# ── Platform + audience filter ────────────────────────────────
gads_f = gads[gads["Audience"].isin(audience_opt)] if "Google Ads"    in platform_opt else gads.iloc[:0]
fb_f   = fb[fb["Audience"].isin(audience_opt)]     if "Facebook Ads"  in platform_opt else fb.iloc[:0]

# ── Spend / clicks (full period, no monthly breakdown available) ──
ga_spend  = gads_f["Cost"].sum();            fb_spend  = fb_f["Amount spent (USD)"].sum()
ga_clicks = gads_f["Clicks"].sum();          fb_res    = fb_f["Results"].sum()
ga_impr   = gads_f["Impr."].sum();           fb_impr   = fb_f["Impressions"].sum()
total_spend  = ga_spend  + fb_spend
total_clicks = ga_clicks + fb_res

# CPC — total spend / total clicks (correct, not mean of per-campaign CPC)
ga_cpc = ga_spend  / ga_clicks if ga_clicks > 0 else 0
fb_cpc = fb_spend  / fb_res    if fb_res    > 0 else 0
tot_cpc= total_spend / total_clicks if total_clicks > 0 else 0

# CTR — weighted impressions for GA; results/impressions for FB
ga_ctr = (gads_f["CTR"] * gads_f["Impr."]).sum() / ga_impr if ga_impr > 0 else 0
fb_ctr = fb_res / fb_impr * 100 if fb_impr > 0 else 0
tot_ctr= (ga_clicks + fb_res) / (ga_impr + fb_impr) * 100 if (ga_impr + fb_impr) > 0 else 0

# ── Conversions filtered by month + platform + audience ──────
# Uses the OVERALL sheet which has all 11 months (Jul–May) for both platforms.
# Per-type breakdowns (fb_type, awareness etc.) only go to April — using them
# causes May to silently return 0 even when selected, breaking the month filter.
def conv_totals(plats, auds, months):
    hl = ql = 0
    ov_ = mf(conv["overall"], months)
    if "School" in auds:
        both = "Facebook Ads" in plats and "Google Ads" in plats
        fb_only = "Facebook Ads" in plats and "Google Ads" not in plats
        ga_only = "Google Ads" in plats and "Facebook Ads" not in plats
        if both:
            hl = ov_["Total_HL"].sum()
            ql = ov_["QL"].sum()
        elif fb_only:
            hl = ov_["FB_HL"].sum()
            # FB QL = Total QL − GA QL (GA HL == GA QL by definition)
            ql = (ov_["QL"] - ov_["GA_HL"]).clip(lower=0).sum()
        elif ga_only:
            hl = ov_["GA_HL"].sum()
            ql = ov_["GA_HL"].sum()
    return hl, ql

hl_tot, ql_tot = conv_totals(platform_opt, audience_opt, month_opt)
cost_hl  = total_spend / hl_tot if hl_tot > 0 else 0
hl_ql_pct= ql_tot / hl_tot * 100 if hl_tot > 0 else 0
conv_rate= hl_tot / total_clicks * 100 if total_clicks > 0 else 0

# ── Pre-filter conv dfs (used in all tabs) ────────────────────
ov_f  = mf(conv["overall"],   month_opt)
aw_f  = mf(conv["awareness"], month_opt)
rt_f  = mf(conv["retarget"],  month_opt)
fbt_f = mf(conv["fb_type"],   month_opt)
fbf_f = mf(conv["fb_form"],   month_opt)

# ══════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════
st.markdown(
    "<h1 style='color:#1a237e;font-size:24px;font-weight:800;margin-bottom:2px;'>"
    "📊 Ads Performance Dashboard</h1>"
    "<p style='color:#607d8b;font-size:12px;margin-top:0;'>"
    "FY 2025–2026 &nbsp;|&nbsp; Jul 1 2025 – May 28 2026 &nbsp;|&nbsp; "
    "💡 Spend metrics = full fiscal period &nbsp;|&nbsp; "
    "Conversion metrics respond to month filter</p>",
    unsafe_allow_html=True)
st.markdown("---")

# ── KPI — Row 1: Conversion metrics (respond to ALL filters incl. month) ──
mo_label = (f"📅 {len(month_opt)} of 11 months"
            if len(month_opt) < 11 else "📅 All 11 Months")
st.markdown(
    f"<p style='font-size:11px;color:#1565c0;font-weight:700;margin-bottom:6px;'>"
    f"CONVERSION METRICS &nbsp;—&nbsp; {mo_label} &nbsp;|&nbsp; "
    f"Platform: {', '.join(platform_opt) if platform_opt else 'None'} &nbsp;|&nbsp; "
    f"Audience: {', '.join(audience_opt) if audience_opt else 'None'}</p>",
    unsafe_allow_html=True)
conv_kpis = [
    ("Hot Leads (HL)",  f"{hl_tot:,.0f}",         mo_label),
    ("Qualified Leads", f"{ql_tot:,.0f}",          "HL→QL confirmed"),
    ("HL → QL%",        f"{hl_ql_pct:.1f}%",       f"QL: {ql_tot:,.0f}"),
    ("Cost / Hot Lead", f"${cost_hl:.0f}",         "Full Spend ÷ HL"),
    ("Conv. Rate",      f"{conv_rate:.2f}%",        "HL ÷ Clicks"),
]
c_cols = st.columns(5)
for col, (lbl, val, sub) in zip(c_cols, conv_kpis):
    col.markdown(kpi(lbl, val, sub), unsafe_allow_html=True)

st.markdown("<p style='font-size:11px;color:#78909c;font-weight:700;margin-bottom:6px;margin-top:14px;'>"
            "SPEND METRICS &nbsp;—&nbsp; 📌 Full Fiscal Period (no monthly breakdown in source files)</p>",
            unsafe_allow_html=True)
spend_kpis = [
    ("Total Spend",    f"${total_spend:,.0f}",      "Google + Facebook"),
    ("Clicks / Results",f"{total_clicks:,.0f}",     "GA Clicks + FB Results"),
    ("Avg CPC",        f"${tot_cpc:.2f}",            f"GA:${ga_cpc:.2f} | FB:${fb_cpc:.2f}"),
    ("Avg CTR",        f"{tot_ctr:.2f}%",            f"GA:{ga_ctr:.2f}% | FB:{fb_ctr:.2f}%"),
]
s_cols = st.columns(4)
for col, (lbl, val, sub) in zip(s_cols, spend_kpis):
    col.markdown(kpi(lbl, val, sub), unsafe_allow_html=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["📈 Overview", "🎯 Campaign Performance", "📅 Monthly Trends"])

# ─────────────────────────────────────────────────────────────
# TAB 1 — Overview
# ─────────────────────────────────────────────────────────────
with tab1:
    # Row 1: three charts
    c1, c2, c3 = st.columns([2.5, 1.3, 1.3])

    # Spend by Campaign Type
    with c1:
        st.markdown("<div class='section-header'>Spend by Campaign Type</div>", unsafe_allow_html=True)
        ga_ct = gads_f.groupby("Campaign Type", as_index=False)["Cost"].sum().rename(columns={"Cost":"Spend"})
        ga_ct["Platform"] = "Google Ads"
        fb_ct = fb_f.groupby("Campaign Type", as_index=False)["Amount spent (USD)"].sum().rename(columns={"Amount spent (USD)":"Spend"})
        fb_ct["Platform"] = "Facebook Ads"
        sct = pd.concat([ga_ct, fb_ct], ignore_index=True)
        if not sct.empty:
            order = sct.groupby("Campaign Type")["Spend"].sum().sort_values(ascending=False).index.tolist()
            with st.container(border=True):
                fig = px.bar(sct, x="Campaign Type", y="Spend",
                             color="Platform", barmode="stack",
                             category_orders={"Campaign Type": order},
                             color_discrete_map=BLUE_MAP, height=420)
                al(fig, title="Spend by Campaign Type",
                   yaxis=dict(title="Spend ($)", tickprefix="$", tickformat=",.0f"),
                   xaxis=dict(title=""))
                # Total labels on top of each stack
                totals = sct.groupby("Campaign Type")["Spend"].sum().reindex(order)
                fig.add_trace(go.Scatter(
                    x=order, y=totals.values,
                    mode="text",
                    text=[f"${v:,.0f}" for v in totals.values],
                    textposition="top center",
                    textfont=dict(size=9, color="#1a237e"),
                    showlegend=False,
                ))
                st.plotly_chart(fig, use_container_width=True)

    # Platform donut
    with c2:
        st.markdown("<div class='section-header'>By Platform</div>", unsafe_allow_html=True)
        pdf = pd.DataFrame({"Platform":["Google Ads","Facebook Ads"],"Spend":[ga_spend,fb_spend]})
        pdf = pdf[pdf["Spend"] > 0]
        if not pdf.empty:
            with st.container(border=True):
                fig2 = go.Figure(go.Pie(
                    labels=pdf["Platform"], values=pdf["Spend"], hole=0.58,
                    marker_colors=["#1a237e","#1877f2"],
                    textinfo="percent+label", textfont=dict(size=11, color="#1a1d23"),
                    hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>",
                ))
                al(fig2, height=420, showlegend=False,
                   xaxis=dict(visible=False), yaxis=dict(visible=False),
                   annotations=[dict(text=f"<b>${total_spend:,.0f}</b>",
                                     x=0.5, y=0.5, showarrow=False,
                                     font=dict(size=12, color="#1a237e"))])
                st.plotly_chart(fig2, use_container_width=True)

    # Audience donut
    with c3:
        st.markdown("<div class='section-header'>By Audience</div>", unsafe_allow_html=True)
        adf = (gads_f.groupby("Audience")["Cost"].sum()
               .add(fb_f.groupby("Audience")["Amount spent (USD)"].sum(), fill_value=0)
               ).reset_index()
        adf.columns = ["Audience","Spend"]; adf = adf[adf["Spend"] > 0]
        if not adf.empty:
            with st.container(border=True):
                fig3 = go.Figure(go.Pie(
                    labels=adf["Audience"], values=adf["Spend"], hole=0.58,
                    marker_colors=["#0288d1","#e65100"],
                    textinfo="percent+label", textfont=dict(size=11, color="#1a1d23"),
                    hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>",
                ))
                al(fig3, height=420, showlegend=False,
                   xaxis=dict(visible=False), yaxis=dict(visible=False))
                st.plotly_chart(fig3, use_container_width=True)

    # Row 2: HL by Platform (monthly)
    st.markdown("<div class='section-header'>Hot Leads by Platform — Monthly</div>", unsafe_allow_html=True)
    if not ov_f.empty:
        hl_traces = []
        if "Facebook Ads" in platform_opt: hl_traces.append(("FB_HL","Facebook Ads","#1877f2"))
        if "Google Ads"    in platform_opt: hl_traces.append(("GA_HL","Google Ads","#1a237e"))
        if hl_traces:
            with st.container(border=True):
                fig_hl = go.Figure()
                for col, name, color in hl_traces:
                    fig_hl.add_trace(go.Bar(
                        x=ov_f["Month"], y=ov_f[col], name=name,
                        marker_color=color,
                        text=ov_f[col].map(lambda v: f"{int(v):,}"),
                        textposition="outside", textfont=dict(size=9, color="#344563"),
                        hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>",
                    ))
                al(fig_hl, title="Monthly Hot Leads by Platform", barmode="group", height=400,
                   yaxis=dict(title="Hot Leads"),
                   legend=dict(orientation="h", y=1.07, x=0))
                st.plotly_chart(fig_hl, use_container_width=True)

    # Platform summary table (fixed CPC + CTR)
    st.markdown("<div class='section-header'>Platform Summary</div>", unsafe_allow_html=True)
    summary = []
    if "Google Ads" in platform_opt:
        summary.append({"Platform":"Google Ads", "Spend ($)": f"${ga_spend:,.2f}",
                         "Impressions":f"{ga_impr:,.0f}", "Clicks":f"{ga_clicks:,.0f}",
                         "Avg CPC ($)":f"${ga_cpc:.2f}", "Avg CTR (%)":f"{ga_ctr:.2f}%"})
    if "Facebook Ads" in platform_opt:
        summary.append({"Platform":"Facebook Ads", "Spend ($)":f"${fb_spend:,.2f}",
                         "Impressions":f"{fb_impr:,.0f}", "Clicks":f"{fb_res:,.0f}",
                         "Avg CPC ($)":f"${fb_cpc:.2f}",   # fixed: sum(spend)/sum(results)
                         "Avg CTR (%)":f"{fb_ctr:.2f}%"})  # fixed: sum(results)/sum(impressions)
    if summary:
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────
# TAB 2 — Campaign Performance
# ─────────────────────────────────────────────────────────────
with tab2:
    # Build per-type summary
    def breakdown():
        rows = []
        for df, plat, sc, cc, ic in [
            (gads_f,"Google Ads","Cost","Clicks","Impr."),
            (fb_f,"Facebook Ads","Amount spent (USD)","Results","Impressions"),
        ]:
            if df.empty: continue
            for (aud, ct), g in df.groupby(["Audience","Campaign Type"]):
                sp=g[sc].sum(); cl=g[cc].sum(); im=g[ic].sum() if ic in g else 0
                cpc=sp/cl if cl>0 else 0
                ctr=(g["CTR"]*g["Impr."]).sum()/im if plat=="Google Ads" and im>0 else (cl/im*100 if im>0 else 0)
                rows.append({"Platform":plat,"Audience":aud,"Campaign Type":ct,
                             "# Camps":len(g),"Spend ($)":sp,"Clicks":cl,"CPC ($)":cpc,"CTR (%)":ctr})
        return pd.DataFrame(rows)

    bdf = breakdown()

    if not bdf.empty:
        col1, col2 = st.columns(2)
        with col1:
            with st.container(border=True):
                sg = bdf.groupby("Campaign Type")["Spend ($)"].sum().sort_values(ascending=False).reset_index()
                sg["label"] = sg["Spend ($)"].map(lambda v: f"${v:,.0f}")
                fig_s = px.bar(sg, x="Campaign Type", y="Spend ($)", text="label",
                               color="Campaign Type",
                               color_discrete_sequence=["#1a237e","#1565c0","#0288d1","#00838f",
                                                         "#2e7d32","#f57f17","#e65100","#6a1b9a",
                                                         "#880e4f","#37474f","#b71c1c","#1b5e20"],
                               height=440, title="Spend by Campaign Type")
                al(fig_s, yaxis=dict(title="Spend ($)",tickprefix="$",tickformat=",.0f"),xaxis=dict(title=""))
                fig_s.update_traces(textposition="outside", textfont_size=9, showlegend=False)
                st.plotly_chart(fig_s, use_container_width=True)

        with col2:
            with st.container(border=True):
                cg = bdf.groupby("Campaign Type")["Clicks"].sum().sort_values(ascending=False).reset_index()
                cg["label"] = cg["Clicks"].map(lambda v: f"{v:,.0f}")
                fig_c = px.bar(cg, x="Campaign Type", y="Clicks", text="label",
                               color="Campaign Type",
                               color_discrete_sequence=["#1a237e","#1565c0","#0288d1","#00838f",
                                                         "#2e7d32","#f57f17","#e65100","#6a1b9a",
                                                         "#880e4f","#37474f","#b71c1c","#1b5e20"],
                               height=440, title="Clicks by Campaign Type")
                al(fig_c, yaxis=dict(title="Clicks"), xaxis=dict(title=""))
                fig_c.update_traces(textposition="outside", textfont_size=9, showlegend=False)
                st.plotly_chart(fig_c, use_container_width=True)

        # Table
        st.markdown("<div class='section-header'>Campaign Type Detail</div>", unsafe_allow_html=True)
        disp = bdf.sort_values("Spend ($)", ascending=False).copy()
        disp["Spend ($)"] = disp["Spend ($)"].map(lambda v: f"${v:,.2f}")
        disp["Clicks"]    = disp["Clicks"].map(lambda v: f"{v:,.0f}")
        disp["CPC ($)"]   = disp["CPC ($)"].map(lambda v: f"${v:.2f}")
        disp["CTR (%)"]   = disp["CTR (%)"].map(lambda v: f"{v:.2f}%")
        st.dataframe(disp, use_container_width=True, hide_index=True)

    # Hot Leads by Campaign Type (HL only)
    st.markdown("<div class='section-header'>Hot Leads (HL) by Campaign Type — Full Period</div>",
                unsafe_allow_html=True)
    CONV_AGG = [
        {"Platform":"Google Ads",  "Audience":"School","Campaign Type":"Awareness",      "Hot Leads":27},
        {"Platform":"Google Ads",  "Audience":"School","Campaign Type":"Retargeting",    "Hot Leads":8},
        {"Platform":"Google Ads",  "Audience":"School","Campaign Type":"Lead Gen (Type)","Hot Leads":3919},
        {"Platform":"Facebook Ads","Audience":"School","Campaign Type":"Awareness",      "Hot Leads":26},
        {"Platform":"Facebook Ads","Audience":"School","Campaign Type":"Retargeting",    "Hot Leads":16},
        {"Platform":"Facebook Ads","Audience":"School","Campaign Type":"Lead Gen (Type)","Hot Leads":9746},
        {"Platform":"Facebook Ads","Audience":"School","Campaign Type":"Lead Form",      "Hot Leads":1469},
    ]
    cdf = pd.DataFrame(CONV_AGG)
    cdf = cdf[cdf["Platform"].isin(platform_opt) & cdf["Audience"].isin(audience_opt)]

    if not cdf.empty:
        c1, c2 = st.columns([1.4, 2])
        with c1:
            d2 = cdf.copy(); d2["Hot Leads"] = d2["Hot Leads"].map(lambda v: f"{v:,}")
            st.dataframe(d2, use_container_width=True, hide_index=True)
        with c2:
            with st.container(border=True):
                cdf2 = cdf.copy()
                cdf2["label"] = cdf2["Hot Leads"].map(lambda v: f"{v:,}")
                fig_hl2 = px.bar(cdf2.sort_values("Hot Leads", ascending=False),
                                 x="Campaign Type", y="Hot Leads", color="Platform",
                                 text="label",
                                 color_discrete_map=BLUE_MAP, barmode="group",
                                 height=380, title="Hot Leads by Campaign Type")
                al(fig_hl2, yaxis=dict(title="Hot Leads"), xaxis=dict(title=""))
                fig_hl2.update_traces(textposition="outside", textfont_size=9)
                st.plotly_chart(fig_hl2, use_container_width=True)

    # Other campaigns breakdown
    st.markdown("<div class='section-header'>Other Campaigns — Sub-Type Breakdown</div>",
                unsafe_allow_html=True)

    other_rows = []
    for df, plat, sc, cc in [
        (gads_f,"Google Ads","Cost","Clicks"),
        (fb_f,"Facebook Ads","Amount spent (USD)","Results"),
    ]:
        others = df[df["Campaign Type"] == "Other"]
        if others.empty: continue
        for sub, g in others.groupby("Other Sub"):
            sp=g[sc].sum(); cl=g[cc].sum()
            other_rows.append({"Platform":plat, "Sub-Type":sub or "Misc",
                                "# Campaigns":len(g), "Spend ($)":sp,
                                "Clicks":cl, "CPC ($)":sp/cl if cl>0 else 0})

    if other_rows:
        odf = pd.DataFrame(other_rows)
        c1, c2 = st.columns([1.4, 2])
        with c1:
            od = odf.sort_values("Spend ($)", ascending=False).copy()
            od["Spend ($)"] = od["Spend ($)"].map(lambda v: f"${v:,.2f}")
            od["Clicks"]    = od["Clicks"].map(lambda v: f"{v:,.0f}")
            od["CPC ($)"]   = od["CPC ($)"].map(lambda v: f"${v:.2f}")
            st.dataframe(od, use_container_width=True, hide_index=True)
        with c2:
            with st.container(border=True):
                odf2 = odf.copy(); odf2["label"] = odf2["Spend ($)"].map(lambda v: f"${v:,.0f}")
                fig_o = px.bar(odf2.sort_values("Spend ($)", ascending=False),
                               x="Sub-Type", y="Spend ($)", color="Platform",
                               text="label", color_discrete_map=BLUE_MAP, barmode="group",
                               height=380, title="Other Campaigns: Spend by Sub-Type")
                al(fig_o, yaxis=dict(title="Spend ($)",tickprefix="$",tickformat=",.0f"),xaxis=dict(title=""))
                fig_o.update_traces(textposition="outside", textfont_size=9)
                st.plotly_chart(fig_o, use_container_width=True)
    else:
        st.info("No 'Other' campaigns in current selection.")


# ─────────────────────────────────────────────────────────────
# TAB 3 — Monthly Trends
# ─────────────────────────────────────────────────────────────
with tab3:
    st.caption("📌 Platform and Month filters (left panel) apply to all charts below.")

    # ── Main: HL by platform, month-on-month ─────────────────
    st.markdown("<div class='section-header'>Hot Leads — Month on Month</div>", unsafe_allow_html=True)

    if not ov_f.empty:
        with st.container(border=True):
            fig_m = go.Figure()
            traces = []
            if "Facebook Ads" in platform_opt: traces.append(("FB_HL","Facebook Ads","#1877f2"))
            if "Google Ads"    in platform_opt: traces.append(("GA_HL","Google Ads","#1a237e"))

            for col, name, color in traces:
                fig_m.add_trace(go.Bar(
                    x=ov_f["Month"], y=ov_f[col], name=name,
                    marker_color=color,
                    text=ov_f[col].map(lambda v: f"{int(v):,}"),
                    textposition="outside", textfont=dict(size=9, color="#344563"),
                    hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>",
                ))
            al(fig_m, title="Monthly Hot Leads by Platform", barmode="group", height=420,
               yaxis=dict(title="Hot Leads"),
               legend=dict(orientation="h", y=1.07, x=0))
            st.plotly_chart(fig_m, use_container_width=True)
    else:
        st.info("No conversion data for selected months.")

    # ── HL → QL% trend ────────────────────────────────────────
    st.markdown("<div class='section-header'>HL → Qualified Lead % by Month</div>", unsafe_allow_html=True)
    if not ov_f.empty:
        with st.container(border=True):
            fig_pct = go.Figure(go.Bar(
                x=ov_f["Month"], y=ov_f["HL_QL_pct"],
                marker_color="#00838f",
                text=ov_f["HL_QL_pct"].map(lambda v: f"{v:.0f}%"),
                textposition="outside", textfont=dict(size=9, color="#344563"),
                hovertemplate="<b>%{x}</b><br>HL→QL: %{y:.1f}%<extra></extra>",
            ))
            al(fig_pct, title="HL → QL Conversion Rate by Month", height=360,
               yaxis=dict(title="HL→QL%", ticksuffix="%"))
            st.plotly_chart(fig_pct, use_container_width=True)

    # ── Campaign Type Monthly Selector ───────────────────────
    st.markdown("<div class='section-header'>Monthly Breakdown by Campaign Type</div>",
                unsafe_allow_html=True)

    type_options = []
    if "Facebook Ads" in platform_opt:
        type_options += ["FB Lead Gen (Type)", "FB Lead Form"]
    type_options += ["Awareness (FB+GA)", "Retargeting (FB+GA)"]

    if type_options:
        type_sel = st.selectbox("Select Campaign Type", type_options)

        def bar_chart(df, ycols, names, colors, title):
            if df.empty: st.info("No data for selected months."); return
            with st.container(border=True):
                fig = go.Figure()
                for col, name, color in zip(ycols, names, colors):
                    fig.add_trace(go.Bar(
                        x=df["Month"], y=df[col], name=name, marker_color=color,
                        text=df[col].map(lambda v: f"{int(v):,}"),
                        textposition="outside", textfont=dict(size=9, color="#344563"),
                        hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}}<extra></extra>",
                    ))
                al(fig, title=title, barmode="group", height=400,
                   yaxis=dict(title="Hot Leads"),
                   legend=dict(orientation="h", y=1.07, x=0))
                st.plotly_chart(fig, use_container_width=True)

        if type_sel == "FB Lead Gen (Type)":
            bar_chart(fbt_f, ["HL","QL"], ["Hot Leads","Qualified Leads"],
                      ["#1877f2","#0d47a1"], "FB Lead Gen — Monthly HL & QL")

        elif type_sel == "FB Lead Form":
            bar_chart(fbf_f, ["HL","QL"], ["Hot Leads","Qualified Leads"],
                      ["#1877f2","#0d47a1"], "FB Lead Form — Monthly HL & QL")

        elif type_sel == "Awareness (FB+GA)":
            cols_t, names_t, colors_t = [], [], []
            if "Facebook Ads" in platform_opt:
                cols_t.append("FB"); names_t.append("Facebook Ads"); colors_t.append("#1877f2")
            if "Google Ads" in platform_opt:
                cols_t.append("GA"); names_t.append("Google Ads"); colors_t.append("#1a237e")
            bar_chart(aw_f, cols_t, names_t, colors_t, "Awareness — Monthly Hot Leads")

        elif type_sel == "Retargeting (FB+GA)":
            cols_t, names_t, colors_t = [], [], []
            if "Facebook Ads" in platform_opt:
                cols_t.append("FB"); names_t.append("Facebook Ads"); colors_t.append("#1877f2")
            if "Google Ads" in platform_opt:
                cols_t.append("GA"); names_t.append("Google Ads"); colors_t.append("#1a237e")
            bar_chart(rt_f, cols_t, names_t, colors_t, "Retargeting — Monthly Hot Leads")
    else:
        st.info("Select at least one platform to view campaign type breakdown.")
