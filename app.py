import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, timedelta

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Pipeline Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #F0F4F8; }
  [data-testid="stSidebar"]          { background: #1A237E; }
  [data-testid="stSidebar"] *        { color: #E8EAF6 !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stSlider label,
  [data-testid="stSidebar"] .stDateInput label {
    color: #B0BEC5 !important; font-size: 0.78rem;
    text-transform: uppercase; letter-spacing: .04em;
  }
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3 { color: #FFFFFF !important; }

  .kpi-card {
    background: #FFFFFF; border-radius: 10px; padding: 20px 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,.08); border-left: 5px solid #1976D2;
    margin-bottom: 4px;
  }
  .kpi-label  { font-size:.75rem; font-weight:600; text-transform:uppercase;
                letter-spacing:.06em; color:#607D8B; margin-bottom:4px; }
  .kpi-value  { font-size:2rem; font-weight:700; color:#1A237E; line-height:1.1; }
  .kpi-trend  { font-size:.82rem; margin-top:6px; }
  .trend-up   { color:#4CAF50; }
  .trend-down { color:#F44336; }
  .trend-flat { color:#9E9E9E; }

  .section-header {
    font-size:1rem; font-weight:700; color:#1A237E;
    border-bottom:2px solid #1976D2; padding-bottom:6px; margin-bottom:12px;
  }
  .alert-banner {
    background:#FFF3E0; border:1px solid #FFB300; border-radius:8px;
    padding:12px 18px; margin-bottom:16px; font-size:.9rem; color:#E65100;
  }
  .insight-box {
    background:linear-gradient(135deg,#E3F2FD 0%,#E8EAF6 100%);
    border-radius:10px; padding:16px 20px; border:1px solid #90CAF9;
    font-size:.88rem; color:#1A237E; line-height:1.7;
  }
  .insight-box strong { color:#0D47A1; }
  .insight-row { display:flex; gap:12px; margin-bottom:10px; align-items:flex-start; }
  .insight-icon { font-size:1.1rem; margin-top:1px; flex-shrink:0; }
  .insight-text { flex:1; }

  .health-green  { background:#E8F5E9; color:#2E7D32; padding:2px 10px; border-radius:12px; font-size:.78rem; font-weight:700; }
  .health-yellow { background:#FFFDE7; color:#F57F17; padding:2px 10px; border-radius:12px; font-size:.78rem; font-weight:700; }
  .health-red    { background:#FFEBEE; color:#C62828; padding:2px 10px; border-radius:12px; font-size:.78rem; font-weight:700; }

  .chart-card {
    background:#FFFFFF; border-radius:10px;
    box-shadow:0 2px 8px rgba(0,0,0,.08); padding:16px 16px 8px;
  }
  [data-testid="stDataFrame"] { border-radius:8px; overflow:hidden; }
  div[data-testid="stDataFrame"] > div { box-shadow:0 2px 8px rgba(0,0,0,.06); border-radius:8px; }
</style>
""", unsafe_allow_html=True)

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    return pd.read_csv("data/opportunities.csv", parse_dates=[
        "expected_close_date", "created_date", "last_activity_date"
    ])

df_raw = load_data()
# --TODAY  = pd.Timestamp("2026-03-31")
TODAY  = pd.Timestamp.now().normalize()

STAGE_ORDER = ["Prospecting", "Qualification", "Proposal", "Negotiation", "Closed Won", "Closed Lost"]
STAGE_COLORS = {
    "Prospecting":   "#90CAF9",
    "Qualification": "#42A5F5",
    "Proposal":      "#1976D2",
    "Negotiation":   "#0D47A1",
    "Closed Won":    "#4CAF50",
    "Closed Lost":   "#EF5350",
}

def fmt_m(v):
    """Show $X.XXM for >=1M, $XXXK for >=1K, else $X."""
    if abs(v) >= 1_000_000:
        return f"${v/1e6:.2f}M"
    if abs(v) >= 1_000:
        return f"${v/1e3:.0f}K"
    return f"${v:.0f}"

# ── Historic slippage (computed on full dataset, filter-independent) ──────────
_past = df_raw[df_raw["expected_close_date"] < TODAY]
_past_val      = _past["deal_value"].sum()
_slipped_val   = _past[_past["stage"].isin(
    ["Prospecting", "Qualification", "Proposal", "Negotiation"])]["deal_value"].sum()
_won_past_val  = _past[_past["stage"] == "Closed Won"]["deal_value"].sum()
_lost_past_val = _past[_past["stage"] == "Closed Lost"]["deal_value"].sum()
# Slippage rate = value still open past due / all past-due value
historic_slip  = (_slipped_val / _past_val) if _past_val else 0.40
# Worst-case multiplier default = proportion of weighted forecast that historically closes
historic_worst_mult = round(max(0.10, 1.0 - historic_slip), 2)

# ── Sidebar Filters ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔍 Filters")
    st.markdown("---")

    st.markdown("**Close Date Range**")
    min_date = df_raw["expected_close_date"].min().date()
    max_date = df_raw["expected_close_date"].max().date()
    date_from = st.date_input("From", value=min_date, min_value=min_date, max_value=max_date, key="d_from")
    date_to   = st.date_input("To",   value=max_date, min_value=min_date, max_value=max_date, key="d_to")

    st.markdown("---")
    all_regions  = sorted(df_raw["region"].unique())
    sel_regions  = st.multiselect("Region", all_regions, default=all_regions)

    all_products = ["All"] + sorted(df_raw["product_line"].unique())
    sel_product  = st.selectbox("Product Line", all_products)

    all_reps = ["All"] + sorted(df_raw["sales_rep"].unique())
    sel_rep  = st.selectbox("Sales Rep", all_reps)

    min_val, max_val = int(df_raw["deal_value"].min()), int(df_raw["deal_value"].max())
    deal_range = st.slider("Deal Size ($)", min_val, max_val, (min_val, max_val),
                           step=5000, format="$%d")

    st.markdown("---")
    st.markdown("**Stages**")
    sel_stages = [s for s in STAGE_ORDER if st.checkbox(s, value=True, key=f"stage_{s}")]

    st.markdown("---")
    st.markdown("**📉 Worst Case Forecast**")
    st.markdown(
        f"<div style='font-size:.75rem;color:#90CAF9;margin-bottom:6px'>"
        f"Historic slippage: <b>{historic_slip*100:.0f}%</b> of past-due pipeline "
        f"value never closed<br>"
        f"({fmt_m(_slipped_val)} still open of {fmt_m(_past_val)} past due)</div>",
        unsafe_allow_html=True,
    )
    worst_mult = st.slider(
        "Worst Case — % of Likely",
        min_value=10, max_value=100,
        value=int(historic_worst_mult * 100),
        step=5, format="%d%%",
        help="Default set from historic slippage in this dataset",
    )

    st.markdown("---")
    if st.button("↺  Reset Filters", use_container_width=True):
        st.rerun()

# ── Apply filters ─────────────────────────────────────────────────────────────
df = df_raw[
    (df_raw["expected_close_date"].dt.date >= date_from) &
    (df_raw["expected_close_date"].dt.date <= date_to) &
    (df_raw["region"].isin(sel_regions)) &
    (df_raw["deal_value"] >= deal_range[0]) &
    (df_raw["deal_value"] <= deal_range[1]) &
    (df_raw["stage"].isin(sel_stages))
].copy()
if sel_product != "All":
    df = df[df["product_line"] == sel_product]
if sel_rep != "All":
    df = df[df["sales_rep"] == sel_rep]

# ── Derived metrics ───────────────────────────────────────────────────────────
open_stages   = ["Prospecting", "Qualification", "Proposal", "Negotiation"]
df_open       = df[df["stage"].isin(open_stages)]
df_closed_won = df[df["stage"] == "Closed Won"]
df_closed_lst = df[df["stage"] == "Closed Lost"]
df_closed     = df[df["stage"].isin(["Closed Won", "Closed Lost"])]

total_pipeline = df_open["deal_value"].sum()
total_val      = df["deal_value"].sum()          # all deals (for % of total references)
open_count     = len(df_open)
avg_deal       = df_open["deal_value"].mean() if open_count else 0
weighted_pipe  = (df_open["deal_value"] * df_open["probability"] / 100).sum()

# ── Three win rate definitions (count-based) ──────────────────────────────────
n_won    = len(df_closed_won)
n_lost   = len(df_closed_lst)
n_closed = len(df_closed)

# 1. Realized: Closed Won / (Closed Won + Closed Lost)
realized_wr = (n_won / n_closed * 100) if n_closed else 0

# 2. Adjusted: treat overdue open deals as losses
overdue_mask = (
    df["stage"].isin(open_stages) &
    (df["expected_close_date"] < TODAY)
)
n_overdue = int(overdue_mask.sum())
adjusted_wr = (n_won / (n_closed + n_overdue) * 100) if (n_closed + n_overdue) else 0

# 3. Forecasted: project open deals forward at stage probability
expected_future_wins = (df_open["probability"] / 100).sum()
forecasted_wr = (
    (n_won + expected_future_wins) / (n_closed + open_count) * 100
) if (n_closed + open_count) else 0

# ── Three win rate definitions (revenue-weighted) ──────────────────────────────
won_val           = df_closed_won["deal_value"].sum()
lost_val          = df_closed_lst["deal_value"].sum()
overdue_val       = df[overdue_mask]["deal_value"].sum()
open_val          = df_open["deal_value"].sum()          # = total_pipeline
weighted_open_val = weighted_pipe                        # already computed above

# 1. Revenue Realized: $ Won / ($ Won + $ Lost)
rev_realized_wr = (won_val / (won_val + lost_val) * 100) if (won_val + lost_val) else 0

# 2. Revenue Adjusted: overdue open value treated as lost revenue
rev_adjusted_wr = (
    won_val / (won_val + lost_val + overdue_val) * 100
) if (won_val + lost_val + overdue_val) else 0

# 3. Revenue Forecasted: ($ Won + weighted open value) / ($ Won + $ Lost + $ Open)
rev_forecasted_wr = (
    (won_val + weighted_open_val) / (won_val + lost_val + open_val) * 100
) if (won_val + lost_val + open_val) else 0

# Pipeline health flag
overdue_pct_count = n_overdue / len(df) * 100 if len(df) else 0

# Stale active deals
stale_count = int((df_open["last_activity_date"] < (TODAY - timedelta(days=14))).sum())

# ── Header ────────────────────────────────────────────────────────────────────
_, col_title, _ = st.columns([0.05, 0.7, 0.25])
with col_title:
    st.markdown("## 📊 Sales Pipeline Dashboard")
    st.caption(f"Data as of {TODAY.strftime('%B %d, %Y')}  ·  {len(df):,} opportunities shown")

st.markdown("---")

# ── Smart Alerts ──────────────────────────────────────────────────────────────
alerts = []
if stale_count > 0:
    stale_val = df_open[df_open["last_activity_date"] < (TODAY - timedelta(days=14))]["deal_value"].sum()
    alerts.append(f"⚠️ <b>{stale_count} deals</b> untouched 14+ days — <b>{fmt_m(stale_val)}</b> at risk")
big_neg = df[(df["stage"] == "Negotiation") & (df["deal_value"] >= 100_000)]
if len(big_neg):
    alerts.append(f"💰 <b>{len(big_neg)} deals ≥$100K</b> in Negotiation — push to close now")

if alerts:
    st.markdown('<div class="alert-banner">' + " &nbsp;|&nbsp; ".join(alerts) + "</div>",
                unsafe_allow_html=True)

# ── Pipeline Health Warning (overdue > 10% of total pipeline) ─────────────────
if n_overdue > 0 and overdue_pct_count >= 10:
    st.markdown(f"""
    <div style="background:#FFF8E1;border:1px solid #F9A825;border-left:5px solid #F57F17;
                border-radius:8px;padding:14px 18px;margin-bottom:12px;font-size:.88rem;color:#4E342E">
      <b>⚠️ Pipeline Health Warning</b><br>
      <b>{n_overdue} deals (${overdue_val/1e6:.1f}M)</b> are past their expected close date
      but not marked as Won or Lost. This may indicate:<br>
      <span style="margin-left:12px">
        &bull; Optimistic close date forecasting by sales reps &nbsp;
        &bull; Stale data not updated in CRM &nbsp;
        &bull; Deals that should be marked Closed Lost
      </span><br>
      <b>Suggested action:</b> Review these {n_overdue} overdue deals with the sales team this week.
      Adjusted Win Rate drops from <b>{realized_wr:.1f}%</b> to <b>{adjusted_wr:.1f}%</b>
      when overdue deals are treated as losses.
    </div>
    """, unsafe_allow_html=True)

# ── Currency formatter ────────────────────────────────────────────────────────
def fmt_m(v):
    """Show $X.XXM for >=1M, $XXXK for >=1K, else $X."""
    if abs(v) >= 1_000_000:
        return f"${v/1e6:.2f}M"
    if abs(v) >= 1_000:
        return f"${v/1e3:.0f}K"
    return f"${v:.0f}"

# ── KPI Cards ─────────────────────────────────────────────────────────────────
def kpi_card(col, label, value, trend_text, trend_dir="flat", border_color="#1976D2"):
    trend_cls = {"up":"trend-up","down":"trend-down","flat":"trend-flat"}[trend_dir]
    arrow     = {"up":"▲","down":"▼","flat":"●"}[trend_dir]
    col.markdown(f"""
    <div class="kpi-card" style="border-left-color:{border_color}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-trend"><span class="{trend_cls}">{arrow} {trend_text}</span></div>
    </div>""", unsafe_allow_html=True)

# Row 1 — Pipeline KPIs
k1, k2, k3 = st.columns(3)
worst_pipe = weighted_pipe * (worst_mult / 100)
kpi_card(k1, "Total Pipeline Value",  f"${total_pipeline/1e6:.2f}M",
         f"{fmt_m(weighted_pipe)} pipeline weighted  ·  {fmt_m(worst_pipe)} worst case",
         "up", "#1976D2")
kpi_card(k2, "Open Opportunities",    f"{open_count:,}",
         f"{len(df_open[df_open['stage']=='Negotiation'])} in Negotiation", "up", "#0D47A1")
kpi_card(k3, "Avg Deal Size",         f"${avg_deal:,.0f}",
         "across active pipeline", "flat", "#7B1FA2")

# Row 2 — Three win rate KPIs
w1, w2, w3 = st.columns(3)

kpi_card(w1, "Win Rate (Closed Deals)",
         f"{realized_wr:.1f}%",
         f"Based on {n_closed} closed deals", "up", "#4CAF50")

# Adjusted win rate — flag if meaningfully lower than realized
adj_dir   = "down" if (realized_wr - adjusted_wr) > 5 else "flat"
adj_color = "#F44336" if (realized_wr - adjusted_wr) > 5 else "#FF6F00"
w2.markdown(f"""
<div class="kpi-card" style="border-left-color:{adj_color}">
  <div class="kpi-label">
    Adjusted Win Rate
    <span title="Treats deals past their expected close date as losses, providing a more conservative view of pipeline health."
          style="cursor:help;color:#90A4AE;font-size:.85rem;margin-left:4px">&#9432;</span>
  </div>
  <div class="kpi-value" style="color:{'#C62828' if adj_dir=='down' else '#E65100'}">{adjusted_wr:.1f}%</div>
  <div class="kpi-trend">
    <span class="trend-{'down' if adj_dir=='down' else 'flat'}">
      {'▼' if adj_dir=='down' else '●'}
      Including {n_overdue} overdue deal{'s' if n_overdue!=1 else ''} as losses
    </span>
  </div>
</div>""", unsafe_allow_html=True)

kpi_card(w3, "Forecasted Win Rate",
         f"{forecasted_wr:.1f}%",
         f"Open pipeline weighted at stage probability",
         "up" if forecasted_wr >= realized_wr else "flat", "#7B1FA2")

# Row 3 — Revenue-weighted win rates
st.markdown(
    "<div style='font-size:.72rem;font-weight:600;text-transform:uppercase;"
    "letter-spacing:.06em;color:#90A4AE;padding:8px 0 4px'>"
    "Revenue-Weighted Win Rates — dollar value of deals, not deal count</div>",
    unsafe_allow_html=True,
)
r1, r2, r3 = st.columns(3)

kpi_card(r1, "Revenue Win Rate (Closed)",
         f"{rev_realized_wr:.1f}%",
         f"{fmt_m(won_val)} won of {fmt_m(won_val+lost_val)} closed",
         "up", "#2E7D32")

rev_adj_dir   = "down" if (rev_realized_wr - rev_adjusted_wr) > 5 else "flat"
rev_adj_color = "#C62828" if rev_adj_dir == "down" else "#E65100"
r2.markdown(f"""
<div class="kpi-card" style="border-left-color:{rev_adj_color}">
  <div class="kpi-label">
    Revenue Adjusted Win Rate
    <span title="Overdue open deal value treated as lost revenue."
          style="cursor:help;color:#90A4AE;font-size:.85rem;margin-left:4px">&#9432;</span>
  </div>
  <div class="kpi-value" style="color:{rev_adj_color}">{rev_adjusted_wr:.1f}%</div>
  <div class="kpi-trend">
    <span class="trend-{rev_adj_dir}">
      {'▼' if rev_adj_dir=='down' else '●'}
      {fmt_m(overdue_val)} overdue value as losses
    </span>
  </div>
</div>""", unsafe_allow_html=True)

kpi_card(r3, "Revenue Forecasted Win Rate",
         f"{rev_forecasted_wr:.1f}%",
         f"{fmt_m(weighted_open_val)} weighted open pipeline",
         "up" if rev_forecasted_wr >= rev_realized_wr else "flat", "#6A1B9A")

st.markdown("<br>", unsafe_allow_html=True)

# ── Monthly Pipeline View  |  Revenue Forecast ───────────────────────────────
col_monthly, col_rev = st.columns([1.3, 1])

# — Monthly Pipeline by Stage ————————————————————————————————————————————————
with col_monthly:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-header">Monthly Pipeline — Deals by Stage & Outcome</div>',
                unsafe_allow_html=True)

    months_all  = pd.date_range("2025-10-01", "2026-06-30", freq="MS")
    month_lbls  = [m.strftime("%b '%y") for m in months_all]

    # Categories and colours (consistent across all months)
    CAT_ORDER  = ["Closed Won", "Closed Lost", "Overdue/Stuck",
                  "Negotiation", "Proposal", "Qualification", "Prospecting"]
    CAT_COLORS = {
        "Closed Won":    "#4CAF50",
        "Closed Lost":   "#EF5350",
        "Overdue/Stuck": "#FF6F00",
        "Negotiation":   "#0D47A1",
        "Proposal":      "#1976D2",
        "Qualification": "#42A5F5",
        "Prospecting":   "#90CAF9",
    }

    cat_vals  = {c: [] for c in CAT_ORDER}
    cat_cnts  = {c: [] for c in CAT_ORDER}
    tot_vals  = []
    tot_cnts  = []
    conv_rates = []    # Won/(Won+Lost) for past months

    for m in months_all:
        mo_end  = m + pd.offsets.MonthEnd(0)
        is_past = mo_end < TODAY
        mo_df   = df[(df["expected_close_date"] >= m) &
                     (df["expected_close_date"] <= mo_end)]

        won_mo  = mo_df[mo_df["stage"] == "Closed Won"]
        lost_mo = mo_df[mo_df["stage"] == "Closed Lost"]
        stuck   = mo_df[mo_df["stage"].isin(open_stages)] if is_past else mo_df.iloc[0:0]
        future  = mo_df[mo_df["stage"].isin(open_stages)] if not is_past else mo_df.iloc[0:0]

        cat_vals["Closed Won"].append(won_mo["deal_value"].sum())
        cat_vals["Closed Lost"].append(lost_mo["deal_value"].sum())
        cat_vals["Overdue/Stuck"].append(stuck["deal_value"].sum())
        cat_cnts["Closed Won"].append(len(won_mo))
        cat_cnts["Closed Lost"].append(len(lost_mo))
        cat_cnts["Overdue/Stuck"].append(len(stuck))

        for s in ["Negotiation", "Proposal", "Qualification", "Prospecting"]:
            s_df = future[future["stage"] == s]
            cat_vals[s].append(s_df["deal_value"].sum())
            cat_cnts[s].append(len(s_df))

        tot_vals.append(mo_df["deal_value"].sum())
        tot_cnts.append(len(mo_df))
        closed_mo = len(won_mo) + len(lost_mo)
        conv_rates.append(len(won_mo) / closed_mo * 100 if closed_mo else None)

    fig_m = go.Figure()
    for cat in CAT_ORDER:
        fig_m.add_trace(go.Bar(
            name=cat,
            x=month_lbls,
            y=cat_vals[cat],
            marker_color=CAT_COLORS[cat],
            customdata=[[v, c] for v, c in zip(cat_vals[cat], cat_cnts[cat])],
            hovertemplate=(
                f"<b>%{{x}} — {cat}</b><br>"
                "Value: $%{customdata[0]:,.0f}<br>"
                "Deals: %{customdata[1]}<extra></extra>"
            ),
        ))

    # Deal count on top of each bar
    for i, (lbl, cnt, tot) in enumerate(zip(month_lbls, tot_cnts, tot_vals)):
        fig_m.add_annotation(
            x=lbl, y=tot, text=f"n={cnt}",
            showarrow=False, yshift=6,
            font=dict(size=8, color="#455A64"),
        )

    # Conversion rate label for past months
    past_end_idx = next((i for i, m in enumerate(months_all)
                         if (m + pd.offsets.MonthEnd(0)) >= TODAY), len(months_all)) - 1
    for i in range(past_end_idx + 1):
        if conv_rates[i] is not None:
            fig_m.add_annotation(
                x=month_lbls[i], y=tot_vals[i],
                text=f"{conv_rates[i]:.0f}%W",
                showarrow=False, yshift=18,
                font=dict(size=7, color="#2E7D32", family="Arial Black"),
            )

    # TODAY divider line (between last past and first future month)
    fig_m.add_shape(
        type="line",
        x0=past_end_idx + 0.5, x1=past_end_idx + 0.5,
        y0=0, y1=1, xref="x", yref="paper",
        line=dict(color="#607D8B", width=2, dash="dash"),
    )
    fig_m.add_annotation(
        x=past_end_idx + 0.5, y=1.02,
        xref="x", yref="paper",
        text="Today", showarrow=False,
        font=dict(size=9, color="#607D8B"),
    )
    fig_m.add_annotation(x=past_end_idx - 1.5, y=1.02,
        xref="x", yref="paper", text="Historical",
        showarrow=False, font=dict(size=9, color="#78909C"))
    fig_m.add_annotation(x=past_end_idx + 2, y=1.02,
        xref="x", yref="paper", text="Forecast",
        showarrow=False, font=dict(size=9, color="#1976D2"))

    fig_m.update_layout(
        barmode="stack",
        height=360,
        margin=dict(l=0, r=10, t=28, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.06,
                    xanchor="left", x=0, font=dict(size=9),
                    traceorder="normal"),
        xaxis=dict(tickfont=dict(size=10), showgrid=False),
        yaxis=dict(tickprefix="$", tickformat=",.0f",
                   tickfont=dict(size=10), showgrid=True, gridcolor="#F0F0F0"),
        hovermode="x unified",
    )
    st.plotly_chart(fig_m, use_container_width=True, config={"displayModeBar": False})

    # Summary strip
    stuck_total_val = sum(cat_vals["Overdue/Stuck"])
    stuck_total_cnt = sum(cat_cnts["Overdue/Stuck"])
    st.markdown(
        f"<div style='font-size:.76rem;color:#607D8B;padding:2px 0 8px'>"
        f"Green=Won &nbsp;|&nbsp; Red=Lost &nbsp;|&nbsp; "
        f"Orange=Overdue/Stuck ({stuck_total_cnt} deals, {fmt_m(stuck_total_val)} missed revenue) "
        f"&nbsp;|&nbsp; Blues=Active Pipeline &nbsp;|&nbsp; "
        f"%W = monthly Won rate"
        f"</div>",
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

# — Revenue Forecast Timeline —
with col_rev:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    months = pd.date_range(TODAY.to_period('M').to_timestamp(), periods=3, freq="MS")
    _hdr_range = f"{months[0].strftime('%b')}–{months[-1].strftime('%b %Y')}"
    st.markdown(f'<div class="section-header">Revenue Forecast — Next 3 Months ({_hdr_range})</div>',
                unsafe_allow_html=True)

    # All three scenarios are rep-committed:
    #   Best Case  — 100% of deal value for deals tagged to that month
    #   Likely Case — deal value × stage probability (rep-entered)
    #   Worst Case  — Likely × slider % (derived from historic slippage)
    best_case, likely_case, worst_case = [], [], []

    for m in months:
        mo_end = m + pd.offsets.MonthEnd(0)
        mo_df  = df[
            (df["expected_close_date"] >= m) &
            (df["expected_close_date"] <= mo_end) &
            (df["stage"].isin(open_stages + ["Closed Won"]))
        ]
        best_case.append(mo_df["deal_value"].sum())
        likely = (mo_df["deal_value"] * mo_df["probability"] / 100).sum()
        likely_case.append(likely)
        worst_case.append(likely * (worst_mult / 100))

    months_str = [m.strftime("%b %Y") for m in months]

    fig_rev = go.Figure()
    fig_rev.add_trace(go.Scatter(
        x=months_str + months_str[::-1],
        y=best_case + worst_case[::-1],
        fill="toself", fillcolor="rgba(25,118,210,0.08)",
        line=dict(color="rgba(0,0,0,0)"),
        name="Forecast Range", hoverinfo="skip",
    ))
    fig_rev.add_trace(go.Scatter(
        x=months_str, y=best_case, mode="lines+markers",
        name="Best Case (100% close rate)",
        line=dict(color="#90CAF9", width=2, dash="dot"), marker=dict(size=6),
        hovertemplate="%{x}<br>Best Case: $%{y:,.0f}<extra></extra>",
    ))
    fig_rev.add_trace(go.Scatter(
        x=months_str, y=likely_case, mode="lines+markers",
        name="Likely Case (stage probability)",
        line=dict(color="#1976D2", width=3), marker=dict(size=8),
        hovertemplate="%{x}<br>Likely Case: $%{y:,.0f}<extra></extra>",
    ))
    fig_rev.add_trace(go.Scatter(
        x=months_str, y=worst_case, mode="lines+markers",
        name=f"Worst Case ({worst_mult}% of Likely)",
        line=dict(color="#F44336", width=2, dash="dash"), marker=dict(size=6),
        hovertemplate="%{x}<br>Worst Case: $%{y:,.0f}<extra></extra>",
    ))

    fig_rev.update_layout(
        height=330, margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        xaxis=dict(tickfont=dict(size=11), showgrid=True, gridcolor="#F0F0F0"),
        yaxis=dict(tickprefix="$", tickformat=",.0f", tickfont=dict(size=11),
                   showgrid=True, gridcolor="#F0F0F0"),
        hovermode="x unified",
    )
    st.plotly_chart(fig_rev, use_container_width=True, config={"displayModeBar": False})

    likely_total = sum(likely_case)
    best_total   = sum(best_case)
    st.markdown(
        f"<div style='font-size:.76rem;color:#607D8B;padding:4px 0 8px'>"
        f"All scenarios use <b>rep-committed close dates</b> and stage probabilities. &nbsp;·&nbsp; "
        f"3-month Likely total: <b>{fmt_m(likely_total)}</b> &nbsp;·&nbsp; "
        f"Best Case upside: <b>{fmt_m(best_total-likely_total)}</b> &nbsp;·&nbsp; "
        f"Worst Case slider adjusts for historic {100-worst_mult}% slippage."
        f"</div>",
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Row 3: Top 10 Table  |  Sales Rep Performance ────────────────────────────
col_table, col_reps = st.columns([1.2, 1])

# — Top 10 Table —
with col_table:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-header">Top 10 Opportunities by Value</div>',
                unsafe_allow_html=True)

    top10 = df.nlargest(10, "deal_value")[
        ["account_name", "deal_value", "stage", "probability",
         "expected_close_date", "sales_rep", "last_activity_date"]
    ].copy()

    def health_score(row):
        if row["stage"] in ("Closed Won", "Closed Lost"):
            return "Closed"
        days = (TODAY - row["last_activity_date"]).days
        if days > 21: return "🔴 At Risk"
        if days > 10: return "🟡 Monitor"
        return "🟢 Healthy"

    top10["Health"]     = top10.apply(health_score, axis=1)
    top10["Close Date"] = top10["expected_close_date"].dt.strftime("%b %d, %Y")
    top10["Deal Value"] = top10["deal_value"].apply(lambda v: f"${v:,.0f}")
    top10["Prob."]      = top10["probability"].apply(lambda p: f"{p}%")

    display_df = top10[["account_name","Deal Value","stage","Prob.","Close Date","sales_rep","Health"]
                       ].rename(columns={"account_name":"Account","stage":"Stage","sales_rep":"Rep"})

    def highlight_rows(row):
        val = float(row["Deal Value"].replace("$", "").replace(",", ""))
        if val >= 100_000:
            return ["font-weight:bold; background-color:#E3F2FD"] * len(row)
        return [""] * len(row)

    st.dataframe(display_df.style.apply(highlight_rows, axis=1),
                 use_container_width=True, height=340, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# — Sales Rep Performance (quota-based) —
with col_reps:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-header">Sales Rep Pipeline Performance</div>',
                unsafe_allow_html=True)

    QUOTA = 500_000  # per rep

    rep_stage = (
        df[df["stage"].isin(open_stages)]
        .groupby(["sales_rep", "stage"], observed=True)["deal_value"]
        .sum().reset_index()
    )
    rep_totals = rep_stage.groupby("sales_rep")["deal_value"].sum().sort_values()
    rep_order  = rep_totals.index.tolist()

    # Keep team_avg for use in insights section below
    team_avg = rep_totals.mean() if len(rep_totals) else 1

    def quota_color(total):
        pct = total / QUOTA
        if pct >= 1.0:  return "#4CAF50"
        if pct >= 0.70: return "#FFB300"
        return "#F44336"

    fig_rep = go.Figure()
    for s in open_stages:
        s_map = dict(zip(
            rep_stage[rep_stage["stage"] == s]["sales_rep"],
            rep_stage[rep_stage["stage"] == s]["deal_value"],
        ))
        fig_rep.add_trace(go.Bar(
            y=rep_order,
            x=[s_map.get(r, 0) for r in rep_order],
            name=s, orientation="h",
            marker_color=STAGE_COLORS[s],
            hovertemplate=f"<b>%{{y}}</b><br>{s}: $%{{x:,.0f}}<extra></extra>",
        ))

    # Quota line
    fig_rep.add_vline(x=QUOTA, line_dash="dash", line_color="#FF7043", line_width=2,
                      annotation_text="Quota", annotation_position="top right",
                      annotation_font=dict(color="#FF7043", size=11))

    # Color-coded dot per rep
    for rep in rep_order:
        fig_rep.add_annotation(
            x=0, y=rep, text="●", showarrow=False,
            xanchor="right", xshift=-4,
            font=dict(color=quota_color(rep_totals[rep]), size=14),
        )

    fig_rep.update_layout(
        barmode="stack", height=340,
        margin=dict(l=20, r=20, t=10, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="left", x=0, font=dict(size=10)),
        xaxis=dict(tickprefix="$", tickformat=",.0f", tickfont=dict(size=10),
                   showgrid=True, gridcolor="#F0F0F0"),
        yaxis=dict(tickfont=dict(size=11)),
    )
    st.plotly_chart(fig_rep, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Row 4: AI Insights  |  Deal Health ────────────────────────────────────────
col_ins, col_health = st.columns([1.6, 1])

# — AI-Generated Insights —
with col_ins:
    st.markdown('<div class="section-header">AI-Generated Pipeline Insights</div>',
                unsafe_allow_html=True)

    # Compute insight components
    top_rep_series = rep_stage.groupby("sales_rep")["deal_value"].sum()
    top_rep     = top_rep_series.idxmax() if len(top_rep_series) else "N/A"
    top_rep_val = top_rep_series.max() if len(top_rep_series) else 0
    top_rep_mult = top_rep_val / team_avg if team_avg else 0

    top_prod = (df[df["stage"].isin(open_stages)]
                .groupby("product_line")["deal_value"].sum()
                .idxmax() if open_count else "N/A")

    next30     = df_open[
        (df_open["expected_close_date"] >= TODAY) &
        (df_open["expected_close_date"] <= TODAY + timedelta(days=30))
    ]
    neg_df     = df[df["stage"] == "Negotiation"]
    neg_val    = neg_df["deal_value"].sum()
    neg_cnt    = len(neg_df)

    # Stage with most value stuck
    top_active_stage = (df_open.groupby("stage")["deal_value"].sum().idxmax()
                        if open_count else "N/A")
    top_active_val   = (df_open.groupby("stage")["deal_value"].sum().max()
                        if open_count else 0)

    closed_total = n_closed
    wr_ctx = "above industry average of ~20%" if realized_wr > 20 else "below industry average of ~20% — review qualification criteria"

    stale_val_at_risk = df_open[
        df_open["last_activity_date"] < (TODAY - timedelta(days=14))
    ]["deal_value"].sum()

    insights = [
        ("📊", f"<strong>Realized Win Rate: {realized_wr:.1f}%</strong> (count) / "
               f"<strong>{rev_realized_wr:.1f}%</strong> (revenue) — "
               f"{n_won} deals won, {fmt_m(won_val)} captured of {fmt_m(won_val+lost_val)} closed. "
               f"Adjusted to <strong>{adjusted_wr:.1f}%</strong> when {n_overdue} overdue deals treated as losses. "
               f"This is {wr_ctx}."),
        ("🔻", f"<strong>Pipeline Conversion Bottleneck:</strong> "
               f"{fmt_m(top_active_val)} ({top_active_val/total_val*100:.0f}% of total value) "
               f"is concentrated in <strong>{top_active_stage}</strong>. "
               f"Focus qualification efforts here to accelerate movement downstream."),
        ("🤝", f"<strong>Near-Term Revenue:</strong> {neg_cnt} deals worth "
               f"<strong>{fmt_m(neg_val)}</strong> are in Negotiation — highest-probability revenue this quarter. "
               f"Additionally, {len(next30)} deals worth <strong>{fmt_m(next30['deal_value'].sum())}</strong> "
               f"are expected to close in the next 30 days."),
        ("🏆", f"<strong>{top_rep}</strong> leads pipeline at "
               f"<strong>{fmt_m(top_rep_val)} ({top_rep_mult:.1f}x team average)</strong>. "
               f"Top product line: <strong>{top_prod}</strong>. "
               + (f"<strong>{stale_count} stale deals</strong> represent <strong>{fmt_m(stale_val_at_risk)}</strong> "
                  f"at risk — reassign or re-engage within 48 hours."
                  if stale_count > 0 else "All active deals have recent activity — pipeline is well-maintained.")),
    ]

    html_rows = "".join(
        f'<div class="insight-row">'
        f'<div class="insight-icon">{icon}</div>'
        f'<div class="insight-text">{text}</div>'
        f'</div>'
        for icon, text in insights
    )
    st.markdown(f'<div class="insight-box">{html_rows}</div>', unsafe_allow_html=True)

# — Deal Health Donut —
with col_health:
    st.markdown('<div class="section-header">Deal Health Overview</div>',
                unsafe_allow_html=True)

    def classify_health(row):
        days = (TODAY - row["last_activity_date"]).days
        if days > 21: return "At Risk"
        if days > 10: return "Monitor"
        return "Healthy"

    df_open_h = df_open.copy()
    df_open_h["health"] = df_open_h.apply(classify_health, axis=1)
    hc = df_open_h["health"].value_counts()
    healthy = hc.get("Healthy", 0)
    monitor = hc.get("Monitor", 0)
    at_risk = hc.get("At Risk", 0)
    total_h = healthy + monitor + at_risk

    fig_h = go.Figure(go.Pie(
        labels=["Healthy", "Monitor", "At Risk"],
        values=[healthy, monitor, at_risk],
        marker_colors=["#4CAF50", "#FFB300", "#F44336"],
        hole=0.55,
        textinfo="label+percent",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>%{value} deals (%{percent})<extra></extra>",
    ))
    fig_h.add_annotation(
        text=f"<b>{total_h}</b><br><span style='font-size:11px'>Deals</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=18, color="#1A237E"),
    )
    fig_h.update_layout(
        height=240, margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=-0.1,
                    xanchor="center", x=0.5, font=dict(size=11)),
    )
    st.plotly_chart(fig_h, use_container_width=True, config={"displayModeBar": False})

    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div style="text-align:center"><span class="health-green">🟢 {healthy}</span></div>',
                unsafe_allow_html=True)
    c2.markdown(f'<div style="text-align:center"><span class="health-yellow">🟡 {monitor}</span></div>',
                unsafe_allow_html=True)
    c3.markdown(f'<div style="text-align:center"><span class="health-red">🔴 {at_risk}</span></div>',
                unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#90A4AE;font-size:.78rem;padding:8px 0'>"
    "Sales Pipeline Dashboard &nbsp;·&nbsp; Streamlit + Plotly &nbsp;·&nbsp; "
    "AI-enhanced insights &nbsp;·&nbsp; Demonstrating AI as a Tableau replacement"
    "</div>",
    unsafe_allow_html=True,
)
