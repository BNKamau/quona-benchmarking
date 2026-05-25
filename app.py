import streamlit as st
import sqlite3
import pandas as pd
import plotly.graph_objects as go
import anthropic
import os
from datetime import datetime
from parsers.excel_parsers import PARSERS, SUPPORTED_COMPANIES

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Quona Portfolio Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Brand palette ─────────────────────────────────────────────────────────────
GREEN  = "#D5FA94"
BLACK  = "#2C2C2A"
BLUE   = "#C5E5FF"
BG     = "#EFF0EA"
WHITE  = "#FFFFFF"
BORDER = "#DDE0D8"
MUTED  = "#888884"
WARN   = "#E65100"
WARN_BG = "#FFF3E0"

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  .stApp {{ background-color:{BG}; color:{BLACK}; }}
  #MainMenu, footer, header {{ visibility:hidden; }}
  .block-container {{ padding-top:1.5rem; padding-bottom:2rem; max-width:1400px; }}

  [data-testid="metric-container"] {{
      background:{WHITE}; border:1px solid {BORDER};
      border-radius:10px; padding:16px 20px;
  }}
  [data-testid="stMetricLabel"] {{ color:{MUTED}; font-size:11px; text-transform:uppercase; letter-spacing:.5px; }}
  [data-testid="stMetricValue"] {{ color:{BLACK}; font-size:22px; font-weight:700; }}

  .stButton > button {{
      background:{GREEN}; color:{BLACK}; border:none;
      border-radius:8px; font-weight:600; padding:8px 20px;
  }}
  .stButton > button:hover {{ background:#bfe07c; color:{BLACK}; }}
  .stButton > button:focus {{ box-shadow:none; border:none; }}

  [data-testid="stHorizontalBlock"] [data-testid="column"]:first-child .stButton > button {{
      background: transparent !important;
      color: {BLACK} !important;
      border: none !important;
      box-shadow: none !important;
      padding: 2px 0 !important;
      font-size: 14px !important;
      font-weight: 600 !important;
      text-decoration: underline !important;
      text-underline-offset: 3px !important;
      text-align: left !important;
      width: auto !important;
      min-width: unset !important;
      border-radius: 0 !important;
  }}
  [data-testid="stHorizontalBlock"] [data-testid="column"]:first-child .stButton > button:hover {{
      background: transparent !important;
      color: #555 !important;
  }}

  [data-testid="stDataFrame"] {{ border-radius:10px; overflow:hidden; }}
  hr {{ border-color:{BORDER}; margin:1.2rem 0; }}
</style>
""", unsafe_allow_html=True)

# ── DB helpers ─────────────────────────────────────────────────────────────────
DB_PATH = "benchmarking.db"

def _conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

# ── Exit comps DB helpers ──────────────────────────────────────────────────────
COMPS_DB = "data/quona_exit_comps.db"
_COMP_NAME_MAP = {"Verto": "Verto FX"}  # benchmarking.db name → portfolio_comp_mapping name

def _comps_conn():
    return sqlite3.connect(COMPS_DB, check_same_thread=False)

@st.cache_data(ttl=300)
def load_comp_mapping(company_name: str) -> pd.DataFrame:
    name = _COMP_NAME_MAP.get(company_name, company_name)
    return pd.read_sql_query(
        "SELECT comp_id, relevance_score, mapping_rationale "
        "FROM portfolio_comp_mapping WHERE portfolio_company = ? "
        "ORDER BY relevance_score DESC",
        _comps_conn(), params=(name,),
    )

@st.cache_data(ttl=300)
def load_comps_detail(comp_ids: tuple) -> pd.DataFrame:
    if not comp_ids:
        return pd.DataFrame()
    ph = ",".join("?" * len(comp_ids))
    return pd.read_sql_query(f"""
        SELECT comp_id, company_name, sub_sector, geography,
               exit_status, exit_year, exit_type, exit_ev_usd_m,
               revenue_at_exit_usd_m, gross_margin_pct, ebitda_margin_pct,
               ev_revenue_multiple, data_confidence, key_narrative_drivers,
               revenue_growth_at_exit
        FROM exit_comps WHERE comp_id IN ({ph})
    """, _comps_conn(), params=list(comp_ids))

@st.cache_data(ttl=300)
def load_stage_snapshots(comp_ids: tuple) -> pd.DataFrame:
    if not comp_ids:
        return pd.DataFrame()
    ph = ",".join("?" * len(comp_ids))
    return pd.read_sql_query(f"""
        SELECT comp_id, company_name, stage, revenue_range_usd_m,
               revenue_growth_pct, gross_margin_pct, ebitda_margin_pct
        FROM comp_stage_snapshots WHERE comp_id IN ({ph})
    """, _comps_conn(), params=list(comp_ids))

@st.cache_data(ttl=300)
def load_companies() -> pd.DataFrame:
    return pd.read_sql_query("""
        SELECT c.id, c.name, c.sector, c.hq_country, c.founded_year,
               k.revenue_usd,
               k.ebitda_usd,
               k.gross_margin_pct,
               COALESCE(
                   k.ebitda_margin_pct,
                   CASE WHEN k.revenue_usd > 0 AND k.ebitda_usd IS NOT NULL
                        THEN ROUND(k.ebitda_usd * 100.0 / k.revenue_usd, 2)
                   END
               ) AS ebitda_margin_pct,
               k.period_end_date,
               k.customer_count,
               k.aum_usd,
               k.gmv_usd
        FROM companies c
        LEFT JOIN kpi_snapshots k
            ON k.company_id = c.id
            AND k.period_end_date = (
                SELECT MAX(k2.period_end_date)
                FROM kpi_snapshots k2 WHERE k2.company_id = c.id
            )
        ORDER BY c.name
    """, _conn())

@st.cache_data(ttl=300)
def load_revenue_growth() -> pd.DataFrame:
    return pd.read_sql_query("""
        WITH ranked AS (
            SELECT company_id, revenue_usd, period_end_date,
                   ROW_NUMBER() OVER (
                       PARTITION BY company_id ORDER BY period_end_date DESC
                   ) AS rn
            FROM kpi_snapshots
            WHERE revenue_usd IS NOT NULL
        )
        SELECT
            r1.company_id AS id,
            CASE
                WHEN r2.revenue_usd > 0
                THEN ROUND((r1.revenue_usd - r2.revenue_usd) * 100.0 / r2.revenue_usd, 1)
            END AS revenue_growth_pct
        FROM ranked r1
        LEFT JOIN ranked r2
            ON r1.company_id = r2.company_id AND r2.rn = 2
        WHERE r1.rn = 1
    """, _conn())

@st.cache_data(ttl=300)
def load_ltm_revenue() -> pd.DataFrame:
    """
    LTM (last 12 months) or ARR-estimated revenue per company.
    - Monthly reporters: sum of last 12 monthly periods
    - Quarterly reporters: sum of last 4 quarterly periods
    - Annual reporters: last annual figure
    - If insufficient history: annualise available data, label 'ARR (est.)'
    """
    conn = _conn()

    periods = pd.read_sql_query("""
        WITH ranked AS (
            SELECT company_id, period_end_date,
                   ROW_NUMBER() OVER (
                       PARTITION BY company_id ORDER BY period_end_date DESC
                   ) AS rn
            FROM kpi_snapshots
        ),
        gaps AS (
            SELECT r1.company_id,
                   CAST(julianday(r1.period_end_date)
                        - julianday(r2.period_end_date) AS INTEGER) AS gap_days
            FROM ranked r1
            JOIN ranked r2
                ON r1.company_id = r2.company_id AND r2.rn = 2
            WHERE r1.rn = 1
        )
        SELECT company_id AS id,
               gap_days,
               CASE WHEN gap_days <= 45  THEN 'monthly'
                    WHEN gap_days <= 135 THEN 'quarterly'
                    ELSE                      'annual' END AS period_type,
               CASE WHEN gap_days <= 45  THEN 12
                    WHEN gap_days <= 135 THEN  4
                    ELSE                       1 END AS needed
        FROM gaps
    """, conn)

    rev = pd.read_sql_query("""
        SELECT company_id AS id, period_end_date, revenue_usd
        FROM kpi_snapshots
        WHERE revenue_usd IS NOT NULL
        ORDER BY company_id, period_end_date DESC
    """, conn)

    ebitda_data = pd.read_sql_query("""
        SELECT company_id AS id, period_end_date, ebitda_usd
        FROM kpi_snapshots
        WHERE ebitda_usd IS NOT NULL
        ORDER BY company_id, period_end_date DESC
    """, conn)

    gm_data = pd.read_sql_query("""
        SELECT company_id AS id, period_end_date, revenue_usd, gross_margin_pct
        FROM kpi_snapshots
        WHERE gross_margin_pct IS NOT NULL AND revenue_usd IS NOT NULL
        ORDER BY company_id, period_end_date DESC
    """, conn)

    companies = pd.read_sql_query("SELECT id FROM companies", conn)

    results = []
    for cid in companies["id"]:
        cid = int(cid)
        crev = rev[rev["id"] == cid]
        n = len(crev)

        pt_row = periods[periods["id"] == cid]
        if pt_row.empty:
            period_type, needed = "monthly", 12
        else:
            period_type = pt_row.iloc[0]["period_type"]
            needed      = int(pt_row.iloc[0]["needed"])

        ltm_ebitda_usd        = None
        ltm_ebitda_margin_pct = None
        ltm_gross_margin_pct  = None

        if n == 0:
            results.append({"id": cid, "ltm_revenue": None,
                            "ltm_label": "—", "ltm_periods_used": 0,
                            "period_type": period_type, "periods_needed": needed,
                            "ltm_ebitda_usd": None, "ltm_ebitda_margin_pct": None,
                            "ltm_gross_margin_pct": None})
        elif n >= needed:
            ltm = float(crev.head(needed)["revenue_usd"].sum())
            top_periods = set(crev.head(needed)["period_end_date"].tolist())
            ce = ebitda_data[
                (ebitda_data["id"] == cid) &
                (ebitda_data["period_end_date"].isin(top_periods))
            ]
            if len(ce) == needed:
                ltm_ebitda_usd = float(ce["ebitda_usd"].sum())
                if ltm > 0:
                    ltm_ebitda_margin_pct = round(ltm_ebitda_usd / ltm * 100, 4)
            # LTM gross margin: weighted average over LTM periods that have GM data
            cgm = gm_data[
                (gm_data["id"] == cid) &
                (gm_data["period_end_date"].isin(top_periods))
            ]
            if len(cgm) > 0:
                gp_sum  = (cgm["revenue_usd"] * cgm["gross_margin_pct"] / 100).sum()
                rev_sum = cgm["revenue_usd"].sum()
                if rev_sum > 0:
                    ltm_gross_margin_pct = round(gp_sum / rev_sum * 100, 4)
            results.append({"id": cid, "ltm_revenue": ltm,
                            "ltm_label": "LTM", "ltm_periods_used": needed,
                            "period_type": period_type, "periods_needed": needed,
                            "ltm_ebitda_usd": ltm_ebitda_usd,
                            "ltm_ebitda_margin_pct": ltm_ebitda_margin_pct,
                            "ltm_gross_margin_pct": ltm_gross_margin_pct})
        else:
            ltm = float(crev["revenue_usd"].sum() * (needed / n))
            # Still compute partial LTM gross margin for ARR-estimated companies
            partial_periods = set(crev["period_end_date"].tolist())
            cgm = gm_data[
                (gm_data["id"] == cid) &
                (gm_data["period_end_date"].isin(partial_periods))
            ]
            if len(cgm) > 0:
                gp_sum  = (cgm["revenue_usd"] * cgm["gross_margin_pct"] / 100).sum()
                rev_sum = cgm["revenue_usd"].sum()
                if rev_sum > 0:
                    ltm_gross_margin_pct = round(gp_sum / rev_sum * 100, 4)
            results.append({"id": cid, "ltm_revenue": ltm,
                            "ltm_label": "ARR (est.)", "ltm_periods_used": n,
                            "period_type": period_type, "periods_needed": needed,
                            "ltm_ebitda_usd": None, "ltm_ebitda_margin_pct": None,
                            "ltm_gross_margin_pct": ltm_gross_margin_pct})

    return pd.DataFrame(results)

@st.cache_data(ttl=300)
def load_all_revenue() -> pd.DataFrame:
    return pd.read_sql_query("""
        SELECT company_id AS id, period_end_date, revenue_usd
        FROM kpi_snapshots
        WHERE revenue_usd IS NOT NULL
        ORDER BY id, period_end_date
    """, _conn())

@st.cache_data(ttl=300)
def load_company_info(company_id: int) -> pd.Series:
    df = pd.read_sql_query(
        "SELECT * FROM companies WHERE id = ?", _conn(), params=(company_id,)
    )
    return df.iloc[0]

@st.cache_data(ttl=300)
def load_kpis(company_id: int) -> pd.DataFrame:
    df = pd.read_sql_query("""
        SELECT period_end_date,
               revenue_usd, gross_margin_pct,
               ebitda_usd, ebitda_margin_pct,
               arr_usd, mrr_usd,
               customer_count, active_clients_count,
               net_revenue_retention_pct, cac_usd, ltv_usd,
               loan_book_gross_usd, par_30_pct, par_90_pct,
               npl_rate_pct, net_yield_pct, nim_pct,
               aum_usd, gmv_usd, tpv_usd,
               unique_borrowers_count
        FROM kpi_snapshots
        WHERE company_id = ?
        ORDER BY period_end_date
    """, _conn(), params=(company_id,))
    df["period_end_date"] = pd.to_datetime(df["period_end_date"])
    mask = (
        df["ebitda_margin_pct"].isna()
        & df["ebitda_usd"].notna()
        & (df["revenue_usd"].fillna(0) > 0)
    )
    df.loc[mask, "ebitda_margin_pct"] = (
        df.loc[mask, "ebitda_usd"] / df.loc[mask, "revenue_usd"] * 100
    ).round(2)
    return df

# ── Formatters ─────────────────────────────────────────────────────────────────
def _is_null(v) -> bool:
    if v is None:
        return True
    try:
        return pd.isna(v)
    except Exception:
        return False

def fmt_usd(v) -> str:
    if _is_null(v):
        return "—"
    v = float(v)
    if v >= 1e9: return f"${v/1e9:.1f}B"
    if v >= 1e6: return f"${v/1e6:.1f}M"
    if v >= 1e3: return f"${v/1e3:.0f}K"
    return f"${v:.0f}"

def fmt_pct(v) -> str:
    if _is_null(v):
        return "—"
    return f"{float(v):.1f}%"

def fmt_int(v) -> str:
    if _is_null(v):
        return "—"
    return f"{int(v):,}"

def fmt_growth(v) -> tuple[str, str]:
    if _is_null(v):
        return "—", MUTED
    v = float(v)
    sign = "+" if v > 0 else ""
    color = "#2E7D32" if v > 0 else ("#C62828" if v < 0 else BLACK)
    return f"{sign}{v:.1f}%", color

def as_of(date_val) -> str:
    if _is_null(date_val):
        return "No data"
    try:
        parsed = pd.to_datetime(date_val)
        if pd.isna(parsed):
            return "No data"
        return parsed.strftime("%b %Y")
    except Exception:
        return "No data"

def fmt_period_label(date_val, period_type: str = "monthly") -> str:
    """Return a short period label: 'Dec \'25', 'Q4 \'25', or 'FY2025'."""
    if _is_null(date_val):
        return ""
    try:
        d = pd.to_datetime(date_val)
        if pd.isna(d):
            return ""
        if period_type == "quarterly":
            q = (d.month - 1) // 3 + 1
            return f"Q{q} '{d.strftime('%y')}"
        elif period_type == "annual":
            return f"FY{d.year}"
        return f"{d.strftime('%b')} '{d.strftime('%y')}"
    except Exception:
        return ""

SECTOR_LABELS = {
    "wealth_management": "Wealth Mgmt",
    "payments":          "Payments",
    "lending":           "Lending",
    "insurtech":         "InsurTech",
    "iot_infrastructure":"IoT Infra",
    "saas":              "SaaS",
    "marketplace":       "Marketplace",
}

def sector_label(s: str) -> str:
    return SECTOR_LABELS.get(s, (s or "").replace("_", " ").title())

# ── Benchmarking helpers ───────────────────────────────────────────────────────
def _parse_pct(s) -> float | None:
    """Parse text metrics like '~40%', '60%+', '(30%)' to float."""
    if pd.isna(s) or s is None:
        return None
    s = str(s).strip()
    neg = s.startswith("(") and ")" in s
    s = s.replace("(", "").replace(")", "").replace("~", "").replace("+", "")
    token = s.split()[0].rstrip("%")
    try:
        return -float(token) if neg else float(token)
    except ValueError:
        return None

def _rev_range_mid(s) -> float | None:
    """'$10-20M' → 15.0, '$300M+' → 300.0"""
    if pd.isna(s) or s is None:
        return None
    s = str(s).replace("$", "").replace("M", "").strip()
    if s.endswith("+"):
        try: return float(s[:-1])
        except: return None
    if "-" in s:
        try:
            lo, hi = s.split("-")
            return (float(lo) + float(hi)) / 2
        except: return None
    try: return float(s)
    except: return None

def compute_comp_benchmarks(comps: pd.DataFrame) -> dict:
    hi = comps[comps["data_confidence"].str.lower().isin(["high", "medium"])] \
        if "data_confidence" in comps.columns else comps

    def _med(col, df=hi):
        v = df[col].dropna() if col in df.columns else pd.Series(dtype=float)
        return float(v.median()) if not v.empty else None

    return {
        "gross_margin_pct":      _med("gross_margin_pct"),
        "ebitda_margin_pct":     _med("ebitda_margin_pct"),
        "ev_revenue_multiple":   _med("ev_revenue_multiple"),
        "revenue_at_exit_usd_m": _med("revenue_at_exit_usd_m"),
        "n_total":   len(comps),
        "n_hi_conf": len(hi),
    }

def compute_gap_analysis(
    gm_pct: float | None,
    em_pct: float | None,
    bench: dict,
    ltm_rev_usd: float | None,
) -> list[dict]:
    rows: list[dict] = []

    def _add(label, co_val, med, ahead_t, behind_t, fmt):
        if co_val is None or med is None:
            rows.append(dict(label=label, company_val=co_val, comp_median=med,
                             delta=None, status="no_data", fmt=fmt))
            return
        delta = co_val - med
        status = "ahead" if delta >= ahead_t else "behind" if delta <= behind_t else "on_track"
        rows.append(dict(label=label, company_val=co_val, comp_median=med,
                         delta=delta, status=status, fmt=fmt))

    _add("Gross Margin",  gm_pct, bench.get("gross_margin_pct"),  5.0, -5.0,  "pct")
    _add("EBITDA Margin", em_pct, bench.get("ebitda_margin_pct"), 5.0, -10.0, "pct")

    rev_m    = ltm_rev_usd / 1e6 if ltm_rev_usd else None
    comp_rev = bench.get("revenue_at_exit_usd_m")
    if rev_m is not None and comp_rev is not None and comp_rev > 0:
        rows.append(dict(
            label="Revenue vs Comp Exit Scale",
            company_val=rev_m, comp_median=comp_rev,
            delta=rev_m / comp_rev * 100,
            status="scale", fmt="usd_m",
        ))
    return rows

# ── Data quality flags ─────────────────────────────────────────────────────────
def compute_data_quality_flags(
    companies: pd.DataFrame,
    ltm: pd.DataFrame,
    all_rev: pd.DataFrame,
) -> dict:
    """Returns {company_id: [flag_string, ...]}."""
    TODAY = pd.Timestamp("2026-05-05")
    STALE_CUTOFF = TODAY - pd.DateOffset(months=6)   # before 2025-11-05

    flags: dict[int, list[str]] = {int(r["id"]): [] for _, r in companies.iterrows()}

    for _, row in companies.iterrows():
        cid = int(row["id"])

        # DATA STALE
        last_date = pd.to_datetime(row.get("period_end_date")) if not _is_null(row.get("period_end_date")) else None
        if last_date is None or last_date < STALE_CUTOFF:
            flags[cid].append(f"DATA STALE (last: {as_of(row.get('period_end_date'))})")

        # NEGATIVE GROSS MARGIN
        gm = row.get("gross_margin_pct")
        if not _is_null(gm) and float(gm) < 0:
            flags[cid].append(f"CHECK: NEGATIVE MARGIN ({fmt_pct(gm)})")

        # UNUSUALLY HIGH GROSS MARGIN
        if not _is_null(gm) and float(gm) > 95:
            flags[cid].append(f"CHECK: UNUSUALLY HIGH MARGIN ({fmt_pct(gm)})")

        # EXTREME EBITDA BURN
        em = row.get("ebitda_margin_pct")
        if not _is_null(em) and float(em) < -200:
            flags[cid].append(f"CHECK: EXTREME BURN ({fmt_pct(em)})")

    # DATA INCOMPLETE (fewer than 6 months of revenue data)
    for _, row in ltm.iterrows():
        cid = int(row["id"])
        pt   = row["period_type"]
        used = int(row["ltm_periods_used"])
        # Convert periods to months for threshold
        months_equiv = used * (3 if pt == "quarterly" else 12 if pt == "annual" else 1)
        if months_equiv < 6 and row["ltm_label"] != "LTM":
            flags[cid].append("DATA INCOMPLETE (<6 mo. of revenue)")

    # REVENUE VOLATILITY (any consecutive period change > 80%)
    for cid in companies["id"]:
        cid = int(cid)
        crev = (
            all_rev[all_rev["id"] == cid]
            .sort_values("period_end_date")
            .copy()
        )
        crev = crev[crev["revenue_usd"] > 0]
        if len(crev) >= 2:
            pct_changes = crev["revenue_usd"].pct_change().abs().dropna()
            if (pct_changes > 0.8).any():
                max_swing = pct_changes.max() * 100
                flags[cid].append(f"CHECK: REVENUE VOLATILITY ({max_swing:.0f}% max swing)")

    return flags

# ── Chart factory ──────────────────────────────────────────────────────────────
def line_chart(
    df: pd.DataFrame,
    y_col: str,
    title: str,
    y_fmt: str = "number",
    fill: bool = True,
) -> go.Figure | None:
    sub = df[["period_end_date", y_col]].dropna()
    if len(sub) < 2:
        return None

    hover = (
        "$%{y:,.0f}" if y_fmt == "usd" else
        "%{y:.1f}%"  if y_fmt == "pct" else
        "%{y:,.0f}"
    )

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=sub["period_end_date"],
        y=sub[y_col],
        mode="lines+markers",
        line=dict(color=BLACK, width=2),
        marker=dict(size=5, color=BLACK, line=dict(width=1.5, color=WHITE)),
        fill="tozeroy" if fill else "none",
        fillcolor="rgba(213,250,148,0.20)" if fill else None,
        hovertemplate=f"%{{x|%b %Y}}<br>{hover}<extra></extra>",
    ))

    fig.update_layout(
        title=dict(text=title, font=dict(color=BLACK, size=13), x=0, pad=dict(l=4)),
        plot_bgcolor=WHITE,
        paper_bgcolor=BG,
        font=dict(color=BLACK, size=11),
        xaxis=dict(showgrid=False, tickformat="%b %Y", tickfont=dict(size=10), linecolor=BORDER),
        yaxis=dict(
            showgrid=True, gridcolor="#EBEBE6",
            ticksuffix="%" if y_fmt == "pct" else "",
            tickprefix="$" if y_fmt == "usd" else "",
            tickfont=dict(size=10),
            zeroline=True, zerolinecolor=BORDER, zerolinewidth=1,
        ),
        margin=dict(l=8, r=8, t=40, b=8),
        height=260,
        hovermode="x unified",
        showlegend=False,
    )
    return fig

def _no_data_box(msg: str = "No data") -> None:
    st.markdown(
        f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
        f"padding:40px;text-align:center;color:{MUTED};font-size:13px'>{msg}</div>",
        unsafe_allow_html=True,
    )

# ── Benchmarking tab renderer ─────────────────────────────────────────────────
def render_benchmarking_tab(
    info: pd.Series,
    kpis: pd.DataFrame,
    ltm_val: float | None,
    ltm_lbl: str,
    ltm_gm_pct: float | None = None,
    ltm_em_pct: float | None = None,
) -> None:
    company_name = info["name"]
    comp_mapping = load_comp_mapping(company_name)

    if comp_mapping.empty:
        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:40px;text-align:center;color:{MUTED};font-size:14px;line-height:2'>"
            f"No comp set mapped yet for <b style='color:{BLACK}'>{company_name}</b>.<br>"
            f"<small>Sector: <b>{sector_label(info.get('sector',''))}</b>"
            f" &nbsp;·&nbsp; Sub-sector: "
            f"<b>{(info.get('sub_sector') or '—').replace('_',' ').title()}</b></small>"
            f"</div>",
            unsafe_allow_html=True,
        )
        return

    comp_ids = tuple(comp_mapping["comp_id"].tolist())
    comps    = load_comps_detail(comp_ids)
    if comps.empty:
        st.info("Comp data not available.")
        return

    comps = (
        comps.merge(comp_mapping, on="comp_id", how="left")
             .sort_values("relevance_score", ascending=False)
             .reset_index(drop=True)
    )

    bench = compute_comp_benchmarks(comps)
    gaps  = compute_gap_analysis(ltm_gm_pct, ltm_em_pct, bench, ltm_val)

    # ── ARR estimation disclaimer ────────────────────────────────────────────
    if ltm_lbl == "ARR (est.)":
        st.markdown(
            f"<div style='background:{WARN_BG};border:1px solid {WARN};border-radius:8px;"
            f"padding:10px 14px;font-size:12px;color:{WARN};margin-bottom:12px'>"
            f"<b>Note:</b> LTM revenue is estimated from ARR — benchmarking comparisons "
            f"(including the implied exit value below) should be treated as "
            f"<b>directional only</b>.</div>",
            unsafe_allow_html=True,
        )

    # ── Comp overview metrics ────────────────────────────────────────────────
    n_total  = bench["n_total"]
    n_hi     = bench["n_hi_conf"]
    comp_rev = bench.get("revenue_at_exit_usd_m")
    comp_gm  = bench.get("gross_margin_pct")
    comp_em  = bench.get("ebitda_margin_pct")
    comp_ev  = bench.get("ev_revenue_multiple")

    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Comps in Set",         f"{n_total}  ({n_hi} high-conf.)")
    b2.metric("Median Rev at Exit",   fmt_usd((comp_rev or 0) * 1e6) if comp_rev else "—")
    b3.metric("Median Gross Margin",  fmt_pct(comp_gm))
    b4.metric("Median EBITDA Margin", fmt_pct(comp_em))

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gap analysis cards ───────────────────────────────────────────────────
    STATUS_CFG = {
        "ahead":    ("#E8F5E9", "#2E7D32", "AHEAD"),
        "on_track": ("#E3F2FD", "#1565C0", "ON TRACK"),
        "behind":   (WARN_BG,   WARN,       "BEHIND"),
        "no_data":  ("#F5F5F5", MUTED,      "NO DATA"),
        "scale":    ("#F3E5F5", "#6A1B9A",  "SCALE"),
    }

    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin-bottom:10px'>Performance vs. Comp Medians</div>",
        unsafe_allow_html=True,
    )

    gap_cols = st.columns(max(len(gaps), 1))
    for col, g in zip(gap_cols, gaps):
        bg, fg, badge = STATUS_CFG.get(g["status"], STATUS_CFG["no_data"])
        co_val  = g["company_val"]
        med_val = g["comp_median"]
        delta   = g["delta"]
        fmt     = g["fmt"]

        if fmt == "pct":
            co_str  = fmt_pct(co_val)
            med_str = fmt_pct(med_val)
            delta_str = (
                f"+{delta:.1f}pp" if delta is not None and delta >= 0
                else f"{delta:.1f}pp" if delta is not None
                else "—"
            )
        elif fmt == "usd_m":
            co_str    = f"${co_val:.1f}M"  if co_val  is not None else "—"
            med_str   = f"${med_val:.1f}M" if med_val is not None else "—"
            delta_str = f"{delta:.0f}% of comp exit scale" if delta is not None else "—"
        else:
            co_str = med_str = delta_str = "—"

        col.markdown(f"""
<div style="background:{WHITE};border:1px solid {BORDER};border-radius:10px;padding:16px 18px;">
  <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;
              color:{MUTED};margin-bottom:6px">{g['label']}</div>
  <div style="display:flex;align-items:baseline;gap:10px;margin-bottom:8px;flex-wrap:wrap">
    <span style="font-size:22px;font-weight:700;color:{BLACK}">{co_str}</span>
    <span style="font-size:12px;color:{MUTED}">vs {med_str} median</span>
  </div>
  <span style="background:{bg};color:{fg};border-radius:4px;
               padding:2px 8px;font-size:11px;font-weight:600">{badge}</span>
  <span style="font-size:11px;color:{MUTED};margin-left:6px">{delta_str}</span>
</div>
""", unsafe_allow_html=True)

    # ── EV multiple context ──────────────────────────────────────────────────
    if comp_ev is not None and ltm_val is not None:
        implied_ev = ltm_val * comp_ev
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:14px 18px;font-size:13px;color:{BLACK}'>"
            f"<span style='font-size:11px;text-transform:uppercase;letter-spacing:.5px;"
            f"color:{MUTED};font-weight:600'>Exit Multiple Context &nbsp; </span>"
            f"At the comp median <b>{comp_ev:.1f}x EV/Revenue</b>, "
            f"{company_name}'s {ltm_lbl} revenue ({fmt_usd(ltm_val)}) "
            f"implies an enterprise value of <b>{fmt_usd(implied_ev)}</b>."
            f"</div>",
            unsafe_allow_html=True,
        )

    # ── Stage snapshots at similar revenue ───────────────────────────────────
    snapshots = load_stage_snapshots(comp_ids)
    if not snapshots.empty and ltm_val is not None:
        ltm_m = ltm_val / 1e6
        snap  = snapshots.copy()
        snap["rev_mid"]  = snap["revenue_range_usd_m"].apply(_rev_range_mid)
        snap["rev_dist"] = (snap["rev_mid"] - ltm_m).abs()
        closest = snap.nsmallest(3, "rev_dist")
        if not closest.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(
                f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                f"color:{MUTED};font-weight:600;margin-bottom:10px'>"
                f"Stage Snapshots at Similar Revenue &nbsp;"
                f"<span style='font-size:11px;font-weight:400;text-transform:none;"
                f"letter-spacing:0;color:{MUTED}'>({company_name} LTM: {fmt_usd(ltm_val)})</span>"
                f"</div>",
                unsafe_allow_html=True,
            )
            disp = closest[[
                "company_name", "stage", "revenue_range_usd_m",
                "revenue_growth_pct", "gross_margin_pct", "ebitda_margin_pct",
            ]].rename(columns={
                "company_name":        "Comp",
                "stage":               "Stage",
                "revenue_range_usd_m": "Rev Range",
                "revenue_growth_pct":  "Rev Growth",
                "gross_margin_pct":    "Gross Margin",
                "ebitda_margin_pct":   "EBITDA Margin",
            }).reset_index(drop=True)
            st.dataframe(disp, use_container_width=True, hide_index=True)

    # ── Peer comp set table ──────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin-bottom:10px'>Peer Comp Set</div>",
        unsafe_allow_html=True,
    )

    REL_COLORS = {
        5: (GREEN,     BLACK),
        4: (BLUE,      BLACK),
        3: ("#FFF9C4", "#795548"),
        2: ("#EFEBE9", MUTED),
        1: ("#F5F5F5", MUTED),
    }

    hdr_html = "".join(
        f"<th style='padding:6px 10px;text-align:left;font-size:10px;"
        f"text-transform:uppercase;letter-spacing:.5px;color:{MUTED};"
        f"border-bottom:2px solid {BORDER};white-space:nowrap'>{h}</th>"
        for h in ["Company", "Sub-sector", "Geography", "Rev at Exit",
                  "Gross Margin", "EBITDA Margin", "EV/Rev", "Relevance", "Confidence"]
    )

    rows_html = ""
    for _, row in comps.iterrows():
        rel           = int(row["relevance_score"]) if not _is_null(row.get("relevance_score")) else 0
        bg_r, fg_r    = REL_COLORS.get(rel, ("#F5F5F5", MUTED))
        rev           = row.get("revenue_at_exit_usd_m")
        gm            = row.get("gross_margin_pct")
        em            = row.get("ebitda_margin_pct")
        ev            = row.get("ev_revenue_multiple")
        conf          = str(row.get("data_confidence", "—")).capitalize()
        rows_html += (
            f"<tr style='border-bottom:1px solid {BORDER}'>"
            f"<td style='padding:8px 10px;font-weight:600;color:{BLACK}'>{row['company_name']}</td>"
            f"<td style='padding:8px 10px;font-size:12px;color:{MUTED};max-width:160px'>{row.get('sub_sector','—')}</td>"
            f"<td style='padding:8px 10px;font-size:12px;color:{MUTED}'>{row.get('geography','—')}</td>"
            f"<td style='padding:8px 10px;font-weight:500'>{'$'+str(round(rev))+'M' if not _is_null(rev) else '—'}</td>"
            f"<td style='padding:8px 10px'>{fmt_pct(gm)}</td>"
            f"<td style='padding:8px 10px'>{fmt_pct(em)}</td>"
            f"<td style='padding:8px 10px'>{f'{ev:.1f}x' if not _is_null(ev) else '—'}</td>"
            f"<td style='padding:8px 10px'>"
            f"<span style='background:{bg_r};color:{fg_r};border-radius:4px;"
            f"padding:2px 7px;font-size:11px;font-weight:600'>{rel}/5</span></td>"
            f"<td style='padding:8px 10px;font-size:11px;color:{MUTED}'>{conf}</td>"
            f"</tr>"
        )

    st.markdown(
        f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;overflow:auto'>"
        f"<table style='width:100%;border-collapse:collapse'>"
        f"<thead><tr style='background:{BG}'>{hdr_html}</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        f"</table></div>",
        unsafe_allow_html=True,
    )

    # ── Mapping rationale expander ───────────────────────────────────────────
    with st.expander("Why these comps? (mapping rationale)"):
        for _, row in comps.iterrows():
            rationale = row.get("mapping_rationale", "")
            if not _is_null(rationale):
                st.markdown(
                    f"**{row['company_name']}** ({row.get('relevance_score','?')}/5) — {rationale}"
                )


# ── DB write helpers ──────────────────────────────────────────────────────────

def _existing_periods(company_id: int) -> set[str]:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    rows = conn.execute(
        "SELECT period_end_date FROM kpi_snapshots WHERE company_id = ?",
        (company_id,),
    ).fetchall()
    conn.close()
    return {r[0] for r in rows}


def _upsert_kpi(company_id: int, data: dict) -> None:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    now    = datetime.utcnow().isoformat()
    period = data["period_end_date"]

    existing = conn.execute(
        "SELECT id FROM kpi_snapshots WHERE company_id=? AND period_end_date=?",
        (company_id, period),
    ).fetchone()

    if existing:
        update_cols = {
            k: v for k, v in data.items()
            if k != "period_end_date" and v is not None
        }
        if update_cols:
            set_clause = ", ".join(f"{k}=?" for k in update_cols)
            conn.execute(
                f"UPDATE kpi_snapshots SET {set_clause}, updated_at=? "
                f"WHERE company_id=? AND period_end_date=?",
                [*update_cols.values(), now, company_id, period],
            )
    else:
        row = {"company_id": company_id, "created_at": now, "updated_at": now,
               **{k: v for k, v in data.items() if v is not None}}
        cols_str    = ", ".join(row.keys())
        placeholders = ", ".join("?" * len(row))
        conn.execute(
            f"INSERT INTO kpi_snapshots ({cols_str}) VALUES ({placeholders})",
            list(row.values()),
        )

    conn.commit()
    conn.close()


def _generate_commentary(
    company_name: str,
    sector: str,
    new_periods: list[dict],
) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return "Commentary unavailable — set the ANTHROPIC_API_KEY environment variable."

    # Comp benchmarks (best-effort)
    bench_txt = ""
    try:
        mapping = load_comp_mapping(company_name)
        if not mapping.empty:
            comps = load_comps_detail(tuple(mapping["comp_id"].tolist()))
            if not comps.empty:
                b = compute_comp_benchmarks(comps)
                bench_txt = (
                    f"\n\nComp set benchmarks (medians, {b['n_total']} exit comps): "
                    f"Gross Margin {fmt_pct(b.get('gross_margin_pct'))}, "
                    f"EBITDA Margin {fmt_pct(b.get('ebitda_margin_pct'))}, "
                    f"Revenue at Exit {fmt_usd((b.get('revenue_at_exit_usd_m') or 0) * 1e6)}, "
                    f"EV/Revenue {b.get('ev_revenue_multiple') or '—'}x."
                )
    except Exception:
        pass

    # Per-period summary lines
    lines = []
    for p in sorted(new_periods, key=lambda x: x["period_end_date"]):
        rev = p.get("revenue_usd")
        gm  = p.get("gross_margin_pct")
        ebt = p.get("ebitda_usd")
        em  = p.get("ebitda_margin_pct") or (
            round(ebt / rev * 100, 1) if (ebt and rev) else None
        )
        parts = [f"Revenue {fmt_usd(rev)}" if rev else "Revenue N/A"]
        if gm is not None:
            parts.append(f"Gross Margin {fmt_pct(gm)}")
        if em is not None:
            parts.append(f"EBITDA Margin {fmt_pct(em)}")
        lines.append(f"  {p['period_end_date']}: {', '.join(parts)}")

    prompt = (
        f"You are an investment analyst at Quona Capital, a fintech-focused VC firm.\n"
        f"Company: {company_name} | Sector: {sector.replace('_', ' ').title()}\n\n"
        f"Newly uploaded performance data:\n" + "\n".join(lines) +
        bench_txt +
        "\n\nWrite a concise 3-4 sentence analyst commentary on this performance update. "
        "Reference specific numbers, compare to comp benchmarks where data is available, "
        "and highlight key trends or concerns. Use professional third-person style."
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text
    except Exception as exc:
        return f"Commentary generation failed: {exc}"


# ── Upload tab renderer ────────────────────────────────────────────────────────

def _build_preview_df(rows: list[dict]) -> pd.DataFrame:
    """Build a display DataFrame from a list of parsed period dicts."""
    preview_rows = []
    for p in sorted(rows, key=lambda x: x["period_end_date"]):
        rev = p.get("revenue_usd")
        gm  = p.get("gross_margin_pct")
        ebt = p.get("ebitda_usd")
        em  = p.get("ebitda_margin_pct") or (
            round(ebt / rev * 100, 1) if (ebt and rev) else None
        )
        row = {
            "Period":        p["period_end_date"],
            "Revenue (USD)": fmt_usd(rev),
            "Gross Margin":  fmt_pct(gm),
            "EBITDA Margin": fmt_pct(em),
        }
        if p.get("tpv_usd")             is not None: row["TPV (USD)"]     = fmt_usd(p["tpv_usd"])
        if p.get("loan_book_gross_usd") is not None: row["Loan Book"]     = fmt_usd(p["loan_book_gross_usd"])
        if p.get("net_yield_pct")       is not None: row["Net Yield"]     = fmt_pct(p["net_yield_pct"])
        if p.get("par_30_pct")          is not None: row["PAR 30"]        = fmt_pct(p["par_30_pct"])
        if p.get("gmv_usd")             is not None: row["GMV (USD)"]     = fmt_usd(p["gmv_usd"])
        if p.get("customer_count")      is not None: row["Customers"]     = fmt_int(p["customer_count"])
        elif p.get("active_clients_count") is not None: row["Active Clients"] = fmt_int(p["active_clients_count"])
        preview_rows.append(row)
    df = pd.DataFrame(preview_rows)
    return df.loc[:, (df != "—").any(axis=0)]


def render_upload_tab(info: pd.Series, company_id: int) -> None:
    company_name = info["name"]

    st.markdown(
        f"<div style='color:{MUTED};font-size:12px;margin-bottom:16px;line-height:1.7'>"
        f"Upload the latest Excel report for <b style='color:{BLACK}'>{company_name}</b>. "
        f"The parser will extract new periods automatically and show a preview "
        f"before writing anything to the database.</div>",
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Select Excel file",
        type=["xlsx"],
        key=f"uploader_{company_id}",
        label_visibility="collapsed",
    )

    # Session-state key names
    ss_fkey       = f"upload_fkey_{company_id}"
    ss_parsed     = f"upload_parsed_{company_id}"
    ss_skip       = f"upload_skip_{company_id}"
    ss_saved      = f"upload_saved_{company_id}"
    ss_snap       = f"upload_snap_{company_id}"   # saved-periods snapshot for success display
    ss_commentary = f"upload_commentary_{company_id}"

    if uploaded is None:
        for k in (ss_fkey, ss_parsed, ss_skip, ss_saved, ss_snap, ss_commentary):
            st.session_state.pop(k, None)
        return

    file_key = f"{uploaded.name}_{uploaded.size}"

    # ── SUCCESS STATE ─────────────────────────────────────────────────────────
    # Must be checked BEFORE the parse block so the post-save rerun renders the
    # success state rather than re-parsing the (now-stale) cached file key.
    if st.session_state.get(ss_saved):
        snap       = st.session_state.get(ss_snap, [])
        commentary = st.session_state.get(ss_commentary, "")
        skipped    = st.session_state.get(ss_skip, 0)

        if skipped:
            st.markdown(
                f"<div style='color:{MUTED};font-size:12px;margin-bottom:8px'>"
                f"{skipped} period(s) already in database — skipped.</div>",
                unsafe_allow_html=True,
            )

        st.markdown(
            f"<div style='background:#E8F5E9;border:1px solid #2E7D32;border-radius:8px;"
            f"padding:12px 18px;font-size:13px;color:#2E7D32;font-weight:600;margin-bottom:14px'>"
            f"✓ {len(snap)} period(s) saved. Charts and benchmarking now reflect the "
            f"updated data.</div>",
            unsafe_allow_html=True,
        )

        if commentary:
            st.markdown(
                f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                f"color:{MUTED};font-weight:600;margin-bottom:8px'>AI Performance Commentary</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                f"padding:18px 22px;font-size:13px;line-height:1.85;color:{BLACK}'>"
                f"{commentary}</div>",
                unsafe_allow_html=True,
            )

        # Reset flag so next upload starts fresh (ss_fkey was cleared on save,
        # so re-uploading the same file will trigger a real re-parse+dedup).
        st.session_state[ss_saved] = False
        return

    # ── PARSE ─────────────────────────────────────────────────────────────────
    # Re-parse whenever the file changes.  ss_fkey is cleared after every save,
    # so re-uploading the same file always goes through this block.
    if st.session_state.get(ss_fkey) != file_key:
        with st.spinner("Reading and parsing Excel file…"):
            try:
                file_bytes = uploaded.read()
                all_rows   = PARSERS[company_name](file_bytes)

                existing = _existing_periods(company_id)
                new_rows = [r for r in all_rows if r["period_end_date"] not in existing]
                skipped  = len(all_rows) - len(new_rows)

                st.session_state[ss_fkey]   = file_key
                st.session_state[ss_parsed] = new_rows
                st.session_state[ss_skip]   = skipped
            except Exception as exc:
                st.error(f"Parse error: {exc}")
                return

    # Safety net: re-filter cached results against the live DB in case the
    # same file key is reused after a save (shouldn't happen with the fkey
    # invalidation above, but guards against any edge case).
    cached_rows = st.session_state.get(ss_parsed, [])
    existing    = _existing_periods(company_id)
    new_rows    = [r for r in cached_rows if r["period_end_date"] not in existing]
    if len(new_rows) != len(cached_rows):
        st.session_state[ss_parsed] = new_rows   # keep cache in sync
    skipped = st.session_state.get(ss_skip, 0)

    if skipped:
        st.markdown(
            f"<div style='color:{MUTED};font-size:12px;margin-bottom:8px'>"
            f"{skipped} period(s) already in database — skipped.</div>",
            unsafe_allow_html=True,
        )

    if not new_rows:
        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:28px;text-align:center;color:{MUTED};font-size:13px'>"
            f"All periods in this file are already in the database. No new data to import.</div>",
            unsafe_allow_html=True,
        )
        return

    # ── PREVIEW TABLE ─────────────────────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin-bottom:10px'>"
        f"{len(new_rows)} new period(s) found — review before saving</div>",
        unsafe_allow_html=True,
    )
    st.dataframe(_build_preview_df(new_rows), use_container_width=True, hide_index=True)

    # ── CONFIRM BUTTON ────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    btn_col, note_col = st.columns([1, 3])
    with btn_col:
        confirm = st.button(
            "Confirm & Save to Database",
            key=f"confirm_{company_id}",
        )
    with note_col:
        st.markdown(
            f"<div style='padding-top:8px;font-size:12px;color:{MUTED}'>"
            f"Saves {len(new_rows)} period(s) for {company_name}. "
            f"Existing periods are not overwritten.</div>",
            unsafe_allow_html=True,
        )

    if confirm:
        with st.spinner(f"Saving {len(new_rows)} period(s) to database…"):
            for p in new_rows:
                _upsert_kpi(company_id, p)

        with st.spinner("Generating AI commentary…"):
            commentary = _generate_commentary(
                company_name, str(info.get("sector", "")), new_rows
            )

        st.session_state[ss_snap]       = list(new_rows)   # snapshot for success display
        st.session_state[ss_commentary] = commentary
        st.session_state[ss_saved]      = True
        # Invalidate the file key cache so re-uploading the same file will
        # trigger a fresh parse+dedup (not reuse stale ss_parsed).
        st.session_state.pop(ss_fkey, None)
        st.cache_data.clear()
        st.rerun()


# ── Session state ─────────────────────────────────────────────────────────────
if "page" not in st.session_state:
    st.session_state.page = "home"
    st.session_state.company_id = None

# ── Persistent header ─────────────────────────────────────────────────────────
st.markdown(f"""
<div style="background:{BLACK};border-radius:12px;padding:14px 28px;
            margin-bottom:20px;display:flex;align-items:baseline;gap:12px;">
  <span style="font-size:22px;font-weight:800;color:{GREEN};letter-spacing:-0.5px;">Quona Capital</span>
  <span style="font-size:13px;color:rgba(255,255,255,0.55);">Portfolio Intelligence</span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HOME PAGE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "home":

    companies = load_companies()
    growth    = load_revenue_growth()
    ltm       = load_ltm_revenue()
    all_rev   = load_all_revenue()

    companies = companies.merge(growth, on="id", how="left")
    companies = companies.merge(ltm,    on="id", how="left")

    flags = compute_data_quality_flags(companies, ltm, all_rev)

    # ── Summary KPIs ──────────────────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    ltm_gm_col = companies["ltm_gross_margin_pct"].combine_first(
        companies["gross_margin_pct"]
    )
    ltm_em_col = companies["ltm_ebitda_margin_pct"].combine_first(
        companies["ebitda_margin_pct"]
    )
    col1.metric("Portfolio Companies", len(companies))
    col2.metric("Combined LTM Revenue",
                fmt_usd(companies["ltm_revenue"].sum()))
    col3.metric("Avg Gross Margin",    fmt_pct(ltm_gm_col.mean()))
    col4.metric("Avg EBITDA Margin",   fmt_pct(ltm_em_col.mean()))

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"### Portfolio Companies &nbsp;"
        f"<small style='color:{MUTED};font-size:13px;font-weight:400'>Click a company to explore</small>",
        unsafe_allow_html=True,
    )

    # ── Column headers ────────────────────────────────────────────────────────
    COLS = [2.5, 1.8, 1.7, 1.3, 1.3, 1.3, 1.2]
    HDRS = ["Company", "Sector / HQ", "LTM Revenue (USD)",
            "Gross Margin", "EBITDA Margin", "Rev Growth", "As of"]

    h = st.columns(COLS)
    for col, lbl in zip(h, HDRS):
        col.markdown(
            f"<div style='font-size:11px;text-transform:uppercase;"
            f"letter-spacing:.6px;color:{MUTED};padding-bottom:4px'>{lbl}</div>",
            unsafe_allow_html=True,
        )
    st.markdown(f"<hr style='margin-top:0'>", unsafe_allow_html=True)

    # ── Company rows ─────────────────────────────────────────────────────────
    for _, row in companies.iterrows():
        r  = st.columns(COLS)
        sl = sector_label(row["sector"])
        cid = int(row["id"])
        company_flags = flags.get(cid, [])

        # Company name + inline flag badges
        with r[0]:
            if st.button(row["name"], key=f"co_{row['id']}"):
                st.session_state.page = "detail"
                st.session_state.company_id = cid
                st.rerun()
            if company_flags:
                # Show first flag as a small badge; more flags in the bottom section
                badge = company_flags[0].split("(")[0].strip()
                more  = f" +{len(company_flags)-1}" if len(company_flags) > 1 else ""
                st.markdown(
                    f"<div style='margin-top:-6px'>"
                    f"<span style='background:{WARN_BG};color:{WARN};border-radius:4px;"
                    f"padding:1px 6px;font-size:10px;font-weight:600'>"
                    f"! {badge}{more}</span></div>",
                    unsafe_allow_html=True,
                )

        with r[1]:
            st.markdown(
                f"<span style='background:{BLUE};border-radius:20px;padding:2px 10px;"
                f"font-size:11px;font-weight:500;color:{BLACK}'>{sl}</span>"
                f"&nbsp;<small style='color:{MUTED}'>{row['hq_country']}</small>",
                unsafe_allow_html=True,
            )

        with r[2]:
            ltm_val   = row.get("ltm_revenue")
            ltm_lbl   = row.get("ltm_label", "")
            pt        = row.get("period_type", "monthly")
            n_used    = int(row.get("ltm_periods_used", 0))
            n_needed  = int(row.get("periods_needed", 12))

            if ltm_lbl == "LTM":
                basis = (
                    "12 mo." if pt == "monthly"   else
                    "4 qtrs." if pt == "quarterly" else
                    "annual"
                )
                sub = f"LTM &middot; {basis}"
            elif ltm_lbl == "ARR (est.)":
                sub = f"ARR est. &middot; {n_used} of {n_needed}"
            else:
                sub = ""

            basis_tag = (
                f"<div style='font-size:10px;color:{MUTED};margin-top:1px'>({sub})</div>"
                if (not _is_null(ltm_val) and sub) else ""
            )
            period_lbl = fmt_period_label(row.get("period_end_date"), pt)
            period_sfx = (
                f"<span style='font-size:13px;color:{MUTED};font-weight:400'> ({period_lbl})</span>"
                if (period_lbl and not _is_null(ltm_val)) else ""
            )
            st.markdown(
                f"<span style='font-weight:600'>{fmt_usd(ltm_val)}</span>{period_sfx}{basis_tag}",
                unsafe_allow_html=True,
            )

        with r[3]:
            gm = row.get("ltm_gross_margin_pct")
            if _is_null(gm):
                gm = row["gross_margin_pct"]
            color = "#2E7D32" if (not _is_null(gm) and float(gm) > 50) else BLACK
            st.markdown(
                f"<span style='color:{color};font-weight:500'>{fmt_pct(gm)}</span>",
                unsafe_allow_html=True,
            )

        with r[4]:
            em = row.get("ltm_ebitda_margin_pct")
            if _is_null(em):
                em = row.get("ebitda_margin_pct")
            color = ("#2E7D32" if (not _is_null(em) and float(em) > 0)
                     else "#C62828" if (not _is_null(em))
                     else BLACK)
            st.markdown(
                f"<span style='color:{color};font-weight:500'>{fmt_pct(em)}</span>",
                unsafe_allow_html=True,
            )

        with r[5]:
            gtxt, gcol = fmt_growth(row.get("revenue_growth_pct"))
            st.markdown(
                f"<span style='color:{gcol};font-weight:500'>{gtxt}</span>",
                unsafe_allow_html=True,
            )

        with r[6]:
            st.markdown(
                f"<small style='color:{MUTED}'>{as_of(row['period_end_date'])}</small>",
                unsafe_allow_html=True,
            )

        st.markdown(f"<hr style='margin:0;border-color:{BORDER}'>", unsafe_allow_html=True)

    # ── Methodology note ──────────────────────────────────────────────────────
    st.markdown(
        f"<div style='color:{MUTED};font-size:11px;margin-top:10px;"
        f"background:{WHITE};border:1px solid {BORDER};border-radius:8px;"
        f"padding:10px 14px;line-height:1.6'>"
        f"<b>LTM Revenue</b> &mdash; Last Twelve Months: sum of the most recent 12 monthly periods "
        f"(or 4 quarterly / 1 annual) so all companies are on a comparable full-year basis. "
        f"Companies with fewer than 12 months of history show an annualised run-rate labelled "
        f"<i>ARR (est.)</i>. Reported figures (not annualised) are shown on the company detail page."
        f"<br><b>Rev Growth</b> &mdash; period-over-period change between the two most recent "
        f"available revenue data points."
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── Data quality flags section ────────────────────────────────────────────
    all_flags = {cid: fl for cid, fl in flags.items() if fl}
    if all_flags:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f"<div style='font-size:13px;font-weight:700;color:{BLACK};"
            f"letter-spacing:.3px;margin-bottom:6px'>DATA QUALITY FLAGS</div>",
            unsafe_allow_html=True,
        )
        name_map = {int(r["id"]): r["name"] for _, r in companies.iterrows()}

        rows_html = ""
        for cid, fl_list in sorted(all_flags.items(), key=lambda x: name_map.get(x[0], "")):
            company_name = name_map.get(cid, f"ID {cid}")
            badges = " ".join(
                f"<span style='background:{WARN_BG};color:{WARN};border-radius:4px;"
                f"padding:2px 8px;font-size:11px;font-weight:600;margin-right:4px'>{f}</span>"
                for f in fl_list
            )
            rows_html += (
                f"<tr>"
                f"<td style='padding:6px 12px 6px 0;font-weight:600;white-space:nowrap;"
                f"color:{BLACK};width:140px'>{company_name}</td>"
                f"<td style='padding:6px 0'>{badges}</td>"
                f"</tr>"
            )

        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:8px;"
            f"padding:12px 16px'>"
            f"<table style='width:100%;border-collapse:collapse'>{rows_html}</table>"
            f"</div>",
            unsafe_allow_html=True,
        )

    # ── LTM summary table (printed to console; also shown as expander) ────────
    with st.expander("LTM Revenue & data quality summary (all companies)"):
        summary_rows = []
        for _, row in companies.iterrows():
            cid = int(row["id"])
            fl  = flags.get(cid, [])
            summary_rows.append({
                "Company":     row["name"],
                "LTM Revenue": fmt_usd(row.get("ltm_revenue")),
                "Basis":       row.get("ltm_label", "—"),
                "Period type": row.get("period_type", "—"),
                "Gross Margin (LTM)": fmt_pct(row.get("ltm_gross_margin_pct") or row.get("gross_margin_pct")),
                "EBITDA Margin (LTM)": fmt_pct(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct")),
                "As of":       as_of(row.get("period_end_date")),
                "Flags":       "; ".join(fl) if fl else "OK",
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# DETAIL PAGE
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "detail":

    if st.button("← Back to Portfolio"):
        st.session_state.page = "home"
        st.session_state.company_id = None
        st.rerun()

    company_id = st.session_state.company_id
    info = load_company_info(company_id)
    kpis = load_kpis(company_id)

    # LTM for this company
    ltm_df  = load_ltm_revenue()
    ltm_row = ltm_df[ltm_df["id"] == company_id]
    ltm_val = float(ltm_row.iloc[0]["ltm_revenue"]) if not ltm_row.empty and not _is_null(ltm_row.iloc[0]["ltm_revenue"]) else None
    ltm_lbl = ltm_row.iloc[0]["ltm_label"] if not ltm_row.empty else "—"

    sl      = sector_label(info["sector"])
    founded = (
        f"· Est. {int(info['founded_year'])}"
        if not _is_null(info.get("founded_year"))
        else ""
    )

    # ── Company header card ────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:{WHITE};border:1px solid {BORDER};border-radius:12px;
                padding:22px 28px;margin-bottom:20px;">
      <div style="display:flex;align-items:center;gap:16px;">
        <div style="background:{GREEN};border-radius:10px;width:52px;height:52px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:22px;font-weight:800;color:{BLACK};flex-shrink:0;">
          {info['name'][0]}
        </div>
        <div>
          <div style="font-size:26px;font-weight:800;color:{BLACK};line-height:1.1">{info['name']}</div>
          <div style="margin-top:6px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
            <span style="background:{BLUE};border-radius:20px;padding:3px 12px;
                         font-size:12px;font-weight:500;color:{BLACK}">{sl}</span>
            <span style="color:{MUTED};font-size:13px">{info['hq_country']} {founded}</span>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    if kpis.empty:
        st.info("No KPI data available for this company yet.")
        st.stop()

    # ── Summary metrics ────────────────────────────────────────────────────
    latest    = kpis.iloc[-1]
    customers = latest.get("customer_count")
    if _is_null(customers):
        customers = latest.get("active_clients_count")

    date_range = (
        f"{kpis['period_end_date'].min().strftime('%b %Y')} – "
        f"{kpis['period_end_date'].max().strftime('%b %Y')}"
    )

    ltm_em_pct = (
        float(ltm_row.iloc[0]["ltm_ebitda_margin_pct"])
        if not ltm_row.empty and not _is_null(ltm_row.iloc[0].get("ltm_ebitda_margin_pct"))
        else None
    )
    ltm_gm_pct = (
        float(ltm_row.iloc[0]["ltm_gross_margin_pct"])
        if not ltm_row.empty and not _is_null(ltm_row.iloc[0].get("ltm_gross_margin_pct"))
        else None
    )
    ebitda_margin_display = ltm_em_pct if ltm_em_pct is not None else (
        float(latest.get("ebitda_margin_pct"))
        if not _is_null(latest.get("ebitda_margin_pct")) else None
    )
    ebitda_margin_label = f"{ltm_lbl} EBITDA Margin" if ltm_em_pct is not None else "EBITDA Margin"
    gm_display = ltm_gm_pct if ltm_gm_pct is not None else (
        float(latest.get("gross_margin_pct"))
        if not _is_null(latest.get("gross_margin_pct")) else None
    )
    gm_label = f"{ltm_lbl} Gross Margin" if ltm_gm_pct is not None else "Gross Margin"

    latest_pt  = ltm_row.iloc[0]["period_type"] if not ltm_row.empty else "monthly"
    latest_plbl = fmt_period_label(latest.get("period_end_date"), latest_pt)
    latest_rev_display = (
        f"{fmt_usd(latest.get('revenue_usd'))} ({latest_plbl})"
        if latest_plbl and not _is_null(latest.get("revenue_usd"))
        else fmt_usd(latest.get("revenue_usd"))
    )

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric(f"{ltm_lbl} Revenue",         fmt_usd(ltm_val))
    m2.metric("Revenue (latest)",            latest_rev_display)
    m3.metric(gm_label,                      fmt_pct(gm_display))
    m4.metric(ebitda_margin_label,           fmt_pct(ebitda_margin_display))
    m5.metric("Customers / Clients",         fmt_int(customers))
    m6.metric("History", f"{len(kpis)} periods  ·  {date_range}")

    _has_upload = info["name"] in SUPPORTED_COMPANIES
    _tab_names  = ["Performance", "Benchmarking"] + (["Upload Data"] if _has_upload else [])
    _tabs       = st.tabs(_tab_names)
    tab_perf    = _tabs[0]
    tab_bench   = _tabs[1]
    tab_upload  = _tabs[2] if _has_upload else None

    with tab_perf:
        # ── Financial performance charts ─────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Financial Performance")

        c1, c2 = st.columns(2)
        with c1:
            fig = line_chart(kpis, "revenue_usd", "Revenue (USD)", y_fmt="usd", fill=True)
            if fig: st.plotly_chart(fig, use_container_width=True)
            else:   _no_data_box("No revenue data")

        with c2:
            fig = line_chart(kpis, "gross_margin_pct", "Gross Margin %", y_fmt="pct", fill=False)
            if fig: st.plotly_chart(fig, use_container_width=True)
            else:   _no_data_box("No gross margin data")

        c3, c4 = st.columns(2)
        with c3:
            fig = line_chart(kpis, "ebitda_usd", "EBITDA (USD)", y_fmt="usd", fill=True)
            if fig: st.plotly_chart(fig, use_container_width=True)
            else:   _no_data_box("No EBITDA data")

        with c4:
            fig = line_chart(kpis, "ebitda_margin_pct", "EBITDA Margin %", y_fmt="pct", fill=False)
            if fig: st.plotly_chart(fig, use_container_width=True)
            else:   _no_data_box("No EBITDA margin data")

        if kpis["customer_count"].notna().any():
            cust_col, cust_lbl = "customer_count", "Customer Count"
        elif kpis["active_clients_count"].notna().any():
            cust_col, cust_lbl = "active_clients_count", "Active Clients"
        else:
            cust_col, cust_lbl = None, None

        if cust_col:
            c5, c6 = st.columns(2)
            with c5:
                fig = line_chart(kpis, cust_col, cust_lbl, y_fmt="number", fill=True)
                if fig: st.plotly_chart(fig, use_container_width=True)
                else:   _no_data_box()

        # ── Lending metrics ───────────────────────────────────────────────────
        LENDING_SNAPSHOT_METRICS = [
            ("loan_book_gross_usd",    "Net Loan Portfolio",  fmt_usd),
            ("net_yield_pct",          "Avg Interest Rate",   fmt_pct),
            ("par_30_pct",             "PAR 30+",             fmt_pct),
            ("par_90_pct",             "PAR 90",              fmt_pct),
            ("active_clients_count",   "Active Clients",      fmt_int),
            ("unique_borrowers_count", "Unique SMEs Funded",  fmt_int),
        ]
        if info["sector"] == "lending":
            snapshot_vals = [
                (lbl, fn(latest.get(k)))
                for k, lbl, fn in LENDING_SNAPSHOT_METRICS
                if not _is_null(latest.get(k))
            ]
            if snapshot_vals:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("#### Lending KPIs (Latest Period)")
                snap_cols = st.columns(len(snapshot_vals))
                for col, (lbl, val_str) in zip(snap_cols, snapshot_vals):
                    col.metric(lbl, val_str)

        LENDING_METRICS = [
            ("loan_book_gross_usd", "Net Loan Portfolio (USD)", "usd",  True),
            ("par_30_pct",          "PAR 30+ %",                "pct",  False),
            ("par_90_pct",          "PAR 90 %",                 "pct",  False),
            ("npl_rate_pct",        "NPL Rate %",               "pct",  False),
            ("net_yield_pct",       "Net Yield %",              "pct",  False),
            ("nim_pct",             "Net Interest Margin %",    "pct",  False),
        ]
        lending_available = [
            m for m in LENDING_METRICS
            if m[0] in kpis.columns and kpis[m[0]].dropna().__len__() >= 2
        ]
        if lending_available:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("#### Lending & Credit Metrics")
            for i in range(0, len(lending_available), 2):
                cols = st.columns(2)
                for j, col in enumerate(cols):
                    if i + j < len(lending_available):
                        c, ttl, fmt, fill = lending_available[i + j]
                        fig = line_chart(kpis, c, ttl, y_fmt=fmt, fill=fill)
                        if fig:
                            col.plotly_chart(fig, use_container_width=True)

        # ── AUM metrics ───────────────────────────────────────────────────────
        AUM_METRICS = [
            ("aum_usd", "Assets Under Management (USD)", "usd", True),
        ]
        aum_available = [
            m for m in AUM_METRICS
            if m[0] in kpis.columns and kpis[m[0]].dropna().__len__() >= 2
        ]
        if aum_available:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("#### Assets Under Management")
            cols = st.columns(2)
            for j, (c, ttl, fmt, fill) in enumerate(aum_available):
                fig = line_chart(kpis, c, ttl, y_fmt=fmt, fill=fill)
                if fig:
                    cols[j].plotly_chart(fig, use_container_width=True)

        # ── Other sector metrics ──────────────────────────────────────────────
        OTHER_METRICS = [
            ("gmv_usd",                   "GMV (USD)",                    "usd",  True),
            ("tpv_usd",                   "Total Payment Volume (USD)",   "usd",  True),
            ("arr_usd",                   "ARR (USD)",                    "usd",  True),
            ("net_revenue_retention_pct", "Net Revenue Retention %",      "pct",  False),
        ]
        shown = {m[0] for m in lending_available} | {m[0] for m in aum_available}
        other_available = [
            m for m in OTHER_METRICS
            if m[0] not in shown
            and m[0] in kpis.columns
            and kpis[m[0]].dropna().__len__() >= 2
        ]
        if other_available:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("#### Sector Metrics")
            for i in range(0, len(other_available), 2):
                cols = st.columns(2)
                for j, col in enumerate(cols):
                    if i + j < len(other_available):
                        c, ttl, fmt, fill = other_available[i + j]
                        fig = line_chart(kpis, c, ttl, y_fmt=fmt, fill=fill)
                        if fig:
                            col.plotly_chart(fig, use_container_width=True)

        # ── Raw data table ────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("Raw data table"):
            candidate_cols = [
                "period_end_date", "revenue_usd", "gross_margin_pct",
                "ebitda_usd", "ebitda_margin_pct",
                "customer_count", "active_clients_count",
                "arr_usd", "aum_usd", "gmv_usd", "tpv_usd",
                "loan_book_gross_usd", "par_30_pct", "par_90_pct",
                "npl_rate_pct", "net_yield_pct", "unique_borrowers_count",
            ]
            show_cols = [
                c for c in candidate_cols
                if c in kpis.columns and kpis[c].notna().any()
            ]
            st.dataframe(
                kpis[show_cols].sort_values("period_end_date", ascending=False).reset_index(drop=True),
                use_container_width=True,
            )

    with tab_bench:
        render_benchmarking_tab(info, kpis, ltm_val, ltm_lbl, ltm_gm_pct, ltm_em_pct)

    if tab_upload is not None:
        with tab_upload:
            render_upload_tab(info, company_id)
