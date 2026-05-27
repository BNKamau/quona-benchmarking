import streamlit as st
import sqlite3
import pandas as pd
import plotly.graph_objects as go
import anthropic
import os
from datetime import datetime, timedelta, timezone
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

  /* Card grid spacing */
  div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] {{
      gap: 14px !important;
  }}

  div[role="radiogroup"] label p {{ color: #2C2C2A !important; font-weight: 600 !important; }}
</style>
""", unsafe_allow_html=True)

# ── DB helpers ─────────────────────────────────────────────────────────────────
DB_PATH = "benchmarking.db"

def _conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def _init_exit_tables() -> None:
    conn = _conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS exit_pathways (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id         INTEGER NOT NULL,
            pathway_name       TEXT    NOT NULL,
            likelihood         TEXT    DEFAULT 'Exploratory',
            estimated_timeline TEXT,
            notes              TEXT,
            created_at         TEXT,
            updated_at         TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS buyer_tracking (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id         INTEGER NOT NULL,
            acquirer_name      TEXT    NOT NULL,
            acquirer_type      TEXT    DEFAULT 'Strategic',
            relationship_owner TEXT,
            last_contact_date  TEXT,
            status             TEXT    DEFAULT 'Not Started',
            sort_order         INTEGER DEFAULT 0,
            created_at         TEXT,
            updated_at         TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS quarterly_actions (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id        INTEGER NOT NULL,
            quarter           TEXT    NOT NULL,
            planned_actions   TEXT    DEFAULT '',
            completed_actions TEXT    DEFAULT '',
            carry_forward     TEXT    DEFAULT '',
            created_at        TEXT,
            updated_at        TEXT,
            UNIQUE(company_id, quarter)
        )
    """)
    conn.commit()
    conn.close()

_init_exit_tables()

# ── Exit comps DB helpers ──────────────────────────────────────────────────────
COMPS_DB = "data/quona_exit_comps.db"
_COMP_NAME_MAP = {"VertoFX": "Verto FX"}  # benchmarking.db name → portfolio_comp_mapping name

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
               revenue_growth_at_exit,
               COALESCE(is_clean_exit, 1)     AS is_clean_exit,
               COALESCE(use_for_margins, 1)   AS use_for_margins,
               COALESCE(use_for_multiples, 1) AS use_for_multiples
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
        SELECT c.id, c.name, c.sector, c.hq_country, c.founded_year, c.fund,
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
               k.gmv_usd,
               k.tpv_usd,
               k.npl_rate_pct,
               k.par_30_pct,
               k.loan_book_gross_usd
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

    def _subset(flag_col):
        if flag_col in hi.columns:
            return hi[hi[flag_col] == 1]
        return hi

    margins_df   = _subset("use_for_margins")
    multiples_df = _subset("use_for_multiples")

    def _med(col, df):
        v = df[col].dropna() if col in df.columns else pd.Series(dtype=float)
        return float(v.median()) if not v.empty else None

    return {
        "gross_margin_pct":      _med("gross_margin_pct",      margins_df),
        "ebitda_margin_pct":     _med("ebitda_margin_pct",     margins_df),
        "ev_revenue_multiple":   _med("ev_revenue_multiple",   multiples_df),
        "revenue_at_exit_usd_m": _med("revenue_at_exit_usd_m", multiples_df),
        "n_total":   len(comps),
        "n_hi_conf": len(margins_df),
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

# ── Affinity deal-intel scan ──────────────────────────────────────────────────

_MA_KEYWORDS = [
    "acquisition", "acquired", "m&a", "merger", "strategic", "term sheet",
    "due diligence", "exit", "ipo", "valuation", "buyout", "transaction",
    "deal close", "invest", "raise", "series",
]

def fetch_affinity_deal_intel(api_key: str) -> list[dict]:
    import requests
    AUTH   = ("", api_key)
    BASE   = "https://api.affinity.co"
    cutoff = datetime.now(tz=timezone.utc) - timedelta(days=365)

    _person_cache: dict[int, str] = {}

    def _person_name(pid: int) -> str:
        if pid in _person_cache:
            return _person_cache[pid]
        try:
            r = requests.get(f"{BASE}/persons/{pid}", auth=AUTH, timeout=10)
            p = r.json()
            name = f"{p.get('first_name','').strip()} {p.get('last_name','').strip()}".strip()
        except Exception:
            name = str(pid)
        _person_cache[pid] = name or str(pid)
        return _person_cache[pid]

    results  = []
    page_token = None

    while True:
        params: dict = {"limit": 100}
        if page_token:
            params["page_token"] = page_token

        r = requests.get(f"{BASE}/notes", params=params, auth=AUTH, timeout=20)
        r.raise_for_status()
        data  = r.json()
        notes = data if isinstance(data, list) else data.get("notes", [])

        for note in notes:
            raw_date = note.get("created_at") or ""
            if not raw_date:
                continue
            note_dt = datetime.fromisoformat(raw_date)
            if note_dt.tzinfo is None:
                note_dt = note_dt.replace(tzinfo=timezone.utc)
            if note_dt < cutoff:
                continue

            content = (note.get("content") or "").strip()
            content_lower = content.lower()
            matched = [kw for kw in _MA_KEYWORDS if kw in content_lower]
            if not matched:
                continue

            creator_id   = note.get("creator_id")
            creator_name = _person_name(creator_id) if creator_id else "Unknown"

            results.append({
                "date":             note_dt.strftime("%Y-%m-%d"),
                "creator_name":     creator_name,
                "snippet":          content[:200] + ("…" if len(content) > 200 else ""),
                "matched_keywords": matched,
            })

        # Pagination — Affinity uses next_page_token or paging dict
        next_token = (
            data.get("next_page_token")
            if isinstance(data, dict)
            else None
        )
        if not next_token or isinstance(data, list):
            break
        page_token = next_token

    results.sort(key=lambda x: x["date"], reverse=True)
    return results


# ── Benchmarking tab renderer ─────────────────────────────────────────────────
def render_benchmarking_tab(
    info: pd.Series,
    kpis: pd.DataFrame,
    ltm_val: float | None,
    ltm_lbl: str,
    ltm_gm_pct: float | None = None,
    ltm_em_pct: float | None = None,
) -> None:
    import math

    company_name = info["name"]
    company_id   = int(info["id"])
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
             .sort_values(
                 ["is_clean_exit", "relevance_score"],
                 ascending=[False, False],
             )
             .reset_index(drop=True)
    )

    bench = compute_comp_benchmarks(comps)
    gaps  = compute_gap_analysis(ltm_gm_pct, ltm_em_pct, bench, ltm_val)

    n_total  = bench["n_total"]
    n_hi     = bench["n_hi_conf"]
    comp_rev = bench.get("revenue_at_exit_usd_m")
    comp_gm  = bench.get("gross_margin_pct")
    comp_em  = bench.get("ebitda_margin_pct")
    comp_ev  = bench.get("ev_revenue_multiple")
    rev_m    = ltm_val / 1e6 if ltm_val else None

    # ── ARR disclaimer ─────────────────────────────────────────────────────────
    if ltm_lbl == "ARR (est.)":
        st.markdown(
            f"<div style='background:{WARN_BG};border:1px solid {WARN};border-radius:8px;"
            f"padding:10px 14px;font-size:12px;color:{WARN};margin-bottom:16px'>"
            f"<b>Note:</b> LTM revenue is estimated from ARR — benchmarking comparisons "
            f"should be treated as <b>directional only</b>.</div>",
            unsafe_allow_html=True,
        )

    # ── Section 1: Summary stat cards ─────────────────────────────────────────
    def _arrow_sublabel(co_val, med_val, suffix="pp"):
        if _is_null(co_val) or _is_null(med_val):
            return f"<div style='font-size:12px;color:{MUTED};margin-top:5px'>Portfolio: —</div>"
        delta = float(co_val) - float(med_val)
        arrow = "↑" if delta > 0 else ("↓" if delta < 0 else "→")
        col   = "#2E7D32" if delta > 0 else ("#C62828" if delta < 0 else MUTED)
        sign  = "+" if delta > 0 else ""
        return (
            f"<div style='font-size:12px;color:{col};font-weight:600;margin-top:5px'>"
            f"{arrow}&nbsp;{fmt_pct(co_val)}"
            f"&nbsp;<span style='font-weight:400;color:{MUTED}'>({sign}{delta:.1f}{suffix} vs median)</span>"
            f"</div>"
        )

    def _stat_card(label, value_str, sub_html=""):
        return (
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:18px 20px'>"
            f"<div style='font-size:10px;text-transform:uppercase;letter-spacing:.6px;"
            f"color:{MUTED};font-weight:600;margin-bottom:6px'>{label}</div>"
            f"<div style='font-size:24px;font-weight:700;color:{BLACK}'>{value_str}</div>"
            f"{sub_html}"
            f"</div>"
        )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            _stat_card(
                "Comps in Set", str(n_total),
                f"<div style='font-size:12px;color:{MUTED};margin-top:5px'>{n_hi} high-confidence</div>",
            ),
            unsafe_allow_html=True,
        )
    with c2:
        rev_str = fmt_usd(comp_rev * 1e6) if comp_rev else "—"
        st.markdown(_stat_card("Median Exit Revenue", rev_str), unsafe_allow_html=True)
    with c3:
        st.markdown(
            _stat_card("Median Gross Margin", fmt_pct(comp_gm), _arrow_sublabel(ltm_gm_pct, comp_gm)),
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            _stat_card("Median EBITDA Margin", fmt_pct(comp_em), _arrow_sublabel(ltm_em_pct, comp_em)),
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Gap analysis (left) + Radar chart (right) ──────────────────
    col_left, col_right = st.columns(2, gap="large")

    with col_left:
        st.markdown(
            f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
            f"color:{MUTED};font-weight:600;margin-bottom:12px'>Performance vs. Comp Medians</div>",
            unsafe_allow_html=True,
        )
        STATUS_CFG = {
            "ahead":    ("#2E7D32", GREEN,     "#E8F5E9", "AHEAD"),
            "on_track": ("#1565C0", BLUE,      "#E3F2FD", "ON TRACK"),
            "behind":   (WARN,     WARN_BG,    WARN_BG,   "BEHIND"),
            "no_data":  (MUTED,    "#F5F5F5",  "#F5F5F5", "NO DATA"),
            "scale":    ("#6A1B9A", "#F3E5F5", "#F3E5F5", "SCALE"),
        }
        for g in gaps:
            border_c, bar_c, badge_bg, badge_txt = STATUS_CFG.get(g["status"], STATUS_CFG["no_data"])
            co_val  = g["company_val"]
            med_val = g["comp_median"]

            if g["fmt"] == "pct":
                co_str    = fmt_pct(co_val)
                med_str   = fmt_pct(med_val)
                delta_str = (
                    f"+{g['delta']:.1f}pp" if g["delta"] is not None and g["delta"] >= 0
                    else f"{g['delta']:.1f}pp" if g["delta"] is not None
                    else "—"
                )
                if co_val is not None and med_val is not None:
                    ref     = max(abs(med_val), abs(co_val), 1)
                    bar_pct = min(max((co_val + ref) / (ref * 2) * 100, 0), 100)
                else:
                    bar_pct = 0
            elif g["fmt"] == "usd_m":
                co_str    = f"${co_val:.1f}M"  if co_val  is not None else "—"
                med_str   = f"${med_val:.1f}M" if med_val is not None else "—"
                delta_str = f"{g['delta']:.0f}% of comp exit scale" if g["delta"] is not None else "—"
                bar_pct   = min(g["delta"] or 0, 100)
            else:
                co_str = med_str = delta_str = "—"
                bar_pct = 0

            st.markdown(
                f"<div style='border-left:4px solid {border_c};background:{WHITE};"
                f"border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:12px;"
                f"box-shadow:0 1px 3px rgba(0,0,0,.04)'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>"
                f"<span style='font-size:10px;text-transform:uppercase;letter-spacing:.5px;"
                f"color:{MUTED};font-weight:600'>{g['label']}</span>"
                f"<span style='background:{badge_bg};color:{border_c};border-radius:4px;"
                f"padding:2px 8px;font-size:10px;font-weight:700'>{badge_txt}</span>"
                f"</div>"
                f"<div style='display:flex;align-items:baseline;gap:8px;margin-bottom:10px;flex-wrap:wrap'>"
                f"<span style='font-size:20px;font-weight:700;color:{BLACK}'>{co_str}</span>"
                f"<span style='font-size:12px;color:{MUTED}'>vs {med_str} median</span>"
                f"<span style='font-size:12px;color:{border_c};font-weight:600'>{delta_str}</span>"
                f"</div>"
                f"<div style='background:{BG};border-radius:4px;height:6px;overflow:hidden'>"
                f"<div style='background:{border_c};height:6px;width:{bar_pct:.0f}%;border-radius:4px'></div>"
                f"</div>"
                f"<div style='display:flex;justify-content:space-between;margin-top:3px'>"
                f"<span style='font-size:10px;color:{MUTED}'>0</span>"
                f"<span style='font-size:10px;color:{MUTED}'>Comp median</span>"
                f"</div></div>",
                unsafe_allow_html=True,
            )

    with col_right:
        # ── Radar chart ───────────────────────────────────────────────────────
        def _norm(val, lo, hi):
            if val is None:
                return 0
            return max(0.0, min(100.0, (val - lo) / (hi - lo) * 100))

        co_gm_r = _norm(ltm_gm_pct, 0, 80)
        md_gm_r = _norm(comp_gm,    0, 80)

        co_em_r = _norm(ltm_em_pct, -80, 40)
        md_em_r = _norm(comp_em,    -80, 40)

        co_rev_r = (
            min(rev_m / comp_rev * 100, 100) if rev_m and comp_rev and comp_rev > 0 else 0
        )
        md_rev_r = 100.0

        hq = str(info.get("hq_country", "")).lower()
        REGION_KEYS = {
            "kenya":        ["ssa", "africa", "kenya", "east africa"],
            "nigeria":      ["ssa", "africa", "nigeria", "west africa"],
            "south africa": ["ssa", "africa", "south africa"],
            "egypt":        ["mena", "north africa", "egypt"],
            "ghana":        ["ssa", "africa", "ghana"],
            "mexico":       ["latam", "latin america", "mexico"],
            "brazil":       ["latam", "latin america", "brazil"],
            "india":        ["south asia", "india"],
            "indonesia":    ["sea", "southeast asia", "indonesia"],
        }
        region_keys = next((v for k, v in REGION_KEYS.items() if k in hq), [hq[:3]] if hq else [])
        if not comps.empty and "geography" in comps.columns and region_keys:
            geo_hits = comps["geography"].apply(
                lambda g: any(k in str(g).lower() for k in region_keys) if not _is_null(g) else False
            )
            co_geo_r = float(geo_hits.sum()) / len(comps) * 100
        else:
            co_geo_r = 50.0
        md_geo_r = 100.0

        try:
            rv_kpi = kpis[kpis["revenue_usd"].notna()].sort_values("period_end_date")
            if len(rv_kpi) >= 2:
                old_rv = rv_kpi.iloc[max(0, len(rv_kpi) - 5)]["revenue_usd"]
                new_rv = rv_kpi.iloc[-1]["revenue_usd"]
                co_growth_raw = (new_rv - old_rv) / old_rv * 100 if old_rv > 0 else None
            else:
                co_growth_raw = None
        except Exception:
            co_growth_raw = None

        grow_col      = pd.to_numeric(
            comps["revenue_growth_at_exit"] if "revenue_growth_at_exit" in comps else pd.Series(dtype=float),
            errors="coerce",
        ).dropna()
        md_growth_raw = float(grow_col.median()) if not grow_col.empty else 60.0
        co_growth_r   = _norm(co_growth_raw, 0, 200)
        md_growth_r   = _norm(md_growth_raw, 0, 200)

        latest_cust = None
        for _cc in ["customer_count", "active_clients_count"]:
            if _cc in kpis.columns:
                _cv = kpis[_cc].dropna()
                if not _cv.empty:
                    latest_cust = float(_cv.iloc[-1])
                    break
        co_cust_r = (
            min(_norm(math.log10(max(latest_cust, 1)), 0, 7), 100)
            if latest_cust and latest_cust > 0 else 0.0
        )
        md_cust_r = 70.0

        cats  = ["Gross Margin", "Revenue Scale", "EBITDA", "Geography", "Growth", "Customers"]
        co_v  = [co_gm_r, co_rev_r, co_em_r, co_geo_r, co_growth_r, co_cust_r]
        md_v  = [md_gm_r, md_rev_r, md_em_r, md_geo_r, md_growth_r, md_cust_r]

        fig_r = go.Figure()
        fig_r.add_trace(go.Scatterpolar(
            r=co_v + [co_v[0]], theta=cats + [cats[0]],
            fill="toself", fillcolor="rgba(213,250,148,0.30)",
            line=dict(color=BLACK, width=2), name=company_name,
            hovertemplate="%{theta}: %{r:.0f}/100<extra></extra>",
        ))
        fig_r.add_trace(go.Scatterpolar(
            r=md_v + [md_v[0]], theta=cats + [cats[0]],
            fill="toself", fillcolor="rgba(197,229,255,0.30)",
            line=dict(color="#1565C0", width=2, dash="dot"), name="Comp Median",
            hovertemplate="%{theta}: %{r:.0f}/100<extra></extra>",
        ))
        fig_r.update_layout(
            polar=dict(
                bgcolor=WHITE,
                radialaxis=dict(visible=False, range=[0, 100]),
                angularaxis=dict(
                    tickfont=dict(size=11, color=BLACK),
                    linecolor=BORDER, gridcolor=BORDER,
                ),
            ),
            showlegend=True,
            legend=dict(
                font=dict(size=11, color=BLACK), bgcolor=WHITE,
                bordercolor=BORDER, borderwidth=1,
                orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
            ),
            paper_bgcolor=BG,
            margin=dict(l=30, r=30, t=20, b=50),
            height=380,
        )
        st.plotly_chart(fig_r, use_container_width=True, config={"displayModeBar": False})

    # ── Section 3: Implied exit value card (sector-aware) ─────────────────────
    if company_name == "Yoco" and ltm_val is not None:
        ltm_revenue = ltm_val  # already in raw USD from caller

        def _sh_val(text):
            st.markdown(
                f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
                f"margin:20px 0 6px 0;letter-spacing:.3px'>{text}</div>",
                unsafe_allow_html=True,
            )

        _sh_val("Implied Valuation Range")
        st.markdown(
            f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
            f"Based on Bruwer ISP analysis and comparable exit multiples. "
            f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
            unsafe_allow_html=True,
        )

        HDR = (
            f"font-size:10px;font-weight:700;color:#93A3A1;"
            f"text-transform:uppercase;letter-spacing:.5px"
        )
        hcols = st.columns([2, 1, 1, 1, 2])
        for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
            with hc:
                st.markdown(f"<div style='{HDR}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown(
            f"<div style='height:2px;background:{BORDER};margin:6px 0 10px'></div>",
            unsafe_allow_html=True,
        )

        def _val_row(pathway_name, tag, tag_bg, tag_fg, multiple_lbl,
                     low, base, high, base_color, note):
            cols = st.columns([2, 1, 1, 1, 2])
            with cols[0]:
                st.markdown(
                    f"<div style='font-size:14px;font-weight:700;color:{BLACK};padding-top:4px'>"
                    f"{pathway_name}</div>"
                    f"<span style='font-size:11px;font-weight:600;background:{tag_bg};color:{tag_fg};"
                    f"border-radius:4px;padding:2px 7px'>{tag}</span>",
                    unsafe_allow_html=True,
                )
            with cols[1]:
                st.markdown(
                    f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{multiple_lbl}</div>",
                    unsafe_allow_html=True,
                )
            with cols[2]:
                st.markdown(
                    f"<div style='font-size:14px;color:{BLACK};padding-top:6px'>{fmt_usd(low)}</div>",
                    unsafe_allow_html=True,
                )
            with cols[3]:
                st.markdown(
                    f"<div style='font-size:14px;font-weight:700;color:{base_color};padding-top:6px'>"
                    f"{fmt_usd(base)}</div>",
                    unsafe_allow_html=True,
                )
            with cols[4]:
                st.markdown(
                    f"<div style='font-size:14px;color:{MUTED};padding-top:6px'>Up to {fmt_usd(high)}</div>",
                    unsafe_allow_html=True,
                )
            st.markdown(
                f"<div style='font-size:11px;color:{MUTED};font-style:italic;margin:4px 0 8px'>{note}</div>",
                unsafe_allow_html=True,
            )
            st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>", unsafe_allow_html=True)

        r = ltm_revenue
        _val_row(
            "Local Strategic Sale",
            "Most likely — 12–24 months", GREEN, BLACK,
            "2–4x Revenue",
            r * 2, r * 3, r * 5,
            "#2E7D32",
            "Consistent with iKhokha ($94M at 4–5x) and TymeBank–Retail Capital ($85–90M at ~2.5x). "
            "SA bank/telco deals capped below $400M.",
        )
        _val_row(
            "Global Strategic Sale",
            "Low feasibility", "#D4D5CE", BLACK,
            "8–13x Revenue",
            r * 8, r * 10, r * 13,
            "#1565C0",
            "Consistent with iZettle–PayPal ($2.2B at 13x) and Paystack–Stripe ($200–250M). "
            "Requires profitability and pan-African narrative.",
        )
        _val_row(
            "Remain Independent",
            "Unattractive", "#D4D5CE", BLACK,
            "2–3x Revenue",
            r * 2, r * 2.5, r * 3,
            BLACK,
            "SA independents rarely exceed $300M. Growth ceiling as banks and telcos consolidate.",
        )

        st.markdown(
            f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
            f"font-size:11px;color:{MUTED};margin-top:8px'>"
            f"Valuation ranges are indicative and based on comparable transaction multiples from "
            f"Bruwer ISP exit analysis (May 2026). Actual exit valuation will depend on buyer appetite, "
            f"competitive dynamics, profitability trajectory, and market conditions at time of exit."
            f"</div>",
            unsafe_allow_html=True,
        )

    elif ltm_val is not None and comp_rev is not None and comp_rev > 0:
        sector = str(info.get("sector", "")).lower()

        # ── derive the latest non-null value for a kpis column ───────────────
        def _latest(col):
            if col in kpis.columns:
                v = kpis[col].dropna()
                return float(v.iloc[-1]) if not v.empty else None
            return None

        # ── compute comp EV/EBITDA from comp rows ────────────────────────────
        def _comp_ev_ebitda():
            hi = (
                comps[comps["data_confidence"].str.lower().isin(["high", "medium"])]
                if "data_confidence" in comps.columns else comps
            )
            vals = []
            for _, r in hi.iterrows():
                ev  = r.get("exit_ev_usd_m")
                rev = r.get("revenue_at_exit_usd_m")
                em  = _parse_pct(r.get("ebitda_margin_pct"))
                if not any(_is_null(x) for x in (ev, rev, em)) and em > 0 and rev > 0:
                    vals.append(float(ev) / (float(rev) * em / 100))
            return float(pd.Series(vals).median()) if vals else None

        # ── sector routing ────────────────────────────────────────────────────
        implied_ev   = None
        multiple_val = None
        method_lbl   = "EV/Revenue"
        method_note  = "Comp median"
        base_val     = ltm_val
        base_lbl     = "LTM Revenue"

        if sector == "lending":
            loan_book = _latest("loan_book_gross_usd")
            if loan_book and loan_book > 0:
                pb_multiple  = 2.0
                implied_ev   = loan_book * pb_multiple
                multiple_val = pb_multiple
                method_lbl   = "P/Book"
                method_note  = "2.0x P/Book (SSA digital lender benchmark)"
                base_val     = loan_book
                base_lbl     = "Gross Loan Book"
            elif comp_ev:
                implied_ev   = ltm_val * comp_ev
                multiple_val = comp_ev
                method_note  = "Comp median EV/Revenue (loan book data unavailable)"

        elif sector == "marketplace":
            gmv = _latest("gmv_usd")
            if gmv and gmv > 0:
                ev_gmv       = 0.5
                implied_ev   = gmv * ev_gmv
                multiple_val = ev_gmv
                method_lbl   = "EV/GMV"
                method_note  = "0.5x EV/GMV (SSA marketplace benchmark)"
                base_val     = gmv
                base_lbl     = "LTM GMV"
            elif comp_ev:
                implied_ev   = ltm_val * comp_ev
                multiple_val = comp_ev
                method_note  = "Comp median EV/Revenue (GMV data unavailable)"

        elif sector == "wealth_management":
            aum = _latest("aum_usd")
            if comp_ev:
                implied_ev   = ltm_val * comp_ev
                multiple_val = comp_ev
                method_note  = (
                    f"Comp median EV/Revenue · AUM multiple is more relevant "
                    f"({fmt_usd(aum)} AUM available)" if aum
                    else "Comp median EV/Revenue · AUM multiple preferred when AUM data available"
                )

        elif sector in ("iot_infrastructure", "saas"):
            if ltm_em_pct is not None and ltm_em_pct > 0:
                ltm_ebitda   = ltm_val * ltm_em_pct / 100
                ev_ebitda    = _comp_ev_ebitda()
                if ev_ebitda:
                    implied_ev   = ltm_ebitda * ev_ebitda
                    multiple_val = ev_ebitda
                    method_lbl   = "EV/EBITDA"
                    method_note  = f"{ev_ebitda:.1f}x EV/EBITDA (comp median, profitable)"
                    base_val     = ltm_ebitda
                    base_lbl     = "LTM EBITDA"
            if implied_ev is None and comp_ev:
                implied_ev   = ltm_val * comp_ev
                multiple_val = comp_ev
                method_note  = "Comp median EV/Revenue (pre-profitability)"

        else:  # payments and default
            if comp_ev:
                implied_ev   = ltm_val * comp_ev
                multiple_val = comp_ev
                method_note  = "Comp median"

        if implied_ev is not None and multiple_val is not None:
            # scale bar always shows revenue position vs comp exit scale
            scale_pct   = min(rev_m / comp_rev * 100, 100) if rev_m else 0
            marker_left = max(5.0, min(scale_pct, 95.0))

            st.markdown(
                f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                f"padding:24px 28px;margin-bottom:20px'>"
                f"<div style='display:flex;justify-content:space-between;align-items:flex-start;"
                f"flex-wrap:wrap;gap:8px;margin-bottom:14px'>"
                f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                f"color:{MUTED};font-weight:600'>Implied Exit Value</div>"
                f"<div style='background:{BG};border-radius:4px;padding:3px 10px;"
                f"font-size:11px;font-weight:600;color:{BLACK}'>"
                f"{multiple_val:.1f}x {method_lbl}</div>"
                f"</div>"
                f"<div style='display:flex;align-items:baseline;gap:14px;flex-wrap:wrap;margin-bottom:20px'>"
                f"<span style='font-size:36px;font-weight:700;color:{BLACK}'>{fmt_usd(implied_ev)}</span>"
                f"<span style='font-size:13px;color:{MUTED}'>{base_lbl}: {fmt_usd(base_val)}</span>"
                f"</div>"
                f"<div style='position:relative;height:14px;margin-bottom:6px'>"
                f"<div style='background:linear-gradient(to right,{BG} 0%,{GREEN} 100%);"
                f"border-radius:7px;height:14px;width:100%'></div>"
                f"<div style='position:absolute;top:50%;left:{marker_left:.1f}%;"
                f"transform:translate(-50%,-50%)'>"
                f"<div style='width:20px;height:20px;border-radius:50%;background:{BLACK};"
                f"border:3px solid {WHITE};box-shadow:0 0 0 2px {BLACK}'></div>"
                f"</div></div>"
                f"<div style='display:flex;justify-content:space-between;margin-bottom:14px'>"
                f"<span style='font-size:11px;color:{MUTED}'>$0 revenue</span>"
                f"<span style='font-size:11px;color:{MUTED}'>Comp median exit revenue ({fmt_usd(comp_rev * 1e6)})</span>"
                f"</div>"
                f"<div style='font-size:12px;color:{MUTED};border-top:1px solid {BORDER};padding-top:12px'>"
                f"<b style='color:{BLACK}'>Methodology:</b> {method_note}. "
                f"{company_name} is at <b style='color:{BLACK}'>{scale_pct:.0f}%</b> of comp median exit revenue scale."
                f"</div></div>",
                unsafe_allow_html=True,
            )

    # ── Section 4: Stage timeline ──────────────────────────────────────────────
    snapshots = load_stage_snapshots(comp_ids)
    if comp_rev is not None and comp_rev > 0 and ltm_val is not None:
        ltm_m      = ltm_val / 1e6
        scale_frac = ltm_m / comp_rev  # 0.0 → 1.0+

        # Fixed 4-node stages — fractions used only for snapshot stat lookups
        STAGE_NODES = [
            ("Early Stage", 0.00, 0.25),
            ("Growth",      0.25, 0.60),
            ("Pre-Exit",    0.60, 0.90),
            ("Exit Ready",  0.90, None),
        ]
        STAGE_SUBLABELS = [
            "< $5M revenue",
            "$6M – $30M revenue",
            "$31M – $100M revenue",
            "> $100M + EBITDA positive",
        ]
        # Stage classification on absolute revenue thresholds
        if ltm_m < 5:
            current_stage_idx = 0  # Early Stage
        elif ltm_m <= 30:
            current_stage_idx = 1  # Growth
        elif ltm_m <= 100:
            current_stage_idx = 2  # Pre-Exit
        elif ltm_em_pct is not None and ltm_em_pct > 0:
            current_stage_idx = 3  # Exit Ready — > $100M + EBITDA positive
        else:
            current_stage_idx = 2  # > $100M but not yet EBITDA positive → Pre-Exit

        # Derive per-stage rev ranges and median GM from snapshot data when available
        # _parse_pct handles "~40%", "(30%)", "60%+" — pd.to_numeric alone cannot
        if not snapshots.empty:
            for _col in ["gross_margin_pct", "ebitda_margin_pct", "revenue_growth_pct"]:
                if _col in snapshots.columns:
                    snapshots[_col] = snapshots[_col].apply(_parse_pct)
            snapshots["rev_mid"] = snapshots["revenue_range_usd_m"].apply(_rev_range_mid)

        def _stage_stats(lo_frac, hi_frac):
            lo_m = comp_rev * lo_frac
            hi_m = comp_rev * hi_frac if hi_frac else None
            if snapshots.empty or "rev_mid" not in snapshots.columns:
                return None, None
            mask = snapshots["rev_mid"] >= lo_m
            if hi_m is not None:
                mask &= snapshots["rev_mid"] < hi_m
            sub = snapshots[mask]
            gm = sub["gross_margin_pct"].dropna().median() if not sub.empty else None
            rev_lo = f"${lo_m:.0f}M"
            rev_hi = f"${hi_m:.0f}M" if hi_m else f"${lo_m:.0f}M+"
            return f"{rev_lo}–{rev_hi}", (float(gm) if not _is_null(gm) else None)

        st.markdown(
            f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
            f"color:{MUTED};font-weight:600;margin-bottom:14px'>Stage Progression</div>",
            unsafe_allow_html=True,
        )

        nodes_html = ""
        for i, (stage_name, lo, hi) in enumerate(STAGE_NODES):
            is_cur     = (i == current_stage_idx)
            dot_bg     = GREEN  if is_cur else WHITE
            dot_border = BLACK  if is_cur else BORDER
            dot_size   = "20px" if is_cur else "14px"
            lbl_fw     = "700"  if is_cur else "400"
            lbl_col    = BLACK  if is_cur else MUTED
            ll_bg      = BORDER if i > 0 else "transparent"
            rl_bg      = BORDER if i < 3 else "transparent"
            dot_shadow = f"0 0 0 3px {GREEN}" if is_cur else "none"

            rev_range_str, gm_val = _stage_stats(lo, hi)
            gm_str = fmt_pct(gm_val) if not _is_null(gm_val) else "—"

            badge_html = (
                f"<div style='background:{GREEN};color:{BLACK};border-radius:4px;"
                f"padding:1px 8px;font-size:10px;font-weight:700;margin-bottom:5px;"
                f"display:inline-block;white-space:nowrap'>{company_name}</div><br>"
                if is_cur else "<br>"
            )

            nodes_html += (
                f"<div style='flex:1;text-align:center;padding:0 8px'>"
                f"{badge_html}"
                f"<div style='font-size:13px;font-weight:{lbl_fw};color:{lbl_col};"
                f"margin-bottom:10px'>{stage_name}</div>"
                f"<div style='display:flex;align-items:center;justify-content:center;"
                f"margin-bottom:10px'>"
                f"<div style='height:2px;flex:1;background:{ll_bg}'></div>"
                f"<div style='width:{dot_size};height:{dot_size};border-radius:50%;"
                f"background:{dot_bg};border:2px solid {dot_border};flex-shrink:0;"
                f"box-shadow:{dot_shadow}'></div>"
                f"<div style='height:2px;flex:1;background:{rl_bg}'></div>"
                f"</div>"
                f"<div style='font-size:10px;color:{MUTED};text-align:center'>{STAGE_SUBLABELS[i]}</div>"
                f"</div>"
            )

        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:24px 20px;display:flex;align-items:flex-start;margin-bottom:20px'>"
            f"{nodes_html}</div>",
            unsafe_allow_html=True,
        )

    # ── Section 5: AI commentary card ─────────────────────────────────────────
    commentary = st.session_state.get(f"upload_commentary_{company_id}")
    if commentary:
        st.markdown(
            f"<div style='border-left:4px solid {GREEN};background:{WHITE};"
            f"border-radius:0 10px 10px 0;padding:16px 20px;margin-bottom:20px;"
            f"box-shadow:0 1px 3px rgba(0,0,0,.05)'>"
            f"<div style='font-size:10px;text-transform:uppercase;letter-spacing:.6px;"
            f"color:{MUTED};font-weight:600;margin-bottom:8px'>AI Commentary</div>"
            f"<div style='font-size:13px;color:{BLACK};line-height:1.65'>{commentary}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

    # ── Section 6: Peer comp table ────────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin-bottom:12px'>Peer Comp Set</div>",
        unsafe_allow_html=True,
    )

    CONF_DOT = {"high": "#2E7D32", "medium": "#F57C00", "low": "#C62828"}
    REL_COLORS = {
        5: (GREEN,     BLACK),
        4: ("#D9F0D9", "#2E7D32"),
        3: ("#FFF9C4", "#795548"),
        2: ("#EFEBE9", MUTED),
        1: ("#F5F5F5", MUTED),
    }

    hdr_html = "".join(
        f"<th style='padding:8px 12px;text-align:left;font-size:10px;text-transform:uppercase;"
        f"letter-spacing:.5px;color:{MUTED};border-bottom:2px solid {BORDER};white-space:nowrap;"
        f"width:{w}'>{h}</th>"
        for h, w in [
            ("Company", "18%"), ("Sub-sector", "14%"), ("Geography", "8%"),
            ("Exit Type", "8%"), ("Year", "5%"),
            ("Rev at Exit", "8%"), ("Gross Margin", "8%"), ("EBITDA Margin", "8%"),
            ("EV/Rev", "7%"), ("Relevance", "8%"), ("Conf.", "6%"),
        ]
    )

    EXIT_TYPE_COLORS = {
        "acquisition":      ("#C5E5FF", "#1565C0"),
        "ipo":              ("#D5FA94", "#2C2C2A"),
        "private funding":  ("#D4D5CE", "#2C2C2A"),
    }

    rows_html      = ""
    separator_done = False
    for idx, row in comps.iterrows():
        is_clean   = int(row.get("is_clean_exit", 1))
        rel        = int(row["relevance_score"]) if not _is_null(row.get("relevance_score")) else 0
        bg_r, fg_r = REL_COLORS.get(rel, ("#F5F5F5", MUTED))
        rev        = row.get("revenue_at_exit_usd_m")
        gm         = row.get("gross_margin_pct")
        em         = row.get("ebitda_margin_pct")
        ev         = row.get("ev_revenue_multiple")
        conf_raw   = str(row.get("data_confidence", "")).lower()
        conf_dot_c = CONF_DOT.get(conf_raw, MUTED)
        row_bg     = WHITE if idx % 2 == 0 else "#F9FAF7"

        # Inject separator row when transitioning to pre-exit comps
        if not is_clean and not separator_done:
            separator_done = True
            n_cols = 11
            rows_html += (
                f"<tr><td colspan='{n_cols}' style='padding:4px 12px;background:#F9FAF7;"
                f"border-top:2px dashed {BORDER};border-bottom:2px dashed {BORDER}'>"
                f"<span style='font-size:10px;font-weight:700;color:{MUTED};"
                f"text-transform:uppercase;letter-spacing:.5px'>"
                f"Pre-exit / Funding Marks — excluded from median calculations</span>"
                f"</td></tr>"
            )

        url       = row.get("announcement_url") or ""
        co_name   = row["company_name"]
        name_html = (
            f"<a href='{url}' target='_blank' rel='noopener noreferrer' "
            f"style='color:{BLACK};text-decoration:underline;text-underline-offset:2px'>"
            f"{co_name}</a>"
            if url else co_name
        )
        # Add Zettle note
        if co_name == "Zettle":
            name_html += f"<div style='font-size:10px;color:{MUTED};font-style:italic'>Same deal as iZettle</div>"

        exit_type_raw = str(row.get("exit_type") or "").strip()
        et_key        = exit_type_raw.lower()
        if not is_clean:
            et_html = (
                f"<span style='font-size:11px;color:{MUTED};font-style:italic'>"
                f"Pre-exit / funding mark</span>"
            )
        else:
            et_bg, et_fg = EXIT_TYPE_COLORS.get(et_key, ("#D4D5CE", "#2C2C2A"))
            et_html = (
                f"<span style='background:{et_bg};color:{et_fg};border-radius:4px;"
                f"padding:2px 7px;font-size:11px;font-weight:600'>{exit_type_raw}</span>"
                if exit_type_raw else "—"
            )

        exit_year_raw = row.get("exit_year")
        year_html     = str(int(exit_year_raw)) if not _is_null(exit_year_raw) else "—"

        rows_html += (
            f"<tr style='background:{row_bg};opacity:{'0.7' if not is_clean else '1'}'>"
            f"<td style='padding:8px 12px;font-weight:600;color:{BLACK};width:18%'>{name_html}</td>"
            f"<td style='padding:8px 12px;font-size:12px;color:{MUTED};width:14%'>"
            f"{(row.get('sub_sector') or '—').replace('_',' ').title()}</td>"
            f"<td style='padding:8px 12px;font-size:12px;color:{MUTED};width:8%'>{row.get('geography','—')}</td>"
            f"<td style='padding:8px 12px;width:8%'>{et_html}</td>"
            f"<td style='padding:8px 12px;font-size:12px;color:{MUTED};width:5%'>{year_html}</td>"
            f"<td style='padding:8px 12px;font-weight:500;width:8%'>"
            f"{'$'+str(round(rev))+'M' if not _is_null(rev) else '—'}</td>"
            f"<td style='padding:8px 12px;width:8%'>{fmt_pct(gm)}</td>"
            f"<td style='padding:8px 12px;width:8%'>{fmt_pct(em)}</td>"
            f"<td style='padding:8px 12px;width:7%'>{f'{ev:.1f}x' if not _is_null(ev) else '—'}</td>"
            f"<td style='padding:8px 12px;width:8%'>"
            f"<span style='background:{bg_r};color:{fg_r};border-radius:4px;"
            f"padding:2px 8px;font-size:11px;font-weight:600'>{rel}/5</span></td>"
            f"<td style='padding:8px 12px;width:6%;text-align:center'>"
            f"<span title='{conf_raw.capitalize()}' style='display:inline-block;width:10px;height:10px;"
            f"border-radius:50%;background:{conf_dot_c}'></span></td>"
            f"</tr>"
        )

    st.markdown(
        f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
        f"overflow:auto;margin-bottom:8px'>"
        f"<table style='width:100%;border-collapse:collapse'>"
        f"<thead><tr style='background:{BG}'>{hdr_html}</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        f"</table></div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:11px;color:{MUTED};font-style:italic;margin-bottom:20px'>"
        f"Median calculations exclude SumUp (pre-exit funding mark, 45.2x) and CloudWalk "
        f"(pre-exit, no EV/Rev) as these are not completed exits. "
        f"Zettle is the post-rebrand entity from the same iZettle–PayPal transaction."
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── Yoco: Affinity deal intelligence scan ────────────────────────────────────
    if company_name == "Yoco":
        st.markdown(
            f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
            f"color:{MUTED};font-weight:600;margin:24px 0 12px'>Deal Intelligence — Affinity Scan</div>",
            unsafe_allow_html=True,
        )

        PILL_COLORS = [
            "#C5E5FF", "#D5FA94", "#FFE0B2", "#F8BBD9", "#D4D5CE",
            "#B2EBF2", "#E1BEE7", "#DCEDC8", "#FFF9C4", "#FFCCBC",
        ]

        def _kw_pills(keywords: list[str]) -> str:
            html = ""
            for i, kw in enumerate(keywords):
                bg = PILL_COLORS[i % len(PILL_COLORS)]
                html += (
                    f"<span style='background:{bg};color:{BLACK};font-size:10px;"
                    f"font-weight:600;border-radius:4px;padding:1px 6px;"
                    f"margin-right:3px;white-space:nowrap'>{kw}</span>"
                )
            return html

        if st.button("Scan Affinity for M&A Intel", key="yoco_affinity_ma_scan"):
            try:
                _api_key = st.secrets.get("AFFINITY_API_KEY", "")
                if not _api_key:
                    st.warning("AFFINITY_API_KEY not set in secrets.toml")
                else:
                    with st.spinner("Scanning all Affinity notes for M&A signals…"):
                        st.session_state["affinity_deal_intel"] = fetch_affinity_deal_intel(_api_key)
            except Exception as exc:
                st.error(f"Affinity scan failed: {exc}")

        intel = st.session_state.get("affinity_deal_intel")
        if intel is not None:
            if intel:
                st.markdown(
                    f"<div style='font-size:12px;color:{MUTED};margin-bottom:12px'>"
                    f"Found <b style='color:{BLACK}'>{len(intel)}</b> notes with M&A signals "
                    f"across Affinity in the last 365 days.</div>",
                    unsafe_allow_html=True,
                )

                # Column headers
                hdr_style = (
                    f"font-size:10px;font-weight:700;color:#93A3A1;"
                    f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
                )
                hcols = st.columns([1, 1, 2, 3, 1])
                for hc, lbl in zip(hcols, ["Date", "Author", "Keywords", "Snippet", "Action"]):
                    with hc:
                        st.markdown(f"<div style='{hdr_style}'>{lbl}</div>", unsafe_allow_html=True)
                st.markdown(
                    f"<div style='height:2px;background:{BORDER};margin-bottom:8px'></div>",
                    unsafe_allow_html=True,
                )

                for i, note in enumerate(intel):
                    row_bg = "#EFF0EA" if i % 2 == 0 else WHITE
                    with st.container():
                        st.markdown(
                            f"<div style='background:{row_bg};border-radius:6px;padding:4px 2px'>",
                            unsafe_allow_html=True,
                        )
                        rcols = st.columns([1, 1, 2, 3, 1])
                        with rcols[0]:
                            st.markdown(
                                f"<div style='font-size:12px;color:{BLACK};padding-top:6px'>"
                                f"{note['date']}</div>",
                                unsafe_allow_html=True,
                            )
                        with rcols[1]:
                            st.markdown(
                                f"<div style='font-size:12px;color:{MUTED};padding-top:6px'>"
                                f"{note['creator_name']}</div>",
                                unsafe_allow_html=True,
                            )
                        with rcols[2]:
                            st.markdown(
                                f"<div style='padding-top:4px'>{_kw_pills(note['matched_keywords'])}</div>",
                                unsafe_allow_html=True,
                            )
                        with rcols[3]:
                            st.markdown(
                                f"<div style='font-size:12px;color:{BLACK};padding-top:6px;"
                                f"line-height:1.4'>{note['snippet'][:150]}"
                                f"{'…' if len(note['snippet']) > 150 else ''}</div>",
                                unsafe_allow_html=True,
                            )
                        with rcols[4]:
                            if st.button("+ Add comp", key=f"add_comp_intel_{i}"):
                                st.session_state[f"add_comp_open_{i}"] = True
                        st.markdown("</div>", unsafe_allow_html=True)

                    if st.session_state.get(f"add_comp_open_{i}"):
                        with st.form(key=f"add_comp_form_{i}"):
                            st.markdown(
                                f"<div style='font-size:12px;font-weight:600;color:{BLACK};"
                                f"margin-bottom:8px'>Add company as comp</div>",
                                unsafe_allow_html=True,
                            )
                            fc1, fc2, fc3, fc4 = st.columns(4)
                            with fc1:
                                new_name = st.text_input("Company name", key=f"ci_name_{i}")
                            with fc2:
                                new_exit_type = st.selectbox(
                                    "Exit type", ["Acquisition", "IPO", "Private Funding"],
                                    key=f"ci_exit_{i}",
                                )
                            with fc3:
                                new_rev = st.number_input(
                                    "Revenue at exit ($M)", min_value=0.0, step=1.0, key=f"ci_rev_{i}"
                                )
                            with fc4:
                                new_mult = st.number_input(
                                    "EV/Rev multiple", min_value=0.0, step=0.1, key=f"ci_mult_{i}"
                                )
                            if st.form_submit_button("Insert into exit_comps"):
                                try:
                                    _conn_comps = sqlite3.connect(
                                        "data/quona_exit_comps.db", check_same_thread=False
                                    )
                                    now_iso = datetime.utcnow().isoformat()
                                    _conn_comps.execute(
                                        """INSERT INTO exit_comps
                                           (company_name, exit_type, revenue_at_exit_usd_m,
                                            ev_revenue_multiple, data_source, created_at, updated_at)
                                           VALUES (?,?,?,?,?,?,?)""",
                                        (new_name, new_exit_type,
                                         new_rev if new_rev > 0 else None,
                                         new_mult if new_mult > 0 else None,
                                         "Affinity Intel", now_iso, now_iso),
                                    )
                                    _conn_comps.commit()
                                    _conn_comps.close()
                                    st.success(f"Added {new_name} to exit_comps.")
                                    st.session_state.pop(f"add_comp_open_{i}", None)
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"Insert failed: {exc}")
            else:
                st.markdown(
                    f"<div style='font-size:13px;color:{MUTED};font-style:italic'>"
                    f"No M&A signals found in Affinity notes from the last 365 days.</div>",
                    unsafe_allow_html=True,
                )

    # ── Mapping rationale expander ─────────────────────────────────────────────
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


# ── Exit tracking DB helpers ──────────────────────────────────────────────────

def _exit_pathways_load(company_id: int) -> list[dict]:
    rows = _conn().execute(
        "SELECT id, pathway_name, likelihood, estimated_timeline, notes "
        "FROM exit_pathways WHERE company_id=? ORDER BY id",
        (company_id,),
    ).fetchall()
    return [{"id": r[0], "pathway_name": r[1], "likelihood": r[2],
             "estimated_timeline": r[3], "notes": r[4]} for r in rows]


def _exit_pathway_save(company_id: int, pid, name: str, likelihood: str,
                       timeline: str, notes: str) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    if pid:
        conn.execute(
            "UPDATE exit_pathways SET pathway_name=?,likelihood=?,"
            "estimated_timeline=?,notes=?,updated_at=? WHERE id=?",
            (name, likelihood, timeline, notes, now, pid),
        )
    else:
        conn.execute(
            "INSERT INTO exit_pathways "
            "(company_id,pathway_name,likelihood,estimated_timeline,notes,created_at,updated_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (company_id, name, likelihood, timeline, notes, now, now),
        )
    conn.commit()
    conn.close()


def _exit_pathway_delete(pid: int) -> None:
    conn = _conn()
    conn.execute("DELETE FROM exit_pathways WHERE id=?", (pid,))
    conn.commit()
    conn.close()


def _buyer_tracking_load(company_id: int) -> pd.DataFrame:
    return pd.read_sql_query(
        "SELECT id, acquirer_name, acquirer_type, relationship_owner, "
        "last_contact_date, status FROM buyer_tracking "
        "WHERE company_id=? ORDER BY sort_order, id",
        _conn(), params=(company_id,),
    )


def _buyer_tracking_replace(company_id: int, df: pd.DataFrame) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    conn.execute("DELETE FROM buyer_tracking WHERE company_id=?", (company_id,))
    for i, row in df.iterrows():
        name = str(row.get("acquirer_name", "")).strip()
        if not name:
            continue
        conn.execute(
            "INSERT INTO buyer_tracking "
            "(company_id,acquirer_name,acquirer_type,relationship_owner,"
            "last_contact_date,status,sort_order,created_at,updated_at) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (company_id, name,
             str(row.get("acquirer_type", "Strategic")),
             str(row.get("relationship_owner", "") or ""),
             str(row.get("last_contact_date", "") or ""),
             str(row.get("status", "Not Started")),
             i, now, now),
        )
    conn.commit()
    conn.close()


def _quarterly_actions_load(company_id: int, quarter: str) -> dict:
    row = _conn().execute(
        "SELECT planned_actions, completed_actions, carry_forward "
        "FROM quarterly_actions WHERE company_id=? AND quarter=?",
        (company_id, quarter),
    ).fetchone()
    return {
        "planned_actions":   (row[0] or "") if row else "",
        "completed_actions": (row[1] or "") if row else "",
        "carry_forward":     (row[2] or "") if row else "",
    }


def _quarterly_actions_save(company_id: int, quarter: str,
                             planned: str, completed: str, carry: str) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    exists = conn.execute(
        "SELECT id FROM quarterly_actions WHERE company_id=? AND quarter=?",
        (company_id, quarter),
    ).fetchone()
    if exists:
        conn.execute(
            "UPDATE quarterly_actions SET planned_actions=?,completed_actions=?,"
            "carry_forward=?,updated_at=? WHERE company_id=? AND quarter=?",
            (planned, completed, carry, now, company_id, quarter),
        )
    else:
        conn.execute(
            "INSERT INTO quarterly_actions "
            "(company_id,quarter,planned_actions,completed_actions,carry_forward,created_at,updated_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (company_id, quarter, planned, completed, carry, now, now),
        )
    conn.commit()
    conn.close()


# ── Exit tab suggestion helpers ───────────────────────────────────────────────

_PATHWAY_DEFAULTS = [
    ("Strategic Acquisition",  "Exploratory", "3–5 years"),
    ("PE / Growth Equity",     "Exploratory", "2–4 years"),
    ("IPO / Public Listing",   "Exploratory", "5–7 years"),
]

_SECTOR_BUYERS: dict[str, list[tuple[str, str]]] = {
    "payments":         [("Pan-African Banking Group",     "Strategic"),
                         ("Global Payments Network",       "Strategic"),
                         ("African Fintech Consolidator",  "Strategic"),
                         ("Growth Equity Fund",            "Financial")],
    "lending":          [("Tier 1 African Bank",           "Strategic"),
                         ("Development Finance (DFI)",     "Financial"),
                         ("Pan-African Fintech Group",     "Strategic"),
                         ("PE / Growth Fund",              "Financial")],
    "wealth_management":[("Regional Asset Manager",        "Strategic"),
                         ("Pan-African Bank (Wealth Arm)", "Strategic"),
                         ("Global EM Investment Manager",  "Financial")],
    "marketplace":      [("African B2B Platform",          "Strategic"),
                         ("Global Marketplace Operator",   "Adjacent"),
                         ("Regional PE Fund",              "Financial")],
    "iot_infrastructure":[("Global IoT / Connectivity Co","Strategic"),
                          ("Pan-African Telco Group",      "Strategic"),
                          ("Infrastructure PE Fund",       "Financial")],
    "saas":             [("Global Vertical SaaS Co",       "Strategic"),
                         ("African Tech Conglomerate",     "Adjacent"),
                         ("Growth Equity (SaaS)",          "Financial")],
    "insurtech":        [("Pan-African Insurance Group",   "Strategic"),
                         ("Global Insurtech Player",       "Adjacent"),
                         ("PE / Growth Fund",              "Financial")],
}


def _suggest_exit_pathways(company_name: str, sector: str) -> list[dict]:
    TYPE_KEYS = {
        "strategic": ("Strategic Acquisition", "3–5 years"),
        "acqui":     ("Strategic Acquisition", "3–5 years"),
        "ipo":        ("IPO / Public Listing",  "5–7 years"),
        "public":     ("IPO / Public Listing",  "5–7 years"),
        "pe":         ("PE / Growth Equity",    "2–4 years"),
        "growth":     ("PE / Growth Equity",    "2–4 years"),
        "financial":  ("PE / Growth Equity",    "2–4 years"),
    }
    seen: dict[str, int] = {}
    try:
        mapping = load_comp_mapping(company_name)
        if not mapping.empty:
            comps = load_comps_detail(tuple(mapping["comp_id"].tolist()))
            if not comps.empty and "exit_type" in comps.columns:
                for et in comps["exit_type"].dropna():
                    for key, (name, _) in TYPE_KEYS.items():
                        if key in str(et).lower():
                            seen[name] = seen.get(name, 0) + 1
                            break
    except Exception:
        pass

    suggestions = []
    timelines   = {n: t for _, (n, t) in TYPE_KEYS.items()}
    for name, count in sorted(seen.items(), key=lambda x: -x[1]):
        suggestions.append({
            "pathway_name":       name,
            "likelihood":         "Exploratory",
            "estimated_timeline": timelines.get(name, "3–5 years"),
            "notes":              f"{count} comp(s) exited via this route",
        })
    for name, timeline, _ in _PATHWAY_DEFAULTS:
        if name not in {s["pathway_name"] for s in suggestions}:
            suggestions.append({
                "pathway_name": name, "likelihood": "Exploratory",
                "estimated_timeline": timeline, "notes": "",
            })
        if len(suggestions) >= 3:
            break
    return suggestions[:3]


def _suggest_buyers(sector: str) -> list[dict]:
    rows = _SECTOR_BUYERS.get(sector, [
        ("Strategic Acquirer (TBD)", "Strategic"),
        ("Financial Sponsor",        "Financial"),
        ("Adjacent Market Player",   "Adjacent"),
    ])
    return [{"acquirer_name": n, "acquirer_type": t,
             "relationship_owner": "", "last_contact_date": "",
             "status": "Not Started"} for n, t in rows]


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


# ── Affinity CRM helpers ──────────────────────────────────────────────────────

def fetch_affinity_interactions(company_name: str) -> list[dict]:
    """Search Affinity for company_name, return notes from last 180 days."""
    import requests

    api_key = st.secrets.get("AFFINITY_API_KEY", "")
    if not api_key:
        raise ValueError("AFFINITY_API_KEY not set in .streamlit/secrets.toml")

    BASE = "https://api.affinity.co"
    AUTH = ("", api_key)

    # Find org
    r = requests.get(f"{BASE}/organizations", params={"term": company_name},
                     auth=AUTH, timeout=15)
    r.raise_for_status()
    orgs = r.json().get("organizations", [])
    if not orgs:
        return []
    org_id = orgs[0]["id"]

    # Fetch notes
    r = requests.get(f"{BASE}/notes", params={"organization_id": org_id},
                     auth=AUTH, timeout=15)
    r.raise_for_status()
    notes = r.json().get("notes", [])

    # Cache person names to minimise API calls
    _person_cache: dict[int, str] = {}

    def _person_name(pid: int) -> str:
        if pid in _person_cache:
            return _person_cache[pid]
        try:
            rp = requests.get(f"{BASE}/persons/{pid}", auth=AUTH, timeout=10)
            p  = rp.json()
            name = f"{p.get('first_name', '')} {p.get('last_name', '')}".strip()
        except Exception:
            name = str(pid)
        _person_cache[pid] = name
        return name

    results = []
    for n in notes:
        raw_date = n.get("created_at", "")
        if not raw_date:
            continue
        note_dt = datetime.fromisoformat(raw_date)
        if note_dt.tzinfo is None:
            note_dt = note_dt.replace(tzinfo=timezone.utc)

        creator_id = n.get("creator_id")
        person_name = _person_name(creator_id) if creator_id else "Unknown"

        itype = "Meeting" if n.get("is_meeting") else "Note"
        content = (n.get("content") or "").strip()
        summary = content[:600] + ("…" if len(content) > 600 else "")

        results.append({
            "date":        note_dt.strftime("%Y-%m-%d"),
            "type":        itype,
            "person_name": person_name,
            "summary":     summary,
            "source":      "affinity",
        })

    results.sort(key=lambda x: x["date"], reverse=True)
    return results


def fetch_slack_messages(company_name: str) -> list[dict]:
    """Find the portco- Slack channel and return messages + thread replies from last 365 days."""
    import requests

    token = st.secrets.get("SLACK_BOT_TOKEN", "")
    if not token:
        raise ValueError("SLACK_BOT_TOKEN not set in .streamlit/secrets.toml")

    BASE    = "https://slack.com/api"
    HEADERS = {"Authorization": f"Bearer {token}"}

    _CHANNEL_MAP = {"VertoFX": "portco-verto"}
    if company_name in _CHANNEL_MAP:
        channel_name = _CHANNEL_MAP[company_name]
    else:
        channel_name = "portco-" + company_name.lower().replace(" ", "-")

    # Find channel ID (paginated)
    channel_id = None
    cursor = ""
    while not channel_id:
        params: dict = {"exclude_archived": "true", "limit": 200}
        if cursor:
            params["cursor"] = cursor
        r = requests.get(f"{BASE}/conversations.list", headers=HEADERS,
                         params=params, timeout=15)
        data = r.json()
        if not data.get("ok"):
            raise ValueError(f"Slack conversations.list error: {data.get('error')}")
        for ch in data.get("channels", []):
            if ch["name"] == channel_name:
                channel_id = ch["id"]
                break
        cursor = data.get("response_metadata", {}).get("next_cursor", "")
        if not cursor:
            break

    if not channel_id:
        return []

    cutoff_ts = str((datetime.now(timezone.utc) - timedelta(days=365)).timestamp())

    _user_cache: dict[str, str] = {}

    def _user_name(uid: str) -> str:
        if uid in _user_cache:
            return _user_cache[uid]
        try:
            rp = requests.get(f"{BASE}/users.info", headers=HEADERS,
                              params={"user": uid}, timeout=10)
            u = rp.json().get("user", {})
            name = u.get("real_name") or u.get("name") or uid
        except Exception:
            name = uid
        _user_cache[uid] = name
        return name

    results = []
    cursor = ""
    while True:
        params = {"channel": channel_id, "oldest": cutoff_ts, "limit": 200}
        if cursor:
            params["cursor"] = cursor
        r = requests.get(f"{BASE}/conversations.history", headers=HEADERS,
                         params=params, timeout=15)
        data = r.json()
        if not data.get("ok"):
            raise ValueError(f"Slack conversations.history error: {data.get('error')}")

        for msg in data.get("messages", []):
            if msg.get("type") != "message" or msg.get("subtype"):
                continue
            ts       = float(msg.get("ts", 0))
            date_str = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
            text     = (msg.get("text") or "").strip()
            results.append({
                "date":             date_str,
                "type":             "Message",
                "person_name":      _user_name(msg.get("user", "")),
                "summary":          text[:600] + ("…" if len(text) > 600 else ""),
                "source":           "slack",
                "is_thread_reply":  False,
            })

            # Fetch thread replies for threaded messages
            if msg.get("reply_count") and msg.get("thread_ts") == msg.get("ts"):
                rr = requests.get(f"{BASE}/conversations.replies", headers=HEADERS,
                                  params={"channel": channel_id, "ts": msg["ts"]},
                                  timeout=15)
                rdata = rr.json()
                if rdata.get("ok"):
                    for reply in rdata.get("messages", [])[1:]:
                        rts   = float(reply.get("ts", 0))
                        rtext = (reply.get("text") or "").strip()
                        results.append({
                            "date":            datetime.fromtimestamp(rts, tz=timezone.utc).strftime("%Y-%m-%d"),
                            "type":            "Thread Reply",
                            "person_name":     _user_name(reply.get("user", "")),
                            "summary":         rtext[:600] + ("…" if len(rtext) > 600 else ""),
                            "source":          "slack",
                            "is_thread_reply": True,
                        })

        cursor = data.get("response_metadata", {}).get("next_cursor", "")
        if not cursor:
            break

    results.sort(key=lambda x: x["date"], reverse=True)
    return results


def classify_exit_relevant(interactions: list[dict]) -> list[dict]:
    """Use Claude to filter interactions for exit signals and extract acquirer hints."""
    if not interactions:
        return []

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable not set")

    import json as _json

    interactions_text = _json.dumps(
        [{"index": i, "source": x.get("source", ""), "date": x.get("date", ""),
          "type": x.get("type", ""), "person": x.get("person_name", ""),
          "summary": x.get("summary", "")}
         for i, x in enumerate(interactions)],
        indent=2,
    )

    prompt = (
        "You are an M&A analyst at a venture capital firm. "
        "Review the following CRM interactions and identify which ones contain "
        "exit-relevant signals: acquisition, M&A, strategic partnership, exit, "
        "buyer, valuation, term sheet, due diligence, secondary, strategic interest, "
        "or any named potential acquirer or investor.\n\n"
        f"Interactions:\n{interactions_text}\n\n"
        "Return a JSON array of objects for ONLY the exit-relevant interactions. "
        "Each object must have:\n"
        "  - index (integer, matching the input index)\n"
        "  - acquirer_hint (string: name of any buyer/acquirer/investor mentioned, "
        "or empty string if none)\n\n"
        "Return ONLY the JSON array, no other text."
    )

    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()

    import json as _json2
    classified = _json2.loads(raw)

    relevant = []
    for item in classified:
        idx = item.get("index")
        if idx is None or idx >= len(interactions):
            continue
        entry = dict(interactions[idx])
        entry["acquirer_hint"] = item.get("acquirer_hint", "")
        relevant.append(entry)
    return relevant


# ── Yoco Affinity helper ──────────────────────────────────────────────────────

def fetch_last_affinity_note_for_buyer(buyer_name: str, affinity_api_key: str) -> dict | None:
    try:
        import requests
        AUTH = ("", affinity_api_key)
        BASE = "https://api.affinity.co"

        r = requests.get(f"{BASE}/organizations", params={"term": buyer_name}, auth=AUTH, timeout=15)
        r.raise_for_status()
        orgs = r.json().get("organizations", [])
        if not orgs:
            return None
        org_id = orgs[0]["id"]

        r = requests.get(f"{BASE}/notes", params={"organization_id": org_id}, auth=AUTH, timeout=15)
        r.raise_for_status()
        notes = r.json().get("notes", [])
        if not notes:
            return None

        def _note_dt(n):
            raw = n.get("created_at", "")
            if not raw:
                return datetime.min.replace(tzinfo=timezone.utc)
            dt = datetime.fromisoformat(raw)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)

        notes.sort(key=_note_dt, reverse=True)
        latest = notes[0]

        raw_date = latest.get("created_at", "")
        note_dt = _note_dt(latest)
        date_str = note_dt.strftime("%Y-%m-%d") if note_dt != datetime.min.replace(tzinfo=timezone.utc) else ""

        cutoff = datetime.now(tz=timezone.utc) - timedelta(days=90)
        if note_dt < cutoff:
            return {"date": date_str, "creator_name": None, "snippet": None, "stale": True}

        creator_name = "Unknown"
        creator_id = latest.get("creator_id")
        if creator_id:
            try:
                rp = requests.get(f"{BASE}/persons/{creator_id}", auth=AUTH, timeout=10)
                p = rp.json()
                creator_name = (f"{p.get('first_name', '')} {p.get('last_name', '')}".strip() or "Unknown")
            except Exception:
                pass

        content = (latest.get("content") or "").strip()
        keywords = {"yoco", "exit", "acquisition", "strategic", "partnership", buyer_name.lower()}
        relevant = [
            s.strip() for s in content.replace("\n", " ").split(".")
            if s.strip() and any(kw in s.lower() for kw in keywords)
        ]
        if relevant:
            summary = ". ".join(relevant[:2]) + "."
            summary = summary[:200] + ("…" if len(summary) > 200 else "")
        else:
            summary = "Note found — no exit-relevant content"

        return {
            "date":         date_str,
            "creator_name": creator_name,
            "snippet":      summary,
            "stale":        False,
        }
    except Exception:
        return None


# ── Yoco custom exit tab ──────────────────────────────────────────────────────

def _render_yoco_exit_tab() -> None:
    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    AMBER = "#FFC107"
    GREEN_DOT = "#D5FA94"
    RED_DOT = "#E57373"
    EMPTY = "#D4D5CE"

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>" if len(valuation) > 1 else ""
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        ("Remain Independent",    ["$150–300M"],             "Protect brand and optionality but risk gradual erosion as banks and telcos consolidate.",
         [AMBER, AMBER, EMPTY],             "Unattractive strategically", False),
        ("SME Bank Build",        ["$250–600M"],             "Partner with Sava to launch SME accounts and credit, creating a stickier ecosystem narrative.",
         [AMBER, AMBER, EMPTY],             "Execution heavy",            False),
        ("Strategic Sale Local",  ["$200–400M", "2–4x rev"], "Sell to Vodacom, MTN, Capitec, FNB or insurers — most realistic path given consolidation wave.",
         [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — 12–24 months", True),
        ("Strategic Sale Global", ["$400–600M"],             "Acquire by Stripe, Adyen or Nubank as Africa market entry — limited near-term appetite.",
         [RED_DOT, EMPTY, EMPTY],           "Low feasibility",            False),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":   ("#D5FA94", "#2C2C2A"),
        "High":        ("#C5E5FF", "#1565C0"),
        "Medium":      ("#D4D5CE", "#2C2C2A"),
        "Low-Medium":  ("#FFCDD2", "#B71C1C"),
        "Low":         ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
                f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>")

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 2])
            with cols[0]:
                st.markdown(
                    f"<div style='padding-top:6px'><span style='font-weight:700;color:#2C2C2A'>{name}</span>"
                    f"{_fit_badge(fit)}</div>",
                    unsafe_allow_html=True,
                )
            with cols[1]:
                st.markdown(
                    f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{activity}</div>",
                    unsafe_allow_html=True,
                )
            with cols[2]:
                st.markdown(
                    f"<div style='font-size:13px;color:#2C2C2A;padding-top:6px'>{rationale}</div>",
                    unsafe_allow_html=True,
                )
            with cols[3]:
                st.checkbox("", key=key)
            with cols[4]:
                if affinity_cache is None:
                    st.markdown(
                        f"<div style='font-size:11px;color:{MUTED};padding-top:8px'>Sync Affinity above</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    note = affinity_cache.get(name)
                    if note is None:
                        st.markdown(
                            f"<div style='font-size:11px;color:{MUTED};font-style:italic;padding-top:8px'>Not in Affinity</div>",
                            unsafe_allow_html=True,
                        )
                    elif note.get("stale"):
                        st.markdown(
                            f"<div style='font-size:11px;color:#E65100;font-weight:600;padding-top:4px'>No update in 90 days</div>"
                            f"<div style='font-size:11px;color:{MUTED}'>Last contact: {note['date']}</div>",
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED}'>{note['snippet']}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("Capitec",    "Very High", "Rapid SME banking rollout under new CEO",
         "Would accelerate SME acquiring with 110k+ merchants. Issuer-backed economics makes Yoco highly valuable. Risk: may replicate organically."),
        ("Vodacom",    "High",      "Active SME lending partnership with Lula",
         "Direct SME merchant access and POS infra. Existing Lula ties could complicate structure but strategically strong fit."),
        ("FNB",        "High",      "Expanding SME payments and digital services",
         "Reach into long-tail merchants. Majority of Yoco merchants bank with FNB already."),
        ("TymeBank",   "High",      "Bought Retail Capital ~$85–90M (2022)",
         "Fills merchant acquiring gap perfectly. Concern: questions on Yoco team and profitability."),
        ("Lesaka",     "Medium",    "Acquired Adumo ~$86–96M (2024)",
         "Would cement largest independent acquirer in SA. Heavy merchant overlap makes integration complex."),
        ("MTN",        "Medium",    "Expanding MoMo into payments and lending",
         "Strengthen SME payments credibility. Patchy SA execution reduces near-term likelihood."),
        ("Nedbank",    "Medium",    "Acquired iKhokha ~$94M (2025)",
         "iKhokha already addressed SME acquiring gap, making another purchase less compelling."),
        ("Old Mutual", "Low-Medium","Launching retail bank",
         "SME entry via Yoco scale but not yet a proven strategic priority."),
    ]

    global_buyers = [
        ("Stripe",      "Very High", "No recent Africa acquisitions",
         "Africa entry via Yoco's 110k merchant POS network and SME payments rails."),
        ("Adyen",       "Very High", "Scaling enterprise globally",
         "African SME acquiring to complement enterprise focus. Would position Yoco as Africa's iZettle."),
        ("Rapyd",       "High",      "Acquired PayU GPO $610M (2023)",
         "Yoco POS complements digital stack. Short-term focus on PayU integration limits near-term appetite."),
        ("Experian",    "High",      "Acquired Compuscan SA $263M (2019)",
         "Transaction data enhances SME credit scoring. Direct outreach already made per Bruwer analysis."),
        ("Nubank",      "High",      "Pan-African expansion signals",
         "Pan-African growth ambitions. Yoco fits SME banking strategy."),
        ("Zoho",        "Medium",    "Offices in Nigeria and Kenya",
         "Full SME OS if POS integrated. Prefers organic growth over acquisition."),
        ("Shopify",     "Medium",    "Scaling POS globally",
         "Omnichannel seller ecosystem in Africa. Favours global tech over regional platforms."),
        ("Amazon",      "Low-Medium","Launched Amazon.co.za (2024)",
         "Enable card and QR acceptance for small sellers linking in-store to marketplace."),
        ("TransUnion",  "Low-Medium","SME data solutions SA (2023)",
         "Real-time merchant data overlap with existing bank and telco feeds."),
    ]

    affinity_cache = st.session_state.get("yoco_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="yoco_affinity_sync"):
            _api_key = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = [b[0] for b in local_buyers] + [g[0] for g in global_buyers]
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["yoco_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key)
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols = st.columns([2, 2, 3, 1, 2])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global = st.tabs(["Local Buyers", "Global Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_yoco_" + name.replace(" ", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_yoco_" + name.replace(" ", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Next Steps Generator ──────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 4px 0;letter-spacing:.3px'>Next Steps Generator</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:14px'>"
        "Tick buyers to re-engage above, then generate a prioritized outreach plan.</div>",
        unsafe_allow_html=True,
    )

    _BUYER_ACTIONS = {
        "Capitec":   "Request meeting with SME banking lead. Pitch profitability milestone and 110k merchant base as de-risked acquisition.",
        "Vodacom":   "Use Lula relationship for warm intro. Frame initial conversation as partnership exploration.",
        "FNB":       "Initiate conversation via existing merchant banking relationship. Explore strategic partnership or M&A dialogue.",
        "TymeBank":  "Re-engage CEO directly. Address profitability concerns with latest financial data.",
        "Experian":  "Follow up on prior outreach. Propose data partnership as first step toward deeper strategic conversation.",
        "Stripe":    "Approach via investment banking intermediary. Frame Yoco as Africa market entry vehicle.",
        "Adyen":     "Approach via investment banking intermediary. Frame Yoco as Africa market entry vehicle.",
        "Rapyd":     "Revisit once PayU integration settles mid-2026. Flag for Q4 outreach.",
    }
    _PRIORITY_ORDER = [b[0] for b in local_buyers] + [b[0] for b in global_buyers]

    if st.button("Generate Q3 2026 Exit Actions for Yoco"):
        ticked = [
            name for name in _PRIORITY_ORDER
            if st.session_state.get("engage_yoco_" + name.replace(" ", ""), False)
        ]

        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name}.")
                st.markdown(
                    f"<div style='padding:10px 14px;margin-bottom:8px;background:#FFFFFF;"
                    f"border:1px solid #D4D5CE;border-radius:8px'>"
                    f"<span style='font-weight:700;color:#2C2C2A'>{name}</span>"
                    f"<span style='color:#2C2C2A;margin-left:10px'>{action}</span></div>",
                    unsafe_allow_html=True,
                )
        else:
            st.markdown(
                f"<div style='color:{MUTED};font-size:13px'>Tick at least one buyer above to generate actions.</div>",
                unsafe_allow_html=True,
            )

        st.markdown("#### GP-Led Secondary — Investor Universe")
        secondary_investors = [
            ("Pantheon Ventures",    "Active GP-led secondary buyer, strong fintech exposure"),
            ("Lexington Partners",   "Large secondary fund with emerging market appetite"),
            ("HarbourVest Partners", "Active in African tech secondaries"),
            ("Verdane",              "European growth equity with fintech focus"),
            ("NewQuest Capital",     "Asia-Pacific secondary specialist with Africa interest"),
            ("TR Capital",           "Emerging market secondary specialist"),
        ]
        for inv_name, inv_desc in secondary_investors:
            st.markdown(
                f"<div style='padding:8px 14px;margin-bottom:6px;background:#FFFFFF;"
                f"border:1px solid #D4D5CE;border-radius:8px'>"
                f"<span style='font-weight:700;color:#2C2C2A'>{inv_name}</span>"
                f"<span style='color:{MUTED};margin-left:10px;font-size:13px'>{inv_desc}</span></div>",
                unsafe_allow_html=True,
            )


# ── Exit Tracking tab ─────────────────────────────────────────────────────────

def render_exit_tab(info: pd.Series, company_id: int) -> None:
    company_name = info["name"]
    sector       = str(info.get("sector", "")).lower()
    _today       = datetime.utcnow()
    cur_q        = f"Q{(_today.month - 1) // 3 + 1} {_today.year}"

    # ── Yoco custom exit tab ───────────────────────────────────────────────────
    if company_name == "Yoco":
        _render_yoco_exit_tab()
        return

    LIKELIHOOD_OPTS = ["Exploratory", "Active", "Advanced", "On Hold"]
    STATUS_OPTS     = ["Not Started", "Warm", "Active", "Passed"]
    TYPE_OPTS       = ["Strategic", "Financial", "Adjacent"]

    LIKELIHOOD_COLORS = {
        "Exploratory": (BLUE,      "#1565C0"),
        "Active":      (GREEN,     "#2E7D32"),
        "Advanced":    ("#D1FAE5", "#065F46"),
        "On Hold":     ("#F5F5F5", MUTED),
    }

    def _sh(text):
        st.markdown(
            f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
            f"margin:20px 0 12px 0;letter-spacing:.3px'>{text}</div>",
            unsafe_allow_html=True,
        )

    # ── Section 1: Exit Pathways ───────────────────────────────────────────────
    _sh("Exit Pathways")

    pathways = _exit_pathways_load(company_id)
    if not pathways:
        for pw in _suggest_exit_pathways(company_name, sector):
            _exit_pathway_save(company_id, None, pw["pathway_name"],
                               pw["likelihood"], pw["estimated_timeline"], pw["notes"])
        pathways = _exit_pathways_load(company_id)

    for pw in pathways:
        pid   = pw["id"]
        lhood = pw["likelihood"] if pw["likelihood"] in LIKELIHOOD_OPTS else "Exploratory"
        badge_bg, badge_fg = LIKELIHOOD_COLORS.get(lhood, (BLUE, "#1565C0"))

        with st.container(border=True):
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:10px;margin-bottom:6px'>"
                f"<span style='font-size:15px;font-weight:600;color:{BLACK}'>"
                f"{pw['pathway_name']}</span>"
                f"<span style='background:{badge_bg};color:{badge_fg};border-radius:4px;"
                f"padding:2px 8px;font-size:11px;font-weight:600'>{lhood}</span>"
                f"</div>",
                unsafe_allow_html=True,
            )
            with st.form(f"pw_{company_id}_{pid}", clear_on_submit=False):
                c1, c2, c3 = st.columns([3, 2, 2])
                name       = c1.text_input("Pathway name",  value=pw["pathway_name"])
                likelihood = c2.selectbox("Likelihood",     LIKELIHOOD_OPTS,
                                          index=LIKELIHOOD_OPTS.index(lhood))
                timeline   = c3.text_input("Est. timeline", value=pw["estimated_timeline"] or "",
                                            placeholder="e.g. 3–5 years")
                notes = st.text_area("Notes", value=pw["notes"] or "", height=80,
                                     placeholder="Context, conditions, next steps…")
                bs, bd, _ = st.columns([1, 1, 6])
                if bs.form_submit_button("Save",   use_container_width=True):
                    _exit_pathway_save(company_id, pid, name, likelihood, timeline, notes)
                    st.rerun()
                if bd.form_submit_button("Delete", use_container_width=True):
                    _exit_pathway_delete(pid)
                    st.rerun()

    with st.expander("＋ Add pathway"):
        with st.form(f"add_pw_{company_id}", clear_on_submit=True):
            c1, c2, c3 = st.columns([3, 2, 2])
            new_name     = c1.text_input("Pathway name",  placeholder="e.g. Strategic Acquisition")
            new_lhood    = c2.selectbox("Likelihood",     LIKELIHOOD_OPTS)
            new_timeline = c3.text_input("Est. timeline", placeholder="e.g. 3–5 years")
            new_notes    = st.text_area("Notes", height=80)
            if st.form_submit_button("Add pathway"):
                if new_name.strip():
                    _exit_pathway_save(company_id, None, new_name.strip(),
                                       new_lhood, new_timeline, new_notes)
                    st.rerun()

    # ── Section 2: Affinity CRM Sync ──────────────────────────────────────────
    _sh("Affinity CRM Sync")

    sync_key = f"crm_sync_results_{company_id}"
    if sync_key not in st.session_state:
        st.session_state[sync_key] = None

    if st.button("Sync from Affinity + Slack", key=f"crm_sync_{company_id}"):
        with st.spinner("Fetching from Affinity and Slack…"):
            import concurrent.futures
            aff_items  = []
            slk_items  = []
            aff_error  = ""
            slk_error  = ""

            def _fetch_aff():
                return fetch_affinity_interactions(company_name)

            def _fetch_slk():
                return fetch_slack_messages(company_name)

            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                fut_aff = pool.submit(_fetch_aff)
                fut_slk = pool.submit(_fetch_slk)
                try:
                    aff_items = fut_aff.result()
                except Exception as e:
                    aff_error = str(e)
                try:
                    slk_items = fut_slk.result()
                except Exception as e:
                    slk_error = str(e)

            if aff_error:
                st.warning(f"Affinity: {aff_error}")
            if slk_error:
                st.warning(f"Slack: {slk_error}")

            combined = aff_items + slk_items
            try:
                relevant = classify_exit_relevant(combined)
            except Exception as e:
                st.error(f"Classification failed: {e}")
                relevant = []

            st.session_state[sync_key] = {
                "aff_count": len(aff_items),
                "slk_count": len(slk_items),
                "relevant":  relevant,
            }

    sync_data = st.session_state[sync_key]
    if sync_data is not None:
        aff_n    = sync_data["aff_count"]
        slk_n    = sync_data["slk_count"]
        relevant = sync_data["relevant"]
        st.markdown(
            f"<div style='font-size:13px;color:{MUTED};margin-bottom:12px'>"
            f"{aff_n} Affinity notes + {slk_n} Slack messages found, "
            f"{len(relevant)} exit-relevant total</div>",
            unsafe_allow_html=True,
        )

        # Auto-add acquirer hints to buyer universe
        hints = [r["acquirer_hint"] for r in relevant if r.get("acquirer_hint")]
        if hints:
            buyers_df_now  = _buyer_tracking_load(company_id)
            existing_names = set(buyers_df_now["acquirer_name"].str.strip().str.lower())
            added = []
            for hint in hints:
                if hint.strip().lower() not in existing_names:
                    contact_date = next(
                        (r["date"] for r in relevant if r.get("acquirer_hint") == hint), ""
                    )
                    new_row = pd.DataFrame([{
                        "acquirer_name":      hint,
                        "acquirer_type":      "Strategic",
                        "relationship_owner": "",
                        "last_contact_date":  contact_date,
                        "status":             "Warm",
                    }])
                    buyers_df_now = pd.concat([buyers_df_now, new_row], ignore_index=True)
                    existing_names.add(hint.strip().lower())
                    added.append(hint)
            if added:
                _buyer_tracking_replace(company_id, buyers_df_now)
                st.success(f"Auto-added to buyer universe: {', '.join(added)}")

        # Show exit-relevant interactions
        if relevant:
            for item in relevant:
                src = item.get("source", "")
                if src == "slack":
                    badge_bg, badge_fg, border = "#F0E6FF", "#4A154B", "#9C27B0"
                else:
                    badge_bg, badge_fg, border = "#C5E5FF", "#1565C0", "#1565C0"
                src_badge = (
                    f"<span style='background:{badge_bg};color:{badge_fg};"
                    f"border-radius:3px;padding:1px 6px;font-size:10px;"
                    f"font-weight:600;margin-right:6px'>"
                    f"{'Slack' if src == 'slack' else 'Affinity'}</span>"
                )
                hint_badge = (
                    f" · <span style='color:#1565C0;font-weight:600'>"
                    f"{item['acquirer_hint']}</span>"
                    if item.get("acquirer_hint") else ""
                )
                st.markdown(
                    f"<div style='border-left:3px solid {border};padding:8px 12px;"
                    f"margin-bottom:8px;background:#F8FAFF;border-radius:0 4px 4px 0'>"
                    f"<div style='font-size:11px;color:{MUTED};margin-bottom:3px'>"
                    f"{src_badge}{item['date']} · {item.get('type','')} · "
                    f"{item.get('person_name','')}{hint_badge}</div>"
                    f"<div style='font-size:13px;color:{BLACK}'>{item.get('summary','')}</div>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.markdown(
                f"<div style='font-size:13px;color:{MUTED}'>No exit-relevant "
                f"interactions found.</div>",
                unsafe_allow_html=True,
            )

    # ── Section 3: Buyer Universe ──────────────────────────────────────────────
    _sh("Buyer Universe")

    buyers_df = _buyer_tracking_load(company_id)
    if buyers_df.empty:
        seed = pd.DataFrame(_suggest_buyers(sector))
        _buyer_tracking_replace(company_id, seed)
        buyers_df = _buyer_tracking_load(company_id)

    display_cols = ["acquirer_name", "acquirer_type", "relationship_owner",
                    "last_contact_date", "status"]
    display_df = (
        buyers_df[display_cols].copy()
        if all(c in buyers_df.columns for c in display_cols)
        else pd.DataFrame(columns=display_cols)
    )

    with st.form(f"buyer_form_{company_id}"):
        edited = st.data_editor(
            display_df,
            column_config={
                "acquirer_name":      st.column_config.TextColumn(
                    "Acquirer", width="large"),
                "acquirer_type":      st.column_config.SelectboxColumn(
                    "Type", options=TYPE_OPTS, width="small"),
                "relationship_owner": st.column_config.TextColumn(
                    "Relationship Owner", width="medium"),
                "last_contact_date":  st.column_config.TextColumn(
                    "Last Contact", width="small",
                    help="Format: YYYY-MM-DD or free text"),
                "status":             st.column_config.SelectboxColumn(
                    "Status", options=STATUS_OPTS, width="small"),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
        )
        if st.form_submit_button("Save buyer universe"):
            _buyer_tracking_replace(company_id, edited)
            st.success("Buyer universe saved.")
            st.rerun()

    # ── Section 3: Quarterly Actions ──────────────────────────────────────────
    _sh(f"Quarterly Actions — {cur_q}")

    qa = _quarterly_actions_load(company_id, cur_q)

    with st.container(border=True):
        with st.form(f"qa_{company_id}_{cur_q.replace(' ', '_')}"):
            c1, c2, c3 = st.columns(3)
            col_cfg = [
                (c1, "Planned Actions",   "planned_actions",   "Actions planned for this quarter…"),
                (c2, "Completed",         "completed_actions", "Actions completed this quarter…"),
                (c3, "Carry Forward",     "carry_forward",     "Items to carry into next quarter…"),
            ]
            text_vals = {}
            for col, hdr, key, ph in col_cfg:
                with col:
                    st.markdown(
                        f"<div style='font-size:12px;font-weight:600;color:{BLACK};"
                        f"margin-bottom:6px'>{hdr}</div>",
                        unsafe_allow_html=True,
                    )
                    text_vals[key] = st.text_area(
                        hdr, value=qa[key], height=200,
                        label_visibility="collapsed", placeholder=ph,
                    )
            if st.form_submit_button("Save actions", use_container_width=False):
                _quarterly_actions_save(
                    company_id, cur_q,
                    text_vals["planned_actions"],
                    text_vals["completed_actions"],
                    text_vals["carry_forward"],
                )
                st.success("Actions saved.")


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

    # ── Filter bar ────────────────────────────────────────────────────────────
    all_sectors = sorted(companies["sector"].dropna().unique().tolist())
    sector_options = ["All"] + [sector_label(s) for s in all_sectors]

    selected_fund = st.radio(
        "Filter by fund",
        options=["All Funds", "Fund I", "Fund II", "Fund III"],
        index=0,
        horizontal=True,
        label_visibility="collapsed",
    )

    filter_col, sort_col = st.columns([4, 1])
    with filter_col:
        selected_sector = st.radio(
            "Filter by sector",
            options=sector_options,
            index=0,
            horizontal=True,
            label_visibility="collapsed",
        )
    with sort_col:
        st.text_input(
            "Search",
            placeholder="Search by name...",
            key="company_search",
            label_visibility="collapsed",
        )

    # Apply fund filter, then sector filter, then name search
    filtered = companies.copy()
    if selected_fund != "All Funds":
        filtered = filtered[filtered["fund"] == selected_fund]
    if selected_sector != "All":
        filtered = filtered[filtered["sector"].apply(sector_label) == selected_sector]
    filtered = filtered[filtered["name"].str.contains(st.session_state.get("company_search", ""), case=False, na=False)]

    filtered = filtered.sort_values("name")

    n_showing = len(filtered)
    st.markdown(
        f"<div style='font-size:11px;color:{MUTED};letter-spacing:.04em;"
        f"margin-bottom:14px'>{n_showing} of {len(companies)} companies</div>",
        unsafe_allow_html=True,
    )

    # ── Card grid — 3 columns ─────────────────────────────────────────────────
    def _sector_metric_pair(row: pd.Series) -> tuple[tuple, tuple]:
        """Return two (label, value, color) tuples for the sector-specific bottom row."""
        sector = str(row.get("sector", "")).lower()

        def _pct_color(v, positive_is_good=True):
            if _is_null(v): return MUTED
            return ("#2E7D32" if float(v) > 0 else "#C62828") if positive_is_good \
                else ("#C62828" if float(v) > 0 else "#2E7D32")

        if sector == "lending":
            npl = row.get("npl_rate_pct")
            loan = row.get("loan_book_gross_usd")
            return (
                ("NPL Rate",   fmt_pct(npl),  _pct_color(npl, positive_is_good=False)),
                ("Loan Book",  fmt_usd(loan), BLACK),
            )
        elif sector == "wealth_management":
            aum = row.get("aum_usd")
            return (
                ("AUM",        fmt_usd(aum),  BLACK),
                ("EBITDA Mgn", fmt_pct(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct")),
                 _pct_color(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct"))),
            )
        elif sector == "payments":
            tpv = row.get("tpv_usd") or row.get("gmv_usd")
            tpv_lbl = "TPV" if not _is_null(row.get("tpv_usd")) else "GMV"
            return (
                (tpv_lbl,     fmt_usd(tpv),  BLACK),
                ("EBITDA Mgn", fmt_pct(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct")),
                 _pct_color(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct"))),
            )
        elif sector == "insurtech":
            return (
                ("EBITDA Mgn", fmt_pct(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct")),
                 _pct_color(row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct"))),
                ("Customers",  fmt_int(row.get("customer_count")), BLACK),
            )
        else:
            em = row.get("ltm_ebitda_margin_pct") or row.get("ebitda_margin_pct")
            cust = row.get("customer_count")
            return (
                ("EBITDA Mgn", fmt_pct(em),   _pct_color(em)),
                ("Customers",  fmt_int(cust),  BLACK),
            )

    def _render_card(col, row: pd.Series) -> None:
        """Render a single company card inside a Streamlit column."""
        cid        = int(row["id"])
        name       = row["name"]
        sl         = sector_label(row.get("sector", ""))
        country    = row.get("hq_country", "")
        ltm_val    = row.get("ltm_revenue")
        ltm_lbl    = row.get("ltm_label", "")
        pt         = row.get("period_type", "monthly")
        period_lbl = fmt_period_label(row.get("period_end_date"), pt)
        asof       = as_of(row.get("period_end_date"))

        # LTM revenue
        rev_str = fmt_usd(ltm_val)
        if ltm_lbl == "LTM":
            basis   = "12 mo." if pt == "monthly" else "4 qtrs." if pt == "quarterly" else "annual"
            rev_sub = f"LTM · {basis}"
        elif ltm_lbl == "ARR (est.)":
            rev_sub = "ARR (est.)"
        else:
            rev_sub = ""
        if period_lbl and not _is_null(ltm_val):
            rev_str = f"{rev_str} ({period_lbl})"

        # Gross margin
        gm = row.get("ltm_gross_margin_pct")
        if _is_null(gm): gm = row.get("gross_margin_pct")
        gm_str   = fmt_pct(gm)
        gm_color = "#2E7D32" if (not _is_null(gm) and float(gm) > 50) else BLACK

        # Revenue growth
        gtxt, gcol = fmt_growth(row.get("revenue_growth_pct"))

        # Sector-specific metric
        (lbl3, val3, col3), _ = _sector_metric_pair(row)

        with col:
            with st.container(border=True):
                # ── Header: name + sector tag ─────────────────────────────
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;"
                    f"align-items:flex-start;margin-bottom:2px'>"
                    f"<div>"
                    f"<div style='font-size:16px;font-weight:800;color:{BLACK};"
                    f"letter-spacing:-0.3px;line-height:1.2'>{name}</div>"
                    f"<div style='font-size:10px;font-weight:600;color:{MUTED};"
                    f"text-transform:uppercase;letter-spacing:.06em;margin-top:2px'>"
                    f"{country}</div>"
                    f"</div>"
                    f"<span style='background:{BLUE};color:{BLACK};border-radius:99px;"
                    f"padding:3px 10px;font-size:10px;font-weight:700;"
                    f"letter-spacing:.04em;white-space:nowrap;flex-shrink:0;"
                    f"margin-left:8px;margin-top:2px'>{sl}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

                st.markdown(f"<hr style='margin:8px 0;border-color:{BORDER}'>",
                            unsafe_allow_html=True)

                # ── Metrics: 2 columns x 2 rows ───────────────────────────
                m1, m2 = st.columns(2)
                m3, m4 = st.columns(2)

                with m1:
                    st.markdown(
                        f"<div style='font-size:9px;font-weight:700;letter-spacing:.12em;"
                        f"text-transform:uppercase;color:{MUTED};margin-bottom:1px'>LTM Revenue</div>"
                        f"<div style='font-size:15px;font-weight:800;color:{BLACK}'>{rev_str}</div>"
                        f"<div style='font-size:9px;color:{MUTED};margin-top:1px'>{rev_sub}</div>",
                        unsafe_allow_html=True,
                    )
                with m2:
                    st.markdown(
                        f"<div style='font-size:9px;font-weight:700;letter-spacing:.12em;"
                        f"text-transform:uppercase;color:{MUTED};margin-bottom:1px'>Gross Margin</div>"
                        f"<div style='font-size:15px;font-weight:800;color:{gm_color}'>{gm_str}</div>",
                        unsafe_allow_html=True,
                    )
                with m3:
                    st.markdown(
                        f"<div style='font-size:9px;font-weight:700;letter-spacing:.12em;"
                        f"text-transform:uppercase;color:{MUTED};margin-bottom:1px'>Rev Growth</div>"
                        f"<div style='font-size:15px;font-weight:800;color:{gcol}'>{gtxt}</div>",
                        unsafe_allow_html=True,
                    )
                with m4:
                    st.markdown(
                        f"<div style='font-size:9px;font-weight:700;letter-spacing:.12em;"
                        f"text-transform:uppercase;color:{MUTED};margin-bottom:1px'>{lbl3}</div>"
                        f"<div style='font-size:15px;font-weight:800;color:{col3}'>{val3}</div>",
                        unsafe_allow_html=True,
                    )

                st.markdown(f"<hr style='margin:8px 0;border-color:{BORDER}'>",
                            unsafe_allow_html=True)

                # ── Footer: as of date + view button ─────────────────────
                st.markdown(
                    f"<div style='font-size:9px;color:{MUTED};font-weight:600;"
                    f"letter-spacing:.06em;margin-bottom:4px'>As of {asof}</div>",
                    unsafe_allow_html=True,
                )
                if st.button("View company →", key=f"co_{cid}", use_container_width=True):
                    st.session_state.page = "detail"
                    st.session_state.company_id = cid
                    st.rerun()

    # Render cards in rows of 3
    rows_iter = list(filtered.iterrows())
    for i in range(0, len(rows_iter), 3):
        chunk = rows_iter[i:i+3]
        cols  = st.columns(3)
        for col, (_, row) in zip(cols, chunk):
            _render_card(col, row)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

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
    _tab_names  = ["Performance", "Benchmarking", "Exit Tracking"] + (["Upload Data"] if _has_upload else [])
    _tabs       = st.tabs(_tab_names)
    tab_perf    = _tabs[0]
    tab_bench   = _tabs[1]
    tab_exit    = _tabs[2]
    tab_upload  = _tabs[3] if _has_upload else None

    with tab_perf:
        # ── Chart palette ─────────────────────────────────────────────────────
        C_REVENUE  = "#378ADD"
        C_GM       = "#1D9E75"
        C_EBITDA_P = "#2E7D32"   # positive EBITDA line
        C_EBITDA_N = "#C62828"   # negative EBITDA line
        C_CLIENTS  = "#7F77DD"
        C_TPV_GMV  = "#378ADD"
        CHART_H    = 280
        CFG        = {"displayModeBar": False}

        # ── 24-month window ───────────────────────────────────────────────────
        kpis_sorted = kpis.copy()
        kpis_sorted["period_end_date"] = pd.to_datetime(kpis_sorted["period_end_date"], errors="coerce")
        kpis_sorted = kpis_sorted.dropna(subset=["period_end_date"]).sort_values("period_end_date")
        if len(kpis_sorted) > 0:
            cutoff = kpis_sorted["period_end_date"].max() - pd.DateOffset(months=24)
            kpis_24 = kpis_sorted[kpis_sorted["period_end_date"] >= cutoff].copy()
        else:
            kpis_24 = kpis_sorted.copy()

        # ── Chart builders ────────────────────────────────────────────────────
        def apply_executive_style(fig, title, y_fmt="number"):
            t = title.lower()
            if "gross margin" in t:
                line_color = "#2E7D32"
            elif "customer" in t or "client" in t:
                line_color = "#93A3A1"
            else:
                line_color = "#2C2C2A"

            lc_r = int(line_color[1:3], 16)
            lc_g = int(line_color[3:5], 16)
            lc_b = int(line_color[5:7], 16)

            last_x = last_y = None
            for trace in fig.data:
                mode = getattr(trace, "mode", "") or ""
                if "lines" in mode:
                    update_kwargs = dict(
                        line=dict(width=2.5, color=line_color),
                        mode="lines",
                        marker=dict(size=0, opacity=0),
                    )
                    if getattr(trace, "fill", None):
                        update_kwargs["fillcolor"] = f"rgba({lc_r},{lc_g},{lc_b},0.08)"
                    trace.update(update_kwargs)
                    if trace.x is not None and len(trace.x) > 0:
                        last_x = trace.x[-1]
                        last_y = trace.y[-1] if trace.y is not None and len(trace.y) > 0 else None
                elif mode == "none" and getattr(trace, "fill", None) == "tozeroy":
                    if trace.y is not None and any(v < 0 for v in trace.y if v is not None):
                        trace.update(fillcolor="rgba(255,138,133,0.08)")

            if last_x is not None and last_y is not None:
                try:
                    float(last_y)
                    fig.add_trace(go.Scatter(
                        x=[last_x], y=[last_y],
                        mode="markers",
                        marker=dict(size=7, color=line_color, line=dict(width=2, color="white")),
                        showlegend=False,
                        hoverinfo="skip",
                    ))
                except (TypeError, ValueError):
                    pass

            fig.update_layout(
                font=dict(family="DM Sans, Trebuchet MS, sans-serif", color="#2C2C2A"),
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=20, r=20, t=40, b=20),
                title=dict(
                    text=title,
                    font=dict(size=14, color="#2C2C2A"),
                    x=0, xanchor="left",
                ),
                legend=dict(
                    orientation="h", yanchor="bottom", y=-0.2,
                    xanchor="left", x=0, font=dict(size=11),
                ),
                hovermode="x unified",
                height=CHART_H,
                showlegend=False,
            )
            fig.update_xaxes(
                showgrid=False,
                showline=False,
                tickfont=dict(size=11, color="#93A3A1"),
                tickcolor="#93A3A1",
                tickformat="%b %Y",
            )
            fig.update_yaxes(
                showgrid=True,
                gridcolor="#EFF0EA",
                gridwidth=1,
                showline=False,
                tickfont=dict(size=11, color="#93A3A1"),
                zeroline=True,
                zerolinecolor="#D4D5CE",
                zerolinewidth=1,
                ticksuffix="%" if y_fmt == "pct" else "",
                tickprefix="$" if y_fmt == "usd" else "",
            )

        def _chart_card(fig):
            st.plotly_chart(fig, use_container_width=True, config=CFG)

        def _simple_chart(col, y_fmt, title, line_color):
            sub = kpis_24[[col, "period_end_date"]].dropna()
            if len(sub) < 2:
                return None
            hover = "$%{y:,.0f}" if y_fmt == "usd" else ("%{y:.1f}%" if y_fmt == "pct" else "%{y:,.0f}")
            r, g, b = int(line_color[1:3], 16), int(line_color[3:5], 16), int(line_color[5:7], 16)
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=sub["period_end_date"], y=sub[col],
                mode="lines+markers",
                line=dict(color=line_color, width=2),
                marker=dict(size=4, color=line_color),
                fill="tozeroy",
                fillcolor=f"rgba({r},{g},{b},0.08)",
                hovertemplate=f"%{{x|%b %Y}}<br>{hover}<extra></extra>",
            ))
            apply_executive_style(fig, title, y_fmt)
            return fig

        def _ebitda_chart(col, y_fmt, title):
            sub = kpis_24[[col, "period_end_date"]].dropna()
            if len(sub) < 2:
                return None
            hover = "$%{y:,.0f}" if y_fmt == "usd" else "%{y:.1f}%"
            vals = sub[col].tolist()
            dates = sub["period_end_date"].tolist()

            # Single line colored by sign of most-recent value
            is_pos = vals[-1] >= 0 if vals else True
            line_color = C_EBITDA_P if is_pos else C_EBITDA_N

            fig = go.Figure()
            # Positive fill (above zero)
            pos_y = [max(v, 0) for v in vals]
            fig.add_trace(go.Scatter(
                x=dates, y=pos_y, mode="none",
                fill="tozeroy", fillcolor="rgba(46,125,50,0.10)",
                showlegend=False, hoverinfo="skip",
            ))
            # Negative fill (below zero)
            neg_y = [min(v, 0) for v in vals]
            fig.add_trace(go.Scatter(
                x=dates, y=neg_y, mode="none",
                fill="tozeroy", fillcolor="rgba(198,40,40,0.10)",
                showlegend=False, hoverinfo="skip",
            ))
            # Main line
            fig.add_trace(go.Scatter(
                x=dates, y=vals,
                mode="lines+markers",
                line=dict(color=line_color, width=2),
                marker=dict(size=4, color=line_color),
                hovertemplate=f"%{{x|%b %Y}}<br>{hover}<extra></extra>",
            ))

            # Annotation: first profitable month or best EBITDA
            ebitda_series = pd.Series(vals, index=dates)
            annotations = []
            pos_months = ebitda_series[ebitda_series > 0]
            if not pos_months.empty:
                first_pos_date = pos_months.index[0]
                all_pos_before = ebitda_series.loc[:first_pos_date]
                if (all_pos_before.iloc[:-1] <= 0).all():
                    annotations.append(dict(
                        x=first_pos_date, y=ebitda_series[first_pos_date],
                        text="First profitable", showarrow=True, arrowhead=2,
                        arrowcolor="#D5FA94", font=dict(size=11, color="#2C2C2A"),
                        bgcolor="white", bordercolor="#D4D5CE", borderwidth=1,
                        borderpad=3, ax=0, ay=-30,
                    ))
                else:
                    best_date = ebitda_series.idxmax()
                    best_val  = ebitda_series[best_date]
                    if best_date == ebitda_series.index[-1]:
                        annotations.append(dict(
                            x=best_date, y=best_val,
                            text="Best EBITDA", showarrow=True, arrowhead=2,
                            arrowcolor="#D5FA94", font=dict(size=11, color="#2C2C2A"),
                            bgcolor="white", bordercolor="#D4D5CE", borderwidth=1,
                            borderpad=3, ax=0, ay=-30,
                        ))

            apply_executive_style(fig, title, y_fmt)
            if annotations:
                fig.update_layout(annotations=annotations)
            return fig

        def _section_header(text):
            st.markdown(
                f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
                f"margin:18px 0 10px 0'>{text}</div>",
                unsafe_allow_html=True,
            )

        # ── Financial Performance ─────────────────────────────────────────────
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        _section_header("Financial Performance")

        # Revenue — full width
        rev_fig = _simple_chart("revenue_usd", "usd", "Revenue (USD)", C_REVENUE)
        if rev_fig:
            _chart_card(rev_fig)
        else:
            _no_data_box("No revenue data")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Gross Margin and EBITDA Margin — 2 columns
        ca, cb = st.columns(2)
        with ca:
            fig = _simple_chart("gross_margin_pct", "pct", "Gross Margin %", C_GM)
            if fig: _chart_card(fig)
            else:   _no_data_box("No gross margin data")
        with cb:
            fig = _ebitda_chart("ebitda_margin_pct", "pct", "EBITDA Margin %")
            if fig: _chart_card(fig)
            else:   _no_data_box("No EBITDA margin data")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # EBITDA (USD) — full column within 2-col grid (leave right col empty)
        cc, cd = st.columns(2)
        with cc:
            fig = _ebitda_chart("ebitda_usd", "usd", "EBITDA (USD)")
            if fig: _chart_card(fig)
            else:   _no_data_box("No EBITDA data")

        # Customer / Active Clients — full width
        if "customer_count" in kpis.columns and kpis["customer_count"].notna().any():
            cust_col, cust_lbl = "customer_count", "Customer Count"
        elif "active_clients_count" in kpis.columns and kpis["active_clients_count"].notna().any():
            cust_col, cust_lbl = "active_clients_count", "Active Clients"
        else:
            cust_col, cust_lbl = None, None

        if cust_col:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            fig = _simple_chart(cust_col, "number", cust_lbl, C_CLIENTS)
            if fig: _chart_card(fig)

        # ── Lending KPIs snapshot ─────────────────────────────────────────────
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
                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                _section_header("Lending KPIs (Latest Period)")
                snap_cols = st.columns(len(snapshot_vals))
                for col, (lbl, val_str) in zip(snap_cols, snapshot_vals):
                    col.metric(lbl, val_str)

        # ── Lending & credit metrics ──────────────────────────────────────────
        LENDING_METRICS = [
            ("loan_book_gross_usd", "Net Loan Portfolio (USD)", "usd"),
            ("par_30_pct",          "PAR 30+ %",                "pct"),
            ("par_90_pct",          "PAR 90 %",                 "pct"),
            ("npl_rate_pct",        "NPL Rate %",               "pct"),
            ("net_yield_pct",       "Net Yield %",              "pct"),
            ("nim_pct",             "Net Interest Margin %",    "pct"),
        ]
        lending_available = [
            (c, t, f) for c, t, f in LENDING_METRICS
            if c in kpis.columns and kpis[c].dropna().__len__() >= 2
        ]
        if lending_available:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            _section_header("Lending & Credit Metrics")
            for i in range(0, len(lending_available), 2):
                cols = st.columns(2)
                for j, col in enumerate(cols):
                    if i + j < len(lending_available):
                        c, t, f = lending_available[i + j]
                        lc = C_REVENUE if f == "usd" else C_GM
                        fig = _simple_chart(c, f, t, lc)
                        if fig:
                            with col: _chart_card(fig)

        # ── AUM ───────────────────────────────────────────────────────────────
        if "aum_usd" in kpis.columns and kpis["aum_usd"].dropna().__len__() >= 2:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            _section_header("Assets Under Management")
            ce, _ = st.columns(2)
            with ce:
                fig = _simple_chart("aum_usd", "usd", "Assets Under Management (USD)", C_REVENUE)
                if fig: _chart_card(fig)

        # ── Sector metrics ────────────────────────────────────────────────────
        OTHER_METRICS = [
            ("gmv_usd",                   "GMV (USD)",                  "usd", C_TPV_GMV),
            ("tpv_usd",                   "Total Payment Volume (USD)", "usd", C_TPV_GMV),
            ("arr_usd",                   "ARR (USD)",                  "usd", C_REVENUE),
            ("net_revenue_retention_pct", "Net Revenue Retention %",   "pct", C_GM),
        ]
        shown = {c for c, *_ in lending_available} | {"aum_usd"}
        other_available = [
            (c, t, f, lc) for c, t, f, lc in OTHER_METRICS
            if c not in shown and c in kpis.columns and kpis[c].dropna().__len__() >= 2
        ]
        if other_available:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            _section_header("Sector Metrics")
            for i in range(0, len(other_available), 2):
                cols = st.columns(2)
                for j, col in enumerate(cols):
                    if i + j < len(other_available):
                        c, t, f, lc = other_available[i + j]
                        fig = _simple_chart(c, f, t, lc)
                        if fig:
                            with col: _chart_card(fig)

    with tab_bench:
        render_benchmarking_tab(info, kpis, ltm_val, ltm_lbl, ltm_gm_pct, ltm_em_pct)

    with tab_exit:
        render_exit_tab(info, company_id)

    if tab_upload is not None:
        with tab_upload:
            render_upload_tab(info, company_id)
