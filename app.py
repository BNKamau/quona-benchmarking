import streamlit as st
import sqlite3
import psycopg2
import psycopg2.extras
import pandas as pd
import plotly.graph_objects as go
import anthropic
import os
from datetime import datetime, timedelta, timezone
from parsers.excel_parsers import PARSERS, SUPPORTED_COMPANIES

# ── SQLite compatibility shim for local dev ────────────────────────────────────
# Translates %s placeholders (psycopg2 style) to ? (sqlite3 style) so the
# same SQL works against both Supabase and the local SQLite fallback.

class _SQLiteShimCursor:
    def __init__(self, cur):
        self._cur = cur

    @staticmethod
    def _fix(sql):
        return sql.replace("%s", "?")

    def execute(self, sql, params=None):
        self._cur.execute(self._fix(sql), params or ())
        return self

    def fetchall(self):         return self._cur.fetchall()
    def fetchone(self):         return self._cur.fetchone()
    def fetchmany(self, n=1000): return self._cur.fetchmany(n)
    def close(self):            self._cur.close()
    def __iter__(self):         return iter(self._cur)

    @property
    def description(self): return self._cur.description

    @property
    def rowcount(self): return self._cur.rowcount


class _SQLiteShim:
    """sqlite3 connection wrapper that accepts psycopg2-style %s placeholders."""

    def __init__(self, path):
        self._c = sqlite3.connect(path, check_same_thread=False)
        self._c.execute("PRAGMA journal_mode=WAL")
        self._c.execute("PRAGMA synchronous=NORMAL")

    @staticmethod
    def _fix(sql):
        return sql.replace("%s", "?")

    def cursor(self):
        return _SQLiteShimCursor(self._c.cursor())

    def execute(self, sql, params=None):
        return _SQLiteShimCursor(self._c.execute(self._fix(sql), params or ()))

    def executemany(self, sql, seq):
        self._c.executemany(self._fix(sql), seq)

    def commit(self):    self._c.commit()
    def close(self):     self._c.close()
    def __enter__(self): return self
    def __exit__(self, *_): self._c.close()


class _PgConn:
    """psycopg2 connection wrapper that adds .execute() returning a cursor,
    matching the sqlite3 / _SQLiteShim interface so all call sites are
    unchanged: conn.execute(sql, params).fetchone() works as expected."""

    def __init__(self, url: str):
        self._c = psycopg2.connect(url, connect_timeout=10)
        self._c.autocommit = False

    def cursor(self):
        return self._c.cursor()

    def execute(self, sql, params=None):
        cur = self._c.cursor()
        cur.execute(sql, params) if params else cur.execute(sql)
        return cur

    def executemany(self, sql, seq):
        cur = self._c.cursor()
        cur.executemany(sql, seq)
        cur.close()

    def commit(self):    self._c.commit()
    def close(self):     self._c.close()
    def __enter__(self): return self
    def __exit__(self, *_): self._c.close()


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

  /* Landing page cards */
  div[data-testid="stVerticalBlockBorderWrapper"] {{
      border-radius: 10px !important;
      border: 1px solid #D4D5CE !important;
      box-shadow: 0 2px 8px rgba(44,44,42,0.06) !important;
      transition: box-shadow 0.2s ease, transform 0.15s ease !important;
      background: white !important;
  }}
  div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
      box-shadow: 0 6px 20px rgba(44,44,42,0.12) !important;
      transform: translateY(-2px) !important;
  }}

  /* View company outlined button */
  button[data-testid*="co_"] > div > p {{
      font-size: 13px !important;
      font-weight: 600 !important;
      color: #2C2C2A !important;
  }}
  button[data-testid*="co_"] {{
      background: white !important;
      border: 1.5px solid #D4D5CE !important;
      border-radius: 6px !important;
      color: #2C2C2A !important;
      margin-top: 4px !important;
  }}
  button[data-testid*="co_"]:hover {{
      border-color: #2C2C2A !important;
      background: #EFF0EA !important;
  }}

  /* Suppress spell-check red underlines */
  div[data-testid="stMarkdownContainer"] {{ -webkit-user-modify: read-only; }}
  div[data-testid="stMarkdownContainer"] p {{ -webkit-spell-check: false !important; }}
  input, textarea, [contenteditable] {{ spellcheck: false !important; }}

  /* Filter radio spacing */
  div[data-testid="stRadio"] {{ margin-bottom: 0px !important; margin-top: 0px !important; }}

  /* Search box */
  div[data-testid="stTextInput"] input {{
      border: 1.5px solid #D4D5CE !important;
      border-radius: 6px !important;
      font-size: 13px !important;
      color: #2C2C2A !important;
      padding: 8px 12px !important;
  }}
  div[data-testid="stTextInput"] input:focus {{
      border-color: #2C2C2A !important;
      box-shadow: none !important;
  }}
</style>
""", unsafe_allow_html=True)

# ── DB helpers ─────────────────────────────────────────────────────────────────
# Absolute path — invariant of the process CWD so reads and writes always
# hit the same file, whether launched from the repo root, a parent dir, or
# a cloud runner that sets a different working directory.
_HERE    = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(_HERE, "benchmarking.db")

def _get_pg_conn():
    """Return a psycopg2 connection to Supabase. Falls back to local SQLite via
    a compatibility shim if SUPABASE_DB_URL is not in secrets (local dev)."""
    url = st.secrets.get("SUPABASE_DB_URL", "")
    if url:
        return _PgConn(url)
    return _SQLiteShim(DB_PATH)

def _conn():
    return _get_pg_conn()

# Companies that should always exist in the registry.
# Seeded automatically on first startup so uploads work on a fresh install.
_PORTFOLIO_COMPANIES = [
    ("Cowrywise", "portfolio", "wealth_management",          "savings_and_investment",        "NG", 2017, "b2c", "NGN"),
    ("Yoco",      "portfolio", "payments",                   "merchant_acquiring",             "ZA", 2015, "b2b", "ZAR"),
    ("Verto",     "portfolio", "payments",                   "cross_border_fx",                "NG", 2019, "b2b", "USD"),
    ("Enza",      "portfolio", "payments",                   "card_issuing_paas",              "KE", 2020, "b2b", "USD"),
    ("Lulalend",  "portfolio", "lending",                    "sme_lending",                   "ZA", 2014, "b2b", "ZAR"),
    ("Khazna",    "portfolio", "lending",                    "consumer_lending",               "EG", 2021, "b2c", "USD"),
    ("TWINCO",    "portfolio", "lending",                    "supply_chain_finance",           "ES", 2019, "b2b", "EUR"),
    ("MaxSoko",   "portfolio", "marketplace",                "ecommerce_embedded_finance",     "EG", 2015, "b2b", "USD"),
    ("SAVA",      "portfolio", "payments",                   "card_issuing_baas",              "ZA", 2022, "b2b", "ZAR"),
    ("AllLife",   "portfolio", "insurtech",                  "life_insurance",                 "ZA", 2004, "b2c", "ZAR"),
    ("OCTA",      "portfolio", "saas",                       "invoice_ar_automation",          "NG", 2023, "b2b", "USD"),
    ("Eseye",     "portfolio", "iot_infrastructure",         "managed_connectivity",           "GB", 2007, "b2b", "GBP"),
    ("POWER",     "portfolio", "lending",                    "earned_wage_access",             "US", 2020, "b2b", "USD"),
]


def _init_db() -> None:
    """Create all tables and seed the company registry on first run.

    Safe to call on every startup: CREATE TABLE IF NOT EXISTS and INSERT OR
    IGNORE ensure existing data is never touched.  ALTER TABLE ADD COLUMN
    is also idempotent (errors are swallowed).
    """
    conn = _conn()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS companies (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            name               TEXT    NOT NULL,
            type               TEXT    NOT NULL DEFAULT 'portfolio',
            sector             TEXT,
            sub_sector         TEXT,
            hq_country         TEXT,
            founded_year       INTEGER,
            business_model     TEXT,
            reporting_currency TEXT,
            fund               TEXT,
            notes              TEXT,
            created_at         TEXT    NOT NULL DEFAULT (datetime('now')),
            updated_at         TEXT    NOT NULL DEFAULT (datetime('now'))
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS kpi_snapshots (
            id                         INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id                 INTEGER NOT NULL REFERENCES companies(id),
            period_end_date            TEXT    NOT NULL,
            reporting_currency         TEXT    NOT NULL DEFAULT 'USD',
            fx_rate_to_usd             REAL,
            revenue_usd                REAL,
            gross_profit_usd           REAL,
            gross_margin_pct           REAL,
            ebitda_usd                 REAL,
            ebitda_margin_pct          REAL,
            arr_usd                    REAL,
            mrr_usd                    REAL,
            customer_count             INTEGER,
            active_clients_count       INTEGER,
            net_revenue_retention_pct  REAL,
            gross_churn_rate_pct       REAL,
            cac_usd                    REAL,
            ltv_usd                    REAL,
            loan_book_gross_usd        REAL,
            npl_rate_pct               REAL,
            par_30_pct                 REAL,
            par_90_pct                 REAL,
            net_yield_pct              REAL,
            cost_of_risk_pct           REAL,
            nim_pct                    REAL,
            leverage_ratio             REAL,
            aum_usd                    REAL,
            gmv_usd                    REAL,
            tpv_usd                    REAL,
            unique_borrowers_count     INTEGER,
            top_3_concentration_pct    REAL,
            insurance_policies_active  INTEGER,
            devices_connected          INTEGER,
            net_income_usd             REAL,
            net_margin_pct             REAL,
            revenue_growth_pct         REAL,
            notes                      TEXT,
            created_at                 TEXT    NOT NULL DEFAULT (datetime('now')),
            updated_at                 TEXT    NOT NULL DEFAULT (datetime('now')),
            UNIQUE (company_id, period_end_date)
        )
    """)

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

    conn.execute("""
        CREATE TABLE IF NOT EXISTS ipo_readiness (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id  INTEGER NOT NULL,
            item_key    TEXT    NOT NULL,
            status      TEXT    NOT NULL DEFAULT 'Not Started',
            notes       TEXT    DEFAULT '',
            updated_at  TEXT    NOT NULL DEFAULT (datetime('now')),
            UNIQUE(company_id, item_key)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS exit_documents (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id   INTEGER NOT NULL,
            doc_name     TEXT    NOT NULL,
            doc_type     TEXT    NOT NULL DEFAULT 'exit_planning',
            file_data    BLOB    NOT NULL,
            file_size    INTEGER,
            uploaded_by  TEXT    DEFAULT '',
            uploaded_at  TEXT    NOT NULL DEFAULT (datetime('now')),
            notes        TEXT    DEFAULT ''
        )
    """)

    conn.commit()

    # Idempotently add columns that were introduced after the initial schema.
    # Only sqlite3.OperationalError "duplicate column name" is expected; anything
    # else is logged so real problems are not silently hidden.
    def _add_col(table, col_def):
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_def}")
            conn.commit()
        except Exception as _e:
            msg = str(_e).lower()
            if "duplicate column" not in msg and "already exists" not in msg:
                print(f"[schema migration] ALTER TABLE {table} ADD COLUMN {col_def}: {_e}")

    for col_def in [
        "net_income_usd         REAL",
        "net_margin_pct         REAL",
        "revenue_growth_pct     REAL",
        "aum_usd                REAL",
        "gmv_usd                REAL",
        "active_clients_count   INTEGER",
        "par_30_pct             REAL",
        "par_90_pct             REAL",
        "top_3_concentration_pct REAL",
        "insurance_policies_active INTEGER",
        "tpv_usd               REAL",
        "devices_connected      INTEGER",
        "unique_borrowers_count INTEGER",
        "fund                   TEXT",
        "sub_sector             TEXT",
    ]:
        table = "kpi_snapshots" if col_def.split()[0] not in ("fund", "sub_sector") else "companies"
        _add_col(table, col_def)

    for col_def in ["source_url TEXT", "discovered_via TEXT", "discovered_at TIMESTAMP"]:
        _add_col("exit_comps", col_def)

    # Seed company registry only on a completely fresh DB.
    # We check count == 0 because companies has no UNIQUE constraint on name,
    # so INSERT OR IGNORE would silently duplicate rows on every startup.
    n_companies = conn.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
    if n_companies == 0:
        for (name, typ, sector, sub_sector, hq, year, biz, currency) in _PORTFOLIO_COMPANIES:
            conn.execute("""
                INSERT INTO companies
                    (name, type, sector, sub_sector, hq_country, founded_year,
                     business_model, reporting_currency)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (name, typ, sector, sub_sector, hq, year, biz, currency))
        conn.commit()
    conn.close()


# Skip on Supabase — tables already exist; _init_db uses SQLite-only DDL syntax.
if not st.secrets.get("SUPABASE_DB_URL", ""):
    _init_db()


# ── Exit document helpers ──────────────────────────────────────────────────────

def save_exit_document(company_id: int, doc_name: str, file_bytes: bytes,
                       doc_type: str = "exit_planning", notes: str = "") -> int:
    conn = _conn()
    cur = conn.cursor()
    cur.execute(
        "DELETE FROM exit_documents WHERE company_id = %s AND doc_name = %s",
        (company_id, doc_name),
    )
    is_pg = bool(st.secrets.get("SUPABASE_DB_URL", ""))
    file_param = psycopg2.Binary(file_bytes) if is_pg else file_bytes
    cur.execute(
        """INSERT INTO exit_documents
               (company_id, doc_name, doc_type, file_data, file_size, uploaded_at, notes)
           VALUES (%s, %s, %s, %s, %s, %s, %s)""",
        (company_id, doc_name, doc_type, file_param, len(file_bytes),
         datetime.utcnow().isoformat(), notes),
    )
    conn.commit()
    conn.close()
    return 0


def list_exit_documents(company_id: int) -> list:
    conn = _conn()
    cur = conn.cursor()
    cur.execute(
        """SELECT id, doc_name, doc_type, file_size, uploaded_at, notes
           FROM exit_documents WHERE company_id = %s ORDER BY uploaded_at DESC""",
        (company_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return [
        {"id": r[0], "doc_name": r[1], "doc_type": r[2],
         "file_size": r[3], "uploaded_at": r[4], "notes": r[5]}
        for r in rows
    ]


def get_exit_document(doc_id: int):
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT file_data FROM exit_documents WHERE id = %s", (doc_id,))
    row = cur.fetchone()
    conn.close()
    return bytes(row[0]) if row else None


def delete_exit_document(doc_id: int) -> None:
    conn = _conn()
    conn.execute("DELETE FROM exit_documents WHERE id = %s", (doc_id,))
    conn.commit()
    conn.close()


# ── Exit comps DB helpers ──────────────────────────────────────────────────────
COMPS_DB = os.path.join(_HERE, "data", "quona_exit_comps.db")
_COMP_NAME_MAP = {"VertoFX": "Verto FX"}  # benchmarking.db name → portfolio_comp_mapping name

def _comps_conn():
    return _get_pg_conn()

@st.cache_data(ttl=300)
def load_comp_mapping(company_name: str) -> pd.DataFrame:
    name = _COMP_NAME_MAP.get(company_name, company_name)
    return pd.read_sql_query(
        "SELECT comp_id, relevance_score, mapping_rationale "
        "FROM portfolio_comp_mapping WHERE portfolio_company = %s "
        "ORDER BY relevance_score DESC",
        _comps_conn(), params=(name,),
    )

@st.cache_data(ttl=300)
def load_comps_detail(comp_ids: tuple) -> pd.DataFrame:
    if not comp_ids:
        return pd.DataFrame()
    ph = ",".join(["%s"] * len(comp_ids))
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
    ph = ",".join(["%s"] * len(comp_ids))
    return pd.read_sql_query(f"""
        SELECT comp_id, company_name, stage, revenue_range_usd_m,
               revenue_growth_pct, gross_margin_pct, ebitda_margin_pct
        FROM comp_stage_snapshots WHERE comp_id IN ({ph})
    """, _comps_conn(), params=list(comp_ids))


def _discover_comps_via_claude(
    company_name: str,
    sub_sector: str,
    geography: str,
    founded_year,
    latest_revenue_usd,
    existing_comps: list[dict],
) -> list[dict]:
    """Use Claude + web search to find recent exits/listings that are genuine
    comparables. Anchored on the company's existing curated comps as exemplars.
    Undisclosed financials left None — never estimated."""
    import json as _json
    import anthropic as _anthropic

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not set in secrets.toml")

    rev_str = f"~${latest_revenue_usd/1e6:.1f}M" if latest_revenue_usd else "undisclosed"
    exemplars = "; ".join(
        f"{c.get('company_name')} ({c.get('sub_sector')}, {c.get('geography')}, "
        f"{c.get('exit_type')})" for c in existing_comps[:12]
    ) or "none on record"

    prompt = (
        f"Find recent M&A exits and IPOs/listings that are GENUINE comparables "
        f"for {company_name}: a {sub_sector} company in {geography}, founded "
        f"{founded_year}, current revenue {rev_str}.\n\n"
        f"These are the comps we already consider strong for this company — "
        f"treat them as the gold standard of what a relevant comp looks like, "
        f"and find more deals that resemble them in business model, stage, and "
        f"market:\n{exemplars}\n\n"
        f"Use web search to verify each deal against a credible source (company "
        f"press release, regulator filing, reputable financial press). "
        f"Prioritise the last 3 years.\n\n"
        f"APPLY THIS RELEVANCE TEST to every candidate and EXCLUDE any that fail:\n"
        f"- Sub-sector must match at the SPECIFIC level, not the broad level. A "
        f"payments infrastructure company is not a comp for a consumer wallet. "
        f"A B2B cross-border FX business is not a comp for a domestic card "
        f"acquirer, even though all are 'payments'.\n"
        f"- Stage and scale must be broadly comparable. Exclude deals where the "
        f"company was an order of magnitude larger or smaller than {company_name} "
        f"is now, unless it is a clear directional precedent.\n"
        f"- Geography: prefer {geography}, but DELIBERATELY include relevant "
        f"deals from other emerging markets (LatAm, SE Asia, MENA) where the "
        f"model is genuinely analogous, for regional diversity. EXCLUDE "
        f"developed-market mega-deals (US/EU large caps) unless a genuine "
        f"strategic precedent for this specific company.\n\n"
        f"It is better to return 3 highly relevant, well-sourced comps than 15 "
        f"loosely related ones. Quality over coverage. Return [] rather than "
        f"padding with marginal deals.\n\n"
        f"Return ONLY a JSON array. Each object uses these exact keys:\n"
        f"company_name, sub_sector, geography, exit_status, exit_year, exit_type "
        f"(IPO | Strategic M&A | Financial Sponsor | Secondary), exit_ev_usd_m, "
        f"revenue_at_exit_usd_m, gross_margin_pct, ebitda_margin_pct, "
        f"ev_revenue_multiple, revenue_growth_at_exit, key_narrative_drivers, "
        f"relevance_score (0-100 for {company_name}), mapping_rationale (one "
        f"line stating WHY it passed the relevance test — if you cannot "
        f"articulate a specific reason, do not return the comp), data_confidence "
        f"(High/Medium/Low), source_url, low_confidence_fields (array of field "
        f"names you are unsure about).\n\n"
        f"CRITICAL RULES:\n"
        f"- If a financial figure is not publicly disclosed, set it to null. "
        f"Never estimate or infer. Null is correct and expected for private deals.\n"
        f"- Every non-null financial figure must be supported by source_url.\n"
        f"- Only include deals verified via web search. Skip rumoured deals.\n"
        f"- *_usd_m in millions USD. *_pct as numbers (45.4 not 0.454).\n"
        f"Return only the JSON array, no other text."
    )

    client = _anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
    )

    text = "".join(
        b.text for b in msg.content if getattr(b, "type", "") == "text"
    ).strip()
    text = text.replace("```json", "").replace("```", "").strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1:
        return []
    rows = _json.loads(text[start:end + 1])

    for r in rows:
        for k, v in list(r.items()):
            if v in ("", None):
                r[k] = None
            elif k == "exit_year":
                try: r[k] = int(float(v))
                except Exception: r[k] = None
            elif k.endswith(("_usd_m", "_pct", "_multiple", "_score")) or k == "revenue_growth_at_exit":
                try: r[k] = float(v)
                except Exception: r[k] = None
    rows.sort(key=lambda r: r.get("relevance_score") or 0, reverse=True)
    return rows


def _save_discovered_comps(portfolio_company_name: str, comps: list[dict]) -> int:
    """Write confirmed Claude-sourced comps to Postgres in a single transaction.
    Converts 0-100 relevance score to 1-5 for portfolio_comp_mapping.
    use_for_margins and use_for_multiples default to 0 so these comps appear
    in the set for reference but don't affect median calculations until verified."""

    def _to_1_5(score_100):
        if score_100 is None:
            return None
        s = float(score_100)
        if s >= 80: return 5
        if s >= 60: return 4
        if s >= 40: return 3
        if s >= 20: return 2
        return 1

    conn = _get_pg_conn()
    saved = 0
    try:
        cur = conn.cursor()
        for c in comps:
            cur.execute("""
                INSERT INTO exit_comps
                  (company_name, sub_sector, geography, exit_status, exit_year,
                   exit_type, exit_ev_usd_m, revenue_at_exit_usd_m,
                   gross_margin_pct, ebitda_margin_pct, ev_revenue_multiple,
                   revenue_growth_at_exit, key_narrative_drivers, data_confidence,
                   source_url, discovered_via, discovered_at,
                   is_clean_exit, use_for_margins, use_for_multiples)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
                        'claude_web_search', NOW(), 1, 0, 0)
                RETURNING comp_id
            """, (
                c.get("company_name"), c.get("sub_sector"), c.get("geography"),
                c.get("exit_status"), c.get("exit_year"), c.get("exit_type"),
                c.get("exit_ev_usd_m"), c.get("revenue_at_exit_usd_m"),
                c.get("gross_margin_pct"), c.get("ebitda_margin_pct"),
                c.get("ev_revenue_multiple"), c.get("revenue_growth_at_exit"),
                c.get("key_narrative_drivers"), c.get("data_confidence"),
                c.get("source_url"),
            ))
            new_comp_id = cur.fetchone()[0]
            cur.execute("""
                INSERT INTO portfolio_comp_mapping
                  (portfolio_company, comp_id, relevance_score, mapping_rationale)
                VALUES (%s,%s,%s,%s)
                ON CONFLICT DO NOTHING
            """, (
                _COMP_NAME_MAP.get(portfolio_company_name, portfolio_company_name),
                new_comp_id,
                _to_1_5(c.get("relevance_score")),
                c.get("mapping_rationale"),
            ))
            saved += 1
        conn.commit()
    except Exception:
        raise
    finally:
        conn.close()
    return saved


def load_companies(db_version: str = "") -> pd.DataFrame:
    _key = "_ws_companies"
    if _key in st.session_state:
        return st.session_state[_key]
    result = pd.read_sql_query("""
        SELECT c.id, c.name, c.sector, c.hq_country, c.founded_year, c.fund,
               k.revenue_usd,
               k.ebitda_usd,
               k.gross_margin_pct,
               COALESCE(
                   k.ebitda_margin_pct,
                   CASE WHEN k.revenue_usd > 0 AND k.ebitda_usd IS NOT NULL
                        THEN ROUND(CAST(k.ebitda_usd * 100.0 / k.revenue_usd AS NUMERIC), 2)
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
    st.session_state[_key] = result
    return result

def load_revenue_growth(db_version: str = "") -> pd.DataFrame:
    _key = "_ws_revenue_growth"
    if _key in st.session_state:
        return st.session_state[_key]
    result = pd.read_sql_query("""
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
                THEN ROUND(CAST((r1.revenue_usd - r2.revenue_usd) * 100.0 / r2.revenue_usd AS NUMERIC), 1)
            END AS revenue_growth_pct
        FROM ranked r1
        LEFT JOIN ranked r2
            ON r1.company_id = r2.company_id AND r2.rn = 2
        WHERE r1.rn = 1
    """, _conn())
    st.session_state[_key] = result
    return result

def load_ltm_revenue(db_version: str = "") -> pd.DataFrame:
    """
    LTM (last 12 months) or ARR-estimated revenue per company.
    - Monthly reporters: sum of last 12 monthly periods
    - Quarterly reporters: sum of last 4 quarterly periods
    - Annual reporters: last annual figure
    - If insufficient history: annualise available data, label 'ARR (est.)'
    """
    _key = "_ws_ltm_revenue"
    if _key in st.session_state:
        return st.session_state[_key]
    conn = _conn()

    # julianday() is SQLite-only; compute gap in Python after a plain window query.
    _p2 = pd.read_sql_query("""
        SELECT company_id AS id, period_end_date
        FROM (
            SELECT company_id, period_end_date,
                   ROW_NUMBER() OVER (
                       PARTITION BY company_id ORDER BY period_end_date DESC
                   ) AS rn
            FROM kpi_snapshots
        ) sub
        WHERE rn <= 2
    """, conn)
    _p2["period_end_date"] = pd.to_datetime(_p2["period_end_date"])
    _prows = []
    for _cid, _grp in _p2.groupby("id"):
        _dates = _grp.sort_values("period_end_date", ascending=False)["period_end_date"]
        _gap   = int((_dates.iloc[0] - _dates.iloc[1]).days) if len(_dates) >= 2 else 30
        _pt    = "monthly" if _gap <= 45 else ("quarterly" if _gap <= 135 else "annual")
        _n     = 12 if _gap <= 45 else (4 if _gap <= 135 else 1)
        _prows.append({"id": _cid, "gap_days": _gap, "period_type": _pt, "needed": _n})
    periods = pd.DataFrame(_prows) if _prows else pd.DataFrame(
        columns=["id", "gap_days", "period_type", "needed"])
    del _p2, _prows

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

    result = pd.DataFrame(results)
    st.session_state["_ws_ltm_revenue"] = result
    return result

def load_ltm_volume(db_version: str = "") -> pd.DataFrame:
    """Compute LTM TPV and GMV per company using Python instead of SQL window functions."""
    _key = "_ws_ltm_volume"
    if _key in st.session_state:
        return st.session_state[_key]
    query = """
        SELECT company_id, tpv_usd, gmv_usd, period_end_date
        FROM kpi_snapshots
        WHERE tpv_usd IS NOT NULL OR gmv_usd IS NOT NULL
        ORDER BY company_id, period_end_date DESC
    """
    df = pd.read_sql_query(query, _conn())

    results = []
    for company_id, group in df.groupby('company_id'):
        last_12 = group.head(12)
        results.append({
            'id': company_id,
            'ltm_tpv_usd': last_12['tpv_usd'].sum() if last_12['tpv_usd'].notna().any() else None,
            'ltm_gmv_usd': last_12['gmv_usd'].sum() if last_12['gmv_usd'].notna().any() else None,
        })

    result = pd.DataFrame(results) if results else pd.DataFrame(columns=['id', 'ltm_tpv_usd', 'ltm_gmv_usd'])
    st.session_state[_key] = result
    return result

def load_all_revenue(db_version: str = "") -> pd.DataFrame:
    _key = "_ws_all_revenue"
    if _key in st.session_state:
        return st.session_state[_key]
    result = pd.read_sql_query("""
        SELECT company_id AS id, period_end_date, revenue_usd
        FROM kpi_snapshots
        WHERE revenue_usd IS NOT NULL
        ORDER BY id, period_end_date
    """, _conn())
    st.session_state[_key] = result
    return result

def load_company_info(company_id: int, db_version: str = "") -> pd.Series:
    _key = f"_ws_company_info_{company_id}"
    if _key in st.session_state:
        return st.session_state[_key]
    df = pd.read_sql_query(
        "SELECT * FROM companies WHERE id = %s", _conn(), params=(company_id,)
    )
    result = df.iloc[0]
    st.session_state[_key] = result
    return result

def load_kpis(company_id: int, db_version: str = "") -> pd.DataFrame:
    _key = f"_ws_kpis_{company_id}"
    if _key in st.session_state:
        return st.session_state[_key]
    print(f"[load_kpis] DB={DB_PATH} company_id={company_id}")
    conn = _conn()
    try:
        df = pd.read_sql_query("""
            SELECT period_end_date,
                   revenue_usd, gross_profit_usd, gross_margin_pct,
                   ebitda_usd, ebitda_margin_pct,
                   net_income_usd, net_margin_pct,
                   revenue_growth_pct,
                   arr_usd, mrr_usd,
                   customer_count, active_clients_count,
                   net_revenue_retention_pct, cac_usd, ltv_usd,
                   loan_book_gross_usd, par_30_pct, par_90_pct,
                   npl_rate_pct, net_yield_pct, nim_pct,
                   aum_usd, gmv_usd, tpv_usd,
                   unique_borrowers_count
            FROM kpi_snapshots
            WHERE company_id = %s
            ORDER BY period_end_date
        """, conn, params=(company_id,))
    finally:
        conn.close()
    df["period_end_date"] = pd.to_datetime(df["period_end_date"])

    rev = df["revenue_usd"].replace(0, float("nan"))

    # Gross margin fallback: derive from gross_profit_usd if margin not stored
    mask = df["gross_margin_pct"].isna() & df["gross_profit_usd"].notna() & rev.notna()
    if mask.any():
        safe_rev = df.loc[mask, "revenue_usd"].replace(0, float("nan"))
        df.loc[mask, "gross_margin_pct"] = (
            df.loc[mask, "gross_profit_usd"].astype(float)
            .div(safe_rev.astype(float))
            .mul(100)
            .round(4)
        )

    # EBITDA margin fallback
    mask = df["ebitda_margin_pct"].isna() & df["ebitda_usd"].notna() & rev.notna()
    if mask.any():
        safe_rev = df.loc[mask, "revenue_usd"].replace(0, float("nan"))
        df.loc[mask, "ebitda_margin_pct"] = (
            df.loc[mask, "ebitda_usd"].astype(float)
            .div(safe_rev.astype(float))
            .mul(100)
            .round(4)
        )

    # Net margin fallback
    mask = df["net_margin_pct"].isna() & df["net_income_usd"].notna() & rev.notna()
    if mask.any():
        safe_rev = df.loc[mask, "revenue_usd"].replace(0, float("nan"))
        df.loc[mask, "net_margin_pct"] = (
            df.loc[mask, "net_income_usd"].astype(float)
            .div(safe_rev.astype(float))
            .mul(100)
            .round(4)
        )

    st.session_state[f"_ws_kpis_{company_id}"] = df
    return df


def _kpi_last_updated(company_id: int) -> str:
    """Read MAX(updated_at) directly from DB — not cached, always fresh."""
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT MAX(updated_at) FROM kpi_snapshots WHERE company_id=%s",
            (company_id,),
        )
        row = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    ts = row[0] if row and row[0] else None
    if not ts:
        return "unknown"
    try:
        if hasattr(ts, "strftime"):
            return ts.strftime("%d %b %Y %H:%M") + " UTC"
        dt = datetime.fromisoformat(str(ts))
        return dt.strftime("%d %b %Y %H:%M") + " UTC"
    except Exception:
        return str(ts)

def _kpi_db_version(company_id: int) -> str:
    """Raw MAX(updated_at) for one company — used as a @st.cache_data key."""
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT MAX(updated_at) FROM kpi_snapshots WHERE company_id=%s",
            (company_id,),
        )
        row = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    return str(row[0]) if (row and row[0]) else ""

def _db_global_version() -> str:
    return ""  # session state warm cache handles invalidation


@st.cache_data(ttl=300)
def _ipo_readiness_load(company_id: int) -> dict:
    """Return {item_key: {status, notes, updated_at}} from ipo_readiness table."""
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT item_key, status, notes, updated_at FROM ipo_readiness WHERE company_id=%s",
            (company_id,)
        )
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()
    return {r[0]: {"status": r[1], "notes": r[2] or "", "updated_at": r[3] or ""} for r in rows}


def _ipo_readiness_save(company_id: int, updates: dict) -> None:
    """Upsert {item_key: {status, notes}} rows into ipo_readiness."""
    conn = _conn()
    try:
        cur = conn.cursor()
        for item_key, data in updates.items():
            cur.execute(
                "INSERT INTO ipo_readiness (company_id, item_key, status, notes, updated_at) "
                "VALUES (%s, %s, %s, %s, NOW()) "
                "ON CONFLICT(company_id, item_key) DO UPDATE SET "
                "status=excluded.status, notes=excluded.notes, updated_at=excluded.updated_at",
                (company_id, item_key, data.get("status", "Not Started"), data.get("notes", ""))
            )
        cur.close()
        conn.commit()
    finally:
        conn.close()


def _warm_cache() -> None:
    """Pre-load all KPI tables into st.session_state once per session.

    Called on every render but exits immediately after the first successful
    warm. Subsequent calls are a single dict lookup — no DB round trips.
    To force a reload after a write, delete the _ws_* keys and set
    _cache_warmed = False before calling st.rerun().
    """
    if st.session_state.get("_cache_warmed"):
        return
    load_companies()
    load_revenue_growth()
    load_ltm_revenue()
    load_ltm_volume()
    load_all_revenue()
    companies_df = st.session_state.get("_ws_companies", pd.DataFrame())
    for cid in companies_df["id"].tolist():
        cid = int(cid)
        load_company_info(cid)
        load_kpis(cid)
    st.session_state["_cache_warmed"] = True


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
    if not s or str(s).lower() == "none":
        return ""
    return SECTOR_LABELS.get(s, s.replace("_", " ").title())

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

    # Force fresh comp data for Lulalend after is_clean_exit flags were corrected in DB
    if company_name == "Lulalend":
        load_comps_detail.clear()
        load_comp_mapping.clear()

    comp_mapping = load_comp_mapping(company_name)

    if comp_mapping.empty:
        if company_name == "Cowrywise":
            st.markdown(
                f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                f"padding:20px 24px;color:{MUTED};font-size:13px;line-height:1.8'>"
                f"No direct wealthtech exit comps exist at scale for Nigeria. "
                f"Benchmarking uses EM digital wealth management proxies. "
                f"Add comps via the Upload Data tab to improve accuracy."
                f"</div>",
                unsafe_allow_html=True,
            )
        elif company_name == "Khazna":
            # Hardcoded EWA comp medians — Payfare is only listed pure-play EWA
            _KH_GM_MED    = 26.0   # Payfare gross margin %
            _KH_EM_MED    = 15.0   # Payfare EBITDA margin %
            _KH_REV_SCALE = 235.0  # Payfare / DailyPay revenue at exit (USD M)
            _KH_N_COMPS   = 3      # Payfare, DailyPay, Wagestream

            st.markdown(
                f"<div style='background:{WARN_BG};border:1px solid {WARN};border-radius:8px;"
                f"padding:10px 14px;font-size:12px;color:{WARN};margin-bottom:16px'>"
                f"<b>Note:</b> Comp set is limited given Khazna's unique positioning as an Egypt/KSA digital "
                f"workforce bank. Payfare is the only publicly listed pure-play EWA comp. Comp medians are based "
                f"on {_KH_N_COMPS} reference points and should be treated as directional benchmarks only.</div>",
                unsafe_allow_html=True,
            )

            # ── Live data from DB ──────────────────────────────────────────────
            _kh_rev    = ltm_val
            _kh_gm     = ltm_gm_pct
            _kh_em     = ltm_em_pct
            _kh_ac     = None
            _kh_arr    = None
            _kh_rev_m  = _kh_rev / 1e6 if _kh_rev else None
            _kh_hist   = 0
            _kh_rev_lt = None
            if not kpis.empty:
                if "active_clients_count" in kpis.columns:
                    _ac_s = kpis["active_clients_count"].dropna()
                    if not _ac_s.empty:
                        _kh_ac = float(_ac_s.iloc[-1])
                if "arr_usd" in kpis.columns:
                    _arr_s = kpis["arr_usd"].dropna()
                    if not _arr_s.empty:
                        _kh_arr = float(_arr_s.iloc[-1])
                if "revenue_usd" in kpis.columns:
                    _rv_s = kpis["revenue_usd"].dropna()
                    if not _rv_s.empty:
                        _kh_rev_lt = float(_rv_s.iloc[-1])
                        _kh_hist   = len(_rv_s)

            # ── Section 1: Summary stat cards ─────────────────────────────────
            def _kh_arrow(co_val, med_val, suffix="pp"):
                if co_val is None or med_val is None:
                    return f"<div style='font-size:12px;color:{MUTED};margin-top:5px'>Portfolio: —</div>"
                delta = float(co_val) - float(med_val)
                arrow = "↑" if delta > 0 else ("↓" if delta < 0 else "→")
                clr   = "#2E7D32" if delta > 0 else ("#C62828" if delta < 0 else MUTED)
                sign  = "+" if delta > 0 else ""
                return (
                    f"<div style='font-size:12px;color:{clr};font-weight:600;margin-top:5px'>"
                    f"{arrow}&nbsp;{fmt_pct(co_val)}"
                    f"&nbsp;<span style='font-weight:400;color:{MUTED}'>({sign}{delta:.1f}{suffix} vs median)</span>"
                    f"</div>"
                )

            def _kh_card(label, value_str, sub_html=""):
                return (
                    f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                    f"padding:18px 20px'>"
                    f"<div style='font-size:10px;text-transform:uppercase;letter-spacing:.6px;"
                    f"color:{MUTED};font-weight:600;margin-bottom:6px'>{label}</div>"
                    f"<div style='font-size:24px;font-weight:700;color:{BLACK}'>{value_str}</div>"
                    f"{sub_html}</div>"
                )

            _pct_scale = (
                f"<div style='font-size:12px;color:#6A1B9A;font-weight:600;margin-top:5px'>"
                f"{(_kh_rev_m / _KH_REV_SCALE * 100):.0f}% of comp exit scale</div>"
                if _kh_rev_m else
                f"<div style='font-size:12px;color:{MUTED};margin-top:5px'>No data uploaded</div>"
            )

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(
                    _kh_card(
                        "LTM Revenue",
                        fmt_usd(_kh_rev) if _kh_rev else "—",
                        f"<div style='font-size:12px;color:{MUTED};margin-top:5px'>"
                        f"Comp exit scale: ${_KH_REV_SCALE:.0f}M</div>",
                    ),
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    _kh_card(f"Comp Exit Scale", f"${_KH_REV_SCALE:.0f}M", _pct_scale),
                    unsafe_allow_html=True,
                )
            with c3:
                st.markdown(
                    _kh_card(
                        "LTM Gross Margin",
                        fmt_pct(_kh_gm) if _kh_gm is not None else "—",
                        _kh_arrow(_kh_gm, _KH_GM_MED),
                    ),
                    unsafe_allow_html=True,
                )
            with c4:
                st.markdown(
                    _kh_card(
                        "LTM EBITDA Margin",
                        fmt_pct(_kh_em) if _kh_em is not None else "—",
                        _kh_arrow(_kh_em, _KH_EM_MED),
                    ),
                    unsafe_allow_html=True,
                )

            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

            # ── Section 2: Gap analysis (left) + Radar chart (right) ──────────
            _KH_STATUS = {
                "ahead":   ("#2E7D32", GREEN,     "#E8F5E9", "AHEAD"),
                "behind":  (WARN,      WARN_BG,   WARN_BG,   "BEHIND"),
                "no_data": (MUTED,     "#F5F5F5", "#F5F5F5", "NO DATA"),
                "scale":   ("#6A1B9A", "#F3E5F5", "#F3E5F5", "SCALE"),
            }

            def _kh_bar(label, co_val, med_val, co_str, med_str, delta_str, bar_pct, status, note="Comp median"):
                bc, _, bb, bt = _KH_STATUS[status]
                return (
                    f"<div style='border-left:4px solid {bc};background:{WHITE};"
                    f"border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:12px;"
                    f"box-shadow:0 1px 3px rgba(0,0,0,.04)'>"
                    f"<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>"
                    f"<span style='font-size:10px;text-transform:uppercase;letter-spacing:.5px;"
                    f"color:{MUTED};font-weight:600'>{label}</span>"
                    f"<span style='background:{bb};color:{bc};border-radius:4px;"
                    f"padding:2px 8px;font-size:10px;font-weight:700'>{bt}</span>"
                    f"</div>"
                    f"<div style='display:flex;align-items:baseline;gap:8px;margin-bottom:10px;flex-wrap:wrap'>"
                    f"<span style='font-size:20px;font-weight:700;color:{BLACK}'>{co_str}</span>"
                    f"<span style='font-size:12px;color:{MUTED}'>vs {med_str} median</span>"
                    f"<span style='font-size:12px;color:{bc};font-weight:600'>{delta_str}</span>"
                    f"</div>"
                    f"<div style='background:{BG};border-radius:4px;height:6px;overflow:hidden'>"
                    f"<div style='background:{bc};height:6px;width:{bar_pct:.0f}%;border-radius:4px'></div>"
                    f"</div>"
                    f"<div style='display:flex;justify-content:space-between;margin-top:3px'>"
                    f"<span style='font-size:10px;color:{MUTED}'>0</span>"
                    f"<span style='font-size:10px;color:{MUTED}'>{note}</span>"
                    f"</div></div>"
                )

            col_left, col_right = st.columns(2, gap="large")

            with col_left:
                st.markdown(
                    f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                    f"color:{MUTED};font-weight:600;margin-bottom:12px'>Performance vs. Comp Medians</div>",
                    unsafe_allow_html=True,
                )

                # Gross Margin bar
                if _kh_gm is not None:
                    _ref_gm = max(abs(_KH_GM_MED), abs(_kh_gm), 1.0)
                    _gm_bar = min(max((_kh_gm + _ref_gm) / (_ref_gm * 2) * 100, 0), 100)
                    _gm_delta = _kh_gm - _KH_GM_MED
                    _gm_ds    = (f"+{_gm_delta:.1f}pp" if _gm_delta >= 0 else f"{_gm_delta:.1f}pp")
                    _gm_stat  = "ahead" if _gm_delta >= 0 else "behind"
                else:
                    _gm_bar, _gm_ds, _gm_stat = 0.0, "—", "no_data"
                st.markdown(
                    _kh_bar("Gross Margin",
                            _kh_gm, _KH_GM_MED,
                            fmt_pct(_kh_gm) if _kh_gm is not None else "—",
                            fmt_pct(_KH_GM_MED),
                            _gm_ds, _gm_bar, _gm_stat,
                            "Payfare benchmark"),
                    unsafe_allow_html=True,
                )

                # EBITDA Margin bar
                if _kh_em is not None:
                    _ref_em = max(abs(_KH_EM_MED), abs(_kh_em), 1.0)
                    _em_bar = min(max((_kh_em + _ref_em) / (_ref_em * 2) * 100, 0), 100)
                    _em_delta = _kh_em - _KH_EM_MED
                    _em_ds    = (f"+{_em_delta:.1f}pp" if _em_delta >= 0 else f"{_em_delta:.1f}pp")
                    _em_stat  = "ahead" if _em_delta >= 0 else "behind"
                else:
                    _em_bar, _em_ds, _em_stat = 0.0, "—", "no_data"
                st.markdown(
                    _kh_bar("EBITDA Margin",
                            _kh_em, _KH_EM_MED,
                            fmt_pct(_kh_em) if _kh_em is not None else "—",
                            fmt_pct(_KH_EM_MED),
                            _em_ds, _em_bar, _em_stat,
                            "Payfare benchmark"),
                    unsafe_allow_html=True,
                )

                # Revenue vs exit scale bar
                if _kh_rev_m is not None:
                    _rv_pct  = min(_kh_rev_m / _KH_REV_SCALE * 100, 100)
                    _rv_ds   = f"{_rv_pct:.0f}% of comp exit scale"
                    _rv_stat = "scale"
                else:
                    _rv_pct, _rv_ds, _rv_stat = 0.0, "—", "no_data"
                st.markdown(
                    _kh_bar("Revenue vs Exit Scale",
                            _kh_rev_m, _KH_REV_SCALE,
                            f"${_kh_rev_m:.1f}M" if _kh_rev_m is not None else "—",
                            f"${_KH_REV_SCALE:.0f}M",
                            _rv_ds, _rv_pct, _rv_stat,
                            f"Comp median exit (${_KH_REV_SCALE:.0f}M)"),
                    unsafe_allow_html=True,
                )

            with col_right:
                # ── Radar chart — Khazna vs comp median ───────────────────────
                def _kh_norm(val, lo, hi):
                    if val is None:
                        return 0.0
                    return max(0.0, min(100.0, (float(val) - lo) / (hi - lo) * 100))

                _kh_gm_r  = _kh_norm(_kh_gm,  0,   80)
                _kh_em_r  = _kh_norm(_kh_em,  -20,  30)
                _kh_rev_r = min((_kh_rev_m / _KH_REV_SCALE * 100) if _kh_rev_m else 0.0, 100.0)
                _kh_ac_r  = (
                    min(math.log10(float(_kh_ac) + 1) / math.log10(7e6 + 1) * 100, 100.0)
                    if (_kh_ac and float(_kh_ac) > 0) else 0.0
                )
                _md_gm_r  = _kh_norm(_KH_GM_MED, 0,   80)
                _md_em_r  = _kh_norm(_KH_EM_MED, -20,  30)
                _md_rev_r = 100.0
                _md_ac_r  = 70.0

                _kh_cats = ["Gross Margin", "EBITDA Margin", "Revenue Scale", "Active Users"]
                _kh_co_v = [_kh_gm_r, _kh_em_r, _kh_rev_r, _kh_ac_r]
                _kh_md_v = [_md_gm_r, _md_em_r, _md_rev_r, _md_ac_r]

                fig_kh = go.Figure()
                fig_kh.add_trace(go.Scatterpolar(
                    r=_kh_co_v + [_kh_co_v[0]], theta=_kh_cats + [_kh_cats[0]],
                    fill="toself", fillcolor="rgba(213,250,148,0.30)",
                    line=dict(color=BLACK, width=2), name="Khazna",
                    hovertemplate="%{theta}: %{r:.0f}/100<extra></extra>",
                ))
                fig_kh.add_trace(go.Scatterpolar(
                    r=_kh_md_v + [_kh_md_v[0]], theta=_kh_cats + [_kh_cats[0]],
                    fill="toself", fillcolor="rgba(197,229,255,0.30)",
                    line=dict(color="#1565C0", width=2, dash="dot"), name="Comp Median",
                    hovertemplate="%{theta}: %{r:.0f}/100<extra></extra>",
                ))
                fig_kh.update_layout(
                    polar=dict(
                        bgcolor=WHITE,
                        radialaxis=dict(visible=False, range=[0, 100]),
                        angularaxis=dict(tickfont=dict(size=11, color=BLACK), linecolor=BORDER, gridcolor=BORDER),
                    ),
                    showlegend=True,
                    legend=dict(
                        font=dict(size=11, color=BLACK), bgcolor=WHITE,
                        bordercolor=BORDER, borderwidth=1,
                        orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                    ),
                    paper_bgcolor=BG, margin=dict(l=30, r=30, t=20, b=50), height=380,
                )
                st.plotly_chart(fig_kh, use_container_width=True, config={"displayModeBar": False})

            # ── Section 3: Implied exit value (ARR multiple slider) ────────────
            _arr_base = _kh_arr if _kh_arr else _kh_rev
            _arr_lbl  = "ARR" if _kh_arr else "LTM Revenue"

            if _arr_base is not None:
                _scale_pct   = min((_kh_rev_m / _KH_REV_SCALE * 100) if _kh_rev_m else 0.0, 100.0)
                _marker_left = max(5.0, min(_scale_pct, 95.0))

                _kh_multiple = st.slider(
                    "ARR Multiple (10–15x)",
                    min_value=10.0, max_value=15.0, value=12.0, step=0.5,
                    key="khazna_arr_slider", format="%.1fx",
                )
                _implied_ev = _arr_base * _kh_multiple

                st.markdown(
                    f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                    f"padding:24px 28px;margin-bottom:20px'>"
                    f"<div style='display:flex;justify-content:space-between;align-items:flex-start;"
                    f"flex-wrap:wrap;gap:8px;margin-bottom:14px'>"
                    f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                    f"color:{MUTED};font-weight:600'>Implied Exit Value</div>"
                    f"<div style='background:{BG};border-radius:4px;padding:3px 10px;"
                    f"font-size:11px;font-weight:600;color:{BLACK}'>{_kh_multiple:.1f}x ARR</div>"
                    f"</div>"
                    f"<div style='display:flex;align-items:baseline;gap:14px;flex-wrap:wrap;margin-bottom:20px'>"
                    f"<span style='font-size:36px;font-weight:700;color:{BLACK}'>{fmt_usd(_implied_ev)}</span>"
                    f"<span style='font-size:13px;color:{MUTED}'>{_arr_lbl}: {fmt_usd(_arr_base)}</span>"
                    f"</div>"
                    f"<div style='position:relative;height:14px;margin-bottom:6px'>"
                    f"<div style='background:linear-gradient(to right,{BG} 0%,{GREEN} 100%);"
                    f"border-radius:7px;height:14px;width:100%'></div>"
                    f"<div style='position:absolute;top:50%;left:{_marker_left:.1f}%;"
                    f"transform:translate(-50%,-50%)'>"
                    f"<div style='width:20px;height:20px;border-radius:50%;background:{BLACK};"
                    f"border:3px solid {WHITE};box-shadow:0 0 0 2px {BLACK}'></div>"
                    f"</div></div>"
                    f"<div style='display:flex;justify-content:space-between;margin-bottom:14px'>"
                    f"<span style='font-size:11px;color:{MUTED}'>$0</span>"
                    f"<span style='font-size:11px;color:{MUTED}'>Comp median exit (${_KH_REV_SCALE:.0f}M)</span>"
                    f"</div>"
                    f"<div style='font-size:12px;color:{MUTED};border-top:1px solid {BORDER};padding-top:12px'>"
                    f"<b style='color:{BLACK}'>Methodology:</b> 10–15x ARR multiple (EWA / digital workforce banking benchmark). "
                    f"Payfare acquired at ~0.6x revenue (depressed public market); DailyPay ~7x; MNT-Halan ~3x. "
                    f"ARR multiple applied given Khazna's recurring lending revenue model. "
                    f"Khazna is at <b style='color:{BLACK}'>{_scale_pct:.0f}%</b> of comp median exit revenue scale."
                    f"</div></div>",
                    unsafe_allow_html=True,
                )

            # ── Section 4: Stage progression tracker ──────────────────────────
            if _kh_rev is not None:
                _ltm_m = _kh_rev / 1e6
                _KH_STAGES = [
                    ("Early Stage", "< $5M revenue"),
                    ("Growth",      "$6M – $30M revenue"),
                    ("Pre-Exit",    "$31M – $100M revenue"),
                    ("Exit Ready",  "> $100M + EBITDA positive"),
                ]
                if _ltm_m < 5:
                    _kh_stage = 0
                elif _ltm_m <= 30:
                    _kh_stage = 1
                elif _ltm_m <= 100:
                    _kh_stage = 2
                elif _kh_em is not None and _kh_em > 0:
                    _kh_stage = 3
                else:
                    _kh_stage = 2

                st.markdown(
                    f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                    f"color:{MUTED};font-weight:600;margin-bottom:14px'>Stage Progression</div>",
                    unsafe_allow_html=True,
                )
                _nodes_html = ""
                for _si, (_sname, _sublbl) in enumerate(_KH_STAGES):
                    _is_cur   = (_si == _kh_stage)
                    _dot_bg   = GREEN  if _is_cur else WHITE
                    _dot_bdr  = BLACK  if _is_cur else BORDER
                    _dot_sz   = "20px" if _is_cur else "14px"
                    _lbl_fw   = "700"  if _is_cur else "400"
                    _lbl_col  = BLACK  if _is_cur else MUTED
                    _ll_bg    = BORDER if _si > 0 else "transparent"
                    _rl_bg    = BORDER if _si < 3 else "transparent"
                    _dot_shad = f"0 0 0 3px {GREEN}" if _is_cur else "none"
                    _badge    = (
                        f"<div style='background:{GREEN};color:{BLACK};border-radius:4px;"
                        f"padding:1px 8px;font-size:10px;font-weight:700;margin-bottom:5px;"
                        f"display:inline-block;white-space:nowrap'>Khazna</div><br>"
                        if _is_cur else "<br>"
                    )
                    _nodes_html += (
                        f"<div style='flex:1;text-align:center;padding:0 8px'>"
                        f"{_badge}"
                        f"<div style='font-size:13px;font-weight:{_lbl_fw};color:{_lbl_col};"
                        f"margin-bottom:10px'>{_sname}</div>"
                        f"<div style='display:flex;align-items:center;justify-content:center;"
                        f"margin-bottom:10px'>"
                        f"<div style='height:2px;flex:1;background:{_ll_bg}'></div>"
                        f"<div style='width:{_dot_sz};height:{_dot_sz};border-radius:50%;"
                        f"background:{_dot_bg};border:2px solid {_dot_bdr};flex-shrink:0;"
                        f"box-shadow:{_dot_shad}'></div>"
                        f"<div style='height:2px;flex:1;background:{_rl_bg}'></div>"
                        f"</div>"
                        f"<div style='font-size:10px;color:{MUTED};text-align:center'>{_sublbl}</div>"
                        f"</div>"
                    )
                st.markdown(
                    f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                    f"padding:24px 20px;display:flex;align-items:flex-start;margin-bottom:20px'>"
                    f"{_nodes_html}</div>",
                    unsafe_allow_html=True,
                )

            # ── Comp set reference table ───────────────────────────────────────
            st.markdown(
                f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
                f"color:{MUTED};font-weight:600;margin-bottom:12px'>Comp Set — EWA and Digital Workforce Banking</div>",
                unsafe_allow_html=True,
            )
            _KH_CONF_DOT  = {"high": "#2E7D32", "medium": "#F57C00", "low": "#C62828"}
            _KH_REL_COLORS = {
                5: (GREEN,     BLACK),
                4: ("#D9F0D9", "#2E7D32"),
                3: ("#FFF9C4", "#795548"),
                2: ("#EFEBE9", MUTED),
                1: ("#F5F5F5", MUTED),
            }
            _KH_ET_COLORS = {
                "acquisition":       ("#C5E5FF", "#1565C0"),
                "financial sponsor": ("#D4D5CE", "#2C2C2A"),
                "ipo":               ("#D5FA94", "#2C2C2A"),
            }
            # (co, sub_sector, geo, et_label, et_key, is_clean, year, rev, gm, em, ev_rev, rel, conf)
            _KH_ROWS = [
                ("Payfare",    "EWA / Gig Banking",       "North America", "Financial Sponsor", "financial sponsor", True,  "2024", "$147M",       "26.0%", "15.0%", "0.6x",        4, "high"),
                ("DailyPay",   "EWA / Workforce Finance", "North America", "Pre-Exit (Pre-IPO)","pre-exit",          False, "—",    "—",           "—",     "—",     "~7x rev est.", 3, "medium"),
                ("Wagestream", "EWA / Workforce Finance", "Europe",        "Investor (Stake)",  "investor",          True,  "—",    "~$300M val.", "—",     "—",     "—",            5, "medium"),
            ]
            _kh_hdr_html = "".join(
                f"<th style='padding:8px 12px;text-align:left;font-size:10px;text-transform:uppercase;"
                f"letter-spacing:.5px;color:{MUTED};border-bottom:2px solid {BORDER};white-space:nowrap;"
                f"width:{w}'>{h}</th>"
                for h, w in [
                    ("Company", "18%"), ("Sub-sector", "14%"), ("Geography", "8%"),
                    ("Exit Type", "10%"), ("Year", "5%"),
                    ("Rev at Exit", "8%"), ("Gross Margin", "8%"), ("EBITDA Margin", "8%"),
                    ("EV/Rev", "8%"), ("Relevance", "7%"), ("Conf.", "6%"),
                ]
            )
            _kh_rows_html = ""
            for _idx, (_co, _sub, _geo, _et_label, _et_key, _is_clean, _yr, _rev, _gm, _em, _ev, _rel, _conf) in enumerate(_KH_ROWS):
                _row_bg   = WHITE if _idx % 2 == 0 else "#F9FAF7"
                _cell     = f"padding:8px 12px;font-size:13px;color:{MUTED}"
                _bg_r, _fg_r = _KH_REL_COLORS.get(_rel, ("#F5F5F5", MUTED))
                _conf_c   = _KH_CONF_DOT.get(_conf, MUTED)
                if not _is_clean:
                    _et_html = (
                        f"<span style='font-size:11px;color:{MUTED};font-style:italic'>"
                        f"Pre-exit / funding mark</span>"
                    )
                else:
                    _et_bg, _et_fg = _KH_ET_COLORS.get(_et_key, ("#D4D5CE", "#2C2C2A"))
                    _et_html = (
                        f"<span style='background:{_et_bg};color:{_et_fg};border-radius:4px;"
                        f"padding:2px 7px;font-size:11px;font-weight:600;white-space:normal'>{_et_label}</span>"
                    )
                _kh_rows_html += (
                    f"<tr style='background:{_row_bg}'>"
                    f"<td style='padding:8px 12px;font-size:13px;font-weight:700;color:{BLACK}'>{_co}</td>"
                    f"<td style='{_cell}'>{_sub}</td>"
                    f"<td style='{_cell}'>{_geo}</td>"
                    f"<td style='padding:8px 12px'>{_et_html}</td>"
                    f"<td style='{_cell}'>{_yr}</td>"
                    f"<td style='padding:8px 12px;font-size:13px;color:{BLACK};font-weight:500'>{_rev}</td>"
                    f"<td style='{_cell}'>{_gm}</td>"
                    f"<td style='{_cell}'>{_em}</td>"
                    f"<td style='padding:8px 12px;font-size:13px;color:{BLACK};font-weight:500'>{_ev}</td>"
                    f"<td style='padding:8px 12px'>"
                    f"<span style='background:{_bg_r};color:{_fg_r};border-radius:4px;"
                    f"padding:2px 8px;font-size:11px;font-weight:600'>{_rel}/5</span></td>"
                    f"<td style='padding:8px 12px;text-align:center'>"
                    f"<span style='display:inline-block;width:10px;height:10px;"
                    f"border-radius:50%;background:{_conf_c}'></span></td>"
                    f"</tr>"
                )
            st.markdown(
                f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                f"overflow:auto;margin-bottom:8px'>"
                f"<table style='width:100%;border-collapse:collapse'>"
                f"<thead><tr style='background:{BG}'>{_kh_hdr_html}</tr></thead>"
                f"<tbody>{_kh_rows_html}</tbody>"
                f"</table></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div style='font-size:11px;color:{MUTED};margin-top:4px;line-height:1.6'>"
                f"Gross margin and EBITDA margin benchmarks based on Payfare (the only publicly listed pure-play EWA comp). "
                f"Khazna's lending margin profile differs from pure EWA — PAR90 (0.4%) and loan book quality are the key "
                f"credit metrics to track alongside revenue growth.</div>",
                unsafe_allow_html=True,
            )
        elif company_name in ("TWINCO", "Twinco"):
            _TWINCO_COMPS = [
                ("Demica",    "Supply Chain Finance Platform", "Acquired by FIS",   "Dec 2024", "$300M",           "$40B AuA — 40% CAGR platform assets"),
                ("Taulia",    "Working Capital / SCF",         "Acquired by SAP",   "Mar 2022", "~$400M",          "$24M ARR at exit · ~17x ARR · $500B+ processed annually"),
                ("C2FO",      "Dynamic Discounting / SCF",     "Private",           "—",        "$1B (2019 val.)", "$186M ARR (2025) · ~5x ARR at last valuation"),
                ("Greensill", "Supply Chain Finance",           "Collapsed",         "Mar 2021", "$1.7B peak",      "Cautionary — fraud and concentration risk"),
                ("Stenn",     "Invoice Finance",                "Administration",    "Dec 2024", "$900M peak",      "Cautionary — HSBC fraud allegations"),
            ]
            _hdr_style = (
                "font-size:10px;font-weight:700;color:#93A3A1;text-transform:uppercase;"
                "letter-spacing:.5px;padding:8px 12px"
            )
            _cols_w = "1fr 1.2fr 1fr 0.7fr 1fr 2fr"
            _hdrs   = ["Company", "Type", "Status", "Date", "Valuation", "Key Metrics / Notes"]
            header_html = (
                f"<div style='display:grid;grid-template-columns:{_cols_w};"
                f"border-bottom:1px solid {BORDER};margin-bottom:4px'>"
                + "".join(f"<div style='{_hdr_style}'>{h}</div>" for h in _hdrs)
                + "</div>"
            )
            rows_html = ""
            for i, (co, typ, status, dt, val, notes) in enumerate(_TWINCO_COMPS):
                bg = "#F7F8F5" if i % 2 == 0 else "#FFFFFF"
                _cautionary = status in ("Collapsed", "Administration")
                val_color   = "#C62828" if _cautionary else BLACK
                _cell = (
                    f"font-size:12px;color:{MUTED};padding:8px 12px;line-height:1.5"
                )
                rows_html += (
                    f"<div style='display:grid;grid-template-columns:{_cols_w};"
                    f"background:{bg};border-radius:4px'>"
                    f"<div style='font-size:13px;font-weight:700;color:{BLACK};padding:8px 12px'>{co}</div>"
                    f"<div style='{_cell}'>{typ}</div>"
                    f"<div style='font-size:12px;color:{val_color};font-weight:600;padding:8px 12px'>{status}</div>"
                    f"<div style='{_cell}'>{dt}</div>"
                    f"<div style='font-size:13px;color:{BLACK};font-weight:600;padding:8px 12px'>{val}</div>"
                    f"<div style='{_cell}'>{notes}</div>"
                    f"</div>"
                )
            st.markdown(
                f"<div style='font-size:11px;color:{MUTED};margin-bottom:10px;line-height:1.6'>"
                f"Exit comp set — supply chain and trade finance platforms globally. "
                f"Revenue multiples not directly comparable given Twinco's unique PO finance model; "
                f"AuA-based and ARR-based valuation approaches both relevant.</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
                f"padding:16px 4px;overflow:hidden'>"
                + header_html + rows_html
                + "</div>",
                unsafe_allow_html=True,
            )
        else:
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

        MULTI_MARKET_COMPANIES = {
            "Verto", "VertoFX", "Enza", "TWINCO", "MaxSoko", "Khazna", "POWER",
        }
        geo_sub   = "Multi-Market" if company_name in MULTI_MARKET_COMPANIES else "Single Market"
        geo_label = f"Geography\n{geo_sub}"

        cats  = ["Gross Margin", "Revenue Scale", "EBITDA", geo_label, "Growth", "Customers"]
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
            margin=dict(l=50, r=50, t=20, b=60),
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
            "Remain Independent — Quona Pursues Secondaries",
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
        EXIT_TYPE_LABELS = {
            "acquisition":        "Acquisition",
            "strategic_ma":       "Strategic M&A",
            "strategic m&a":      "Strategic M&A",
            "ipo":                "IPO",
            "financial_sponsor":  "Financial Sponsor",
            "financial sponsor":  "Financial Sponsor",
            "secondary":          "Secondary",
            "private":            "Private",
            "collapsed":          "Collapsed",
            "administration":     "Administration",
        }
        exit_type_raw     = str(row.get("exit_type") or "").strip()
        et_key            = exit_type_raw.lower()
        exit_type_display = EXIT_TYPE_LABELS.get(et_key, exit_type_raw.replace("_", " ").title())
        if not is_clean:
            et_html = (
                f"<span style='font-size:11px;color:{MUTED};font-style:italic'>"
                f"Pre-exit / funding mark</span>"
            )
        else:
            et_bg, et_fg = EXIT_TYPE_COLORS.get(et_key, ("#D4D5CE", "#2C2C2A"))
            et_html = (
                f"<span style='background:{et_bg};color:{et_fg};border-radius:4px;"
                f"padding:2px 7px;font-size:11px;font-weight:600'>{exit_type_display}</span>"
                if exit_type_raw else "—"
            )

        exit_year_raw = row.get("exit_year")
        year_html     = str(int(exit_year_raw)) if not _is_null(exit_year_raw) else "—"

        rows_html += (
            f"<tr style='background:{row_bg};opacity:{'0.7' if not is_clean else '1'}'>"
            f"<td style='padding:8px 12px;font-weight:600;color:{BLACK};width:18%'>{name_html}</td>"
            f"<td style='padding:8px 12px;font-size:12px;color:{MUTED};width:14%'>"
            f"{(row.get('sub_sector') or '—').replace('_',' ').title()}</td>"
            f"<td style='padding:8px 12px;font-size:12px;color:{MUTED};width:8%'>{row.get('geography') or '—'}</td>"
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
    # Build dynamic footnote based on what's actually in this comp set
    pre_exit_names  = [r["company_name"] for _, r in comps.iterrows() if not int(r.get("is_clean_exit", 1))]
    no_margin_names = [r["company_name"] for _, r in comps.iterrows()
                       if int(r.get("is_clean_exit", 1)) and
                       _is_null(r.get("gross_margin_pct")) and _is_null(r.get("ebitda_margin_pct"))]
    estimated_names = [r["company_name"] for _, r in comps.iterrows()
                       if str(r.get("data_confidence", "")).lower() in ("low", "medium") and
                       int(r.get("is_clean_exit", 1))]

    footnote_parts = []
    if pre_exit_names:
        footnote_parts.append(
            f"Median calculations exclude {', '.join(pre_exit_names)} "
            f"({'pre-exit funding mark' if len(pre_exit_names) == 1 else 'pre-exit funding marks'}) "
            f"as these are not completed exits."
        )
    if no_margin_names:
        footnote_parts.append(
            f"{', '.join(no_margin_names)} margins not publicly disclosed and excluded from margin medians."
        )
    if estimated_names:
        footnote_parts.append(
            f"{', '.join(estimated_names)} figures are estimates based on available public data."
        )

    if footnote_parts:
        st.markdown(
            f"<div style='font-size:11px;color:{MUTED};font-style:italic;margin-bottom:20px'>"
            f"{' '.join(footnote_parts)}</div>",
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
                                    _conn_comps = _comps_conn()
                                    now_iso = datetime.utcnow().isoformat()
                                    _conn_comps.execute(
                                        """INSERT INTO exit_comps
                                           (company_name, exit_type, revenue_at_exit_usd_m,
                                            ev_revenue_multiple, data_source, created_at, updated_at)
                                           VALUES (%s,%s,%s,%s,%s,%s,%s)""",
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

    # ── Claude comp discovery ──────────────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin:28px 0 6px'>"
        f"Discover Comps with Claude (Web Search)</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:14px;line-height:1.6'>"
        f"Claude surfaces and sources recent deals. Private deal financials are often "
        f"undisclosed and will show as blank. Claude-sourced comps are added for "
        f"reference and excluded from benchmark medians until you verify and enable "
        f"them.</div>",
        unsafe_allow_html=True,
    )

    _sug_key = f"comp_suggestions_{company_id}"

    if st.button("Search for recent comps", key=f"discover_comps_{company_id}"):
        _c2 = _conn()
        _ss_row = _c2.execute(
            "SELECT sub_sector FROM companies WHERE id = %s", (company_id,)
        ).fetchone()
        _c2.close()
        _sub_sector = (_ss_row[0] if _ss_row else None) or info.get("sub_sector") or ""
        try:
            with st.spinner("Searching the web for recent exits and listings…"):
                _found = _discover_comps_via_claude(
                    company_name=company_name,
                    sub_sector=_sub_sector,
                    geography=info.get("hq_country") or "",
                    founded_year=info.get("founded_year"),
                    latest_revenue_usd=ltm_val,
                    existing_comps=comps.to_dict("records"),
                )
            st.session_state[_sug_key] = _found
        except Exception as _exc:
            st.error(f"Discovery failed: {_exc}")

    _suggestions = st.session_state.get(_sug_key) or []

    if _suggestions:
        def _fmt_fld(val, key, low_conf_keys):
            if val is None:
                return f"<span style='color:{MUTED}'>&mdash;</span>"
            color = WARN if key in low_conf_keys else BLACK
            if key.endswith("_pct"):
                return f"<span style='color:{color}'>{float(val):.1f}%</span>"
            if key.endswith("_usd_m"):
                return f"<span style='color:{color}'>${float(val):.0f}M</span>"
            if key.endswith("_multiple"):
                return f"<span style='color:{color}'>{float(val):.1f}x</span>"
            return f"<span style='color:{color}'>{val}</span>"

        st.markdown(
            f"<div style='background:{BG};border:2px dashed {BORDER};"
            f"border-radius:10px;padding:14px 18px 6px;margin-top:4px'>"
            f"<div style='font-size:10px;font-weight:700;text-transform:uppercase;"
            f"letter-spacing:.6px;color:{MUTED};margin-bottom:14px'>"
            f"Suggested — not yet in your comp set &nbsp;·&nbsp; "
            f"{len(_suggestions)} found, sorted by relevance</div></div>",
            unsafe_allow_html=True,
        )

        _selected_idxs: list[int] = []
        for _i, _sug in enumerate(_suggestions):
            _low_conf: list = _sug.get("low_confidence_fields") or []
            if isinstance(_low_conf, str):
                import json as _js
                try: _low_conf = _js.loads(_low_conf)
                except Exception: _low_conf = [_low_conf]
            _rel_score = float(_sug.get("relevance_score") or 0)
            _opacity   = "0.6" if _rel_score < 60 else "1"

            _c_sel, _c_info, _c_fin, _c_score = st.columns([0.5, 4.5, 3.5, 1.5])

            with _c_sel:
                _checked = st.checkbox(
                    "", key=f"cs_{company_id}_{_i}",
                    value=False, label_visibility="collapsed",
                )
                if _checked:
                    _selected_idxs.append(_i)

            with _c_info:
                _src  = _sug.get("source_url") or ""
                _co   = _sug.get("company_name") or "—"
                _et   = _sug.get("exit_type") or "—"
                _yr   = _sug.get("exit_year")
                _geo  = _sug.get("geography") or "—"
                _rat  = _sug.get("mapping_rationale") or ""
                _yr_s = str(int(_yr)) if _yr else "—"
                _name_html = (
                    f"<a href='{_src}' target='_blank' rel='noopener noreferrer' "
                    f"style='color:{BLACK};font-weight:700;text-decoration:underline;"
                    f"text-underline-offset:2px'>{_co}</a>"
                    if _src else
                    f"<span style='font-weight:700;color:{BLACK}'>{_co}</span>"
                )
                st.markdown(
                    f"<div style='opacity:{_opacity};padding-top:2px'>"
                    + _name_html
                    + f"<span style='color:{MUTED};font-size:11px;margin-left:8px'>"
                    + f"{_et} &nbsp;·&nbsp; {_yr_s} &nbsp;·&nbsp; {_geo}</span>"
                    + (
                        f"<div style='font-size:11px;color:{MUTED};margin-top:3px;"
                        f"line-height:1.4'>{_rat}</div>"
                        if _rat else ""
                    )
                    + "</div>",
                    unsafe_allow_html=True,
                )

            with _c_fin:
                _rv_h = _fmt_fld(_sug.get("revenue_at_exit_usd_m"), "revenue_at_exit_usd_m", _low_conf)
                _ev_h = _fmt_fld(_sug.get("exit_ev_usd_m"),         "exit_ev_usd_m",         _low_conf)
                _gm_h = _fmt_fld(_sug.get("gross_margin_pct"),      "gross_margin_pct",      _low_conf)
                _em_h = _fmt_fld(_sug.get("ebitda_margin_pct"),     "ebitda_margin_pct",     _low_conf)
                _mx_h = _fmt_fld(_sug.get("ev_revenue_multiple"),   "ev_revenue_multiple",   _low_conf)
                _conf = str(_sug.get("data_confidence") or "").capitalize()
                st.markdown(
                    f"<div style='font-size:12px;opacity:{_opacity};line-height:1.9'>"
                    f"Rev: {_rv_h} &nbsp;·&nbsp; EV: {_ev_h}<br>"
                    f"GM: {_gm_h} &nbsp;·&nbsp; EM: {_em_h} &nbsp;·&nbsp; EV/Rev: {_mx_h}"
                    + (f"<br><span style='font-size:10px;color:{MUTED}'>Confidence: {_conf}</span>" if _conf else "")
                    + "</div>",
                    unsafe_allow_html=True,
                )

            with _c_score:
                _sc_clr = "#2E7D32" if _rel_score >= 70 else (WARN if _rel_score >= 50 else "#C62828")
                st.markdown(
                    f"<div style='text-align:center;opacity:{_opacity};padding-top:4px'>"
                    f"<span style='font-size:16px;font-weight:700;color:{_sc_clr}'>"
                    f"{int(_rel_score)}</span>"
                    f"<span style='font-size:10px;color:{MUTED}'>/100</span></div>",
                    unsafe_allow_html=True,
                )

            if _i < len(_suggestions) - 1:
                st.markdown(
                    f"<div style='height:1px;background:{BORDER};margin:6px 0 8px'></div>",
                    unsafe_allow_html=True,
                )

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if st.button("Add selected to comp set", key=f"add_disc_{company_id}"):
            _to_save = [_suggestions[_j] for _j in _selected_idxs]
            if not _to_save:
                st.warning("Tick at least one comp to add.")
            else:
                try:
                    _n = _save_discovered_comps(company_name, _to_save)
                    load_comp_mapping.clear()
                    load_comps_detail.clear()
                    load_stage_snapshots.clear()
                    st.session_state.pop(_sug_key, None)
                    st.success(f"Added {_n} comp{'s' if _n != 1 else ''} — reloading…")
                    st.rerun()
                except Exception as _exc:
                    st.error(f"Save failed: {_exc}")


# ── DB write helpers ──────────────────────────────────────────────────────────

def _existing_periods(company_id: int) -> set[str]:
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT period_end_date FROM kpi_snapshots WHERE company_id = %s",
            (company_id,),
        )
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()
    return {r[0] for r in rows}


def _upsert_kpi(company_id: int, data: dict) -> None:
    conn   = _conn()
    now    = datetime.utcnow().isoformat()
    period = data["period_end_date"]
    print(f"[_upsert_kpi] DB={DB_PATH} company_id={company_id} period={period}")
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id FROM kpi_snapshots WHERE company_id=%s AND period_end_date=%s",
            (company_id, period),
        )
        existing = cur.fetchone()

        if existing:
            update_cols = {
                k: v for k, v in data.items()
                if k != "period_end_date" and v is not None
            }
            if update_cols:
                set_clause = ", ".join(f"{k}=%s" for k in update_cols)
                cur.execute(
                    f"UPDATE kpi_snapshots SET {set_clause}, updated_at=%s "
                    f"WHERE company_id=%s AND period_end_date=%s",
                    [*update_cols.values(), now, company_id, period],
                )
        else:
            row = {"company_id": company_id, "created_at": now, "updated_at": now,
                   **{k: v for k, v in data.items() if v is not None}}
            cols_str     = ", ".join(row.keys())
            placeholders = ", ".join(["%s"] * len(row))
            cur.execute(
                f"INSERT INTO kpi_snapshots ({cols_str}) VALUES ({placeholders})",
                list(row.values()),
            )
        cur.close()
        conn.commit()
    finally:
        conn.close()


def _recompute_growth(company_id: int) -> None:
    """Backfill revenue_growth_pct for every period of a company using the DB record order.

    Called after each upload batch so the first period in the batch (which the
    parser leaves as None) gets its growth filled from the prior DB period.
    """
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT id, revenue_usd FROM kpi_snapshots
               WHERE company_id = %s AND revenue_usd IS NOT NULL
               ORDER BY period_end_date""",
            (company_id,),
        )
        rows = cur.fetchall()
        for i, (row_id, rev) in enumerate(rows):
            if i == 0 or rows[i - 1][1] is None:
                growth = None
            else:
                prior = rows[i - 1][1]
                growth = round((rev - prior) / prior * 100, 4) if prior > 0 else None
            cur.execute(
                "UPDATE kpi_snapshots SET revenue_growth_pct = %s WHERE id = %s",
                (growth, row_id),
            )
        cur.close()
        conn.commit()
    finally:
        conn.close()


# ── Exit tracking DB helpers ──────────────────────────────────────────────────

def _exit_pathways_load(company_id: int) -> list[dict]:
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, pathway_name, likelihood, estimated_timeline, notes "
            "FROM exit_pathways WHERE company_id=%s ORDER BY id",
            (company_id,),
        )
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()
    return [{"id": r[0], "pathway_name": r[1], "likelihood": r[2],
             "estimated_timeline": r[3], "notes": r[4]} for r in rows]


def _exit_pathway_save(company_id: int, pid, name: str, likelihood: str,
                       timeline: str, notes: str) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    try:
        cur = conn.cursor()
        if pid:
            cur.execute(
                "UPDATE exit_pathways SET pathway_name=%s,likelihood=%s,"
                "estimated_timeline=%s,notes=%s,updated_at=%s WHERE id=%s",
                (name, likelihood, timeline, notes, now, pid),
            )
        else:
            cur.execute(
                "INSERT INTO exit_pathways "
                "(company_id,pathway_name,likelihood,estimated_timeline,notes,created_at,updated_at) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (company_id, name, likelihood, timeline, notes, now, now),
            )
        cur.close()
        conn.commit()
    finally:
        conn.close()


def _exit_pathway_delete(pid: int) -> None:
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM exit_pathways WHERE id=%s", (pid,))
        cur.close()
        conn.commit()
    finally:
        conn.close()


def _buyer_tracking_load(company_id: int) -> pd.DataFrame:
    return pd.read_sql_query(
        "SELECT id, acquirer_name, acquirer_type, relationship_owner, "
        "last_contact_date, status FROM buyer_tracking "
        "WHERE company_id=%s ORDER BY sort_order, id",
        _conn(), params=(company_id,),
    )


def _buyer_tracking_replace(company_id: int, df: pd.DataFrame) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM buyer_tracking WHERE company_id=%s", (company_id,))
        for i, row in df.iterrows():
            name = str(row.get("acquirer_name", "")).strip()
            if not name:
                continue
            cur.execute(
                "INSERT INTO buyer_tracking "
                "(company_id,acquirer_name,acquirer_type,relationship_owner,"
                "last_contact_date,status,sort_order,created_at,updated_at) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (company_id, name,
                 str(row.get("acquirer_type", "Strategic")),
                 str(row.get("relationship_owner", "") or ""),
                 str(row.get("last_contact_date", "") or ""),
                 str(row.get("status", "Not Started")),
                 i, now, now),
            )
        cur.close()
        conn.commit()
    finally:
        conn.close()


def _quarterly_actions_load(company_id: int, quarter: str) -> dict:
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT planned_actions, completed_actions, carry_forward "
            "FROM quarterly_actions WHERE company_id=%s AND quarter=%s",
            (company_id, quarter),
        )
        row = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    return {
        "planned_actions":   (row[0] or "") if row else "",
        "completed_actions": (row[1] or "") if row else "",
        "carry_forward":     (row[2] or "") if row else "",
    }


def _quarterly_actions_save(company_id: int, quarter: str,
                             planned: str, completed: str, carry: str) -> None:
    conn = _conn()
    now  = datetime.utcnow().isoformat()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id FROM quarterly_actions WHERE company_id=%s AND quarter=%s",
            (company_id, quarter),
        )
        exists = cur.fetchone()
        if exists:
            cur.execute(
                "UPDATE quarterly_actions SET planned_actions=%s,completed_actions=%s,"
                "carry_forward=%s,updated_at=%s WHERE company_id=%s AND quarter=%s",
                (planned, completed, carry, now, company_id, quarter),
            )
        else:
            cur.execute(
                "INSERT INTO quarterly_actions "
                "(company_id,quarter,planned_actions,completed_actions,carry_forward,created_at,updated_at) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (company_id, quarter, planned, completed, carry, now, now),
            )
        cur.close()
        conn.commit()
    finally:
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


def _box_get_latest_kpi_file(company_name: str) -> tuple[bytes, str, str] | tuple[None, None, None]:
    """
    Navigate Box folder structure to find the most recent KPI file for a company.

    Searches both the Management Accounts folder and Board Meetings folder,
    collects xlsx and pdf files, and returns the single most recently
    modified file across both sources.

    Returns (file_bytes, filename, file_type) where file_type is "xlsx" or "pdf",
    or (None, None, None) if not found.
    """
    import re as _re
    import io

    def _strip_prefix(name: str) -> str:
        return _re.sub(r'^\d+\s*', '', name).lower()

    def _score_folder(folder, keywords):
        n = folder.name.lower()
        return sum(1 for kw in keywords if kw in n)

    def _collect_files_from_source(source_folder):
        """Navigate year → month inside source_folder, collect xlsx/pdf files."""
        collected = []
        year_items = client.folders.get_folder_items(source_folder.id, limit=200).entries
        year_folders = []
        for f in year_items:
            if f.type.value != "folder":
                continue
            m = _re.search(r'20\d{2}', f.name)
            if m:
                year_folders.append((f, int(m.group())))

        st.write(f"  DEBUG source_folder={source_folder.name!r} year_folders={[f.name for f,_ in year_folders]}")

        if not year_folders:
            for f in year_items:
                if f.type.value != "file":
                    continue
                nl = f.name.lower()
                mod = getattr(f, "modified_at", None)
                if nl.endswith(".xlsx") or nl.endswith(".pdf"):
                    ft = "xlsx" if nl.endswith(".xlsx") else "pdf"
                    st.write(f"    DEBUG direct file: {f.name!r} modified_at={mod!r}")
                    collected.append({"id": f.id, "name": f.name,
                                      "modified_at": mod or "",
                                      "file_type": ft})
            return collected

        for year_folder, _ in sorted(year_folders, key=lambda x: x[1], reverse=True):
            month_items = client.folders.get_folder_items(year_folder.id, limit=200).entries
            month_folders = [i for i in month_items if i.type.value == "folder"]

            st.write(f"  DEBUG year={year_folder.name!r} month_folders={[f.name for f in month_folders]}")

            if not month_folders:
                candidates = month_items
                st.write(f"  DEBUG no month subfolders — using year folder directly")
            else:
                def _month_sort(f):
                    try:
                        return int(f.name.strip().split()[0])
                    except Exception:
                        return 0
                latest_month = sorted(month_folders, key=_month_sort, reverse=True)[0]
                st.write(f"  DEBUG selected month={latest_month.name!r}")
                candidates = client.folders.get_folder_items(latest_month.id, limit=200).entries

            for f in candidates:
                if f.type.value != "file":
                    continue
                nl = f.name.lower()
                mod = getattr(f, "modified_at", None)
                if nl.endswith(".xlsx") or nl.endswith(".pdf"):
                    ft = "xlsx" if nl.endswith(".xlsx") else "pdf"
                    st.write(f"    DEBUG candidate: {f.name!r} modified_at={mod!r} type={ft}")
                    collected.append({"id": f.id, "name": f.name,
                                      "modified_at": mod or "",
                                      "file_type": ft})
            if collected:
                break

        return collected

    try:
        from box_sdk_gen import BoxClient, BoxDeveloperTokenAuth

        token = st.secrets.get("BOX_DEV_TOKEN", "")
        if not token:
            return None, None, None

        auth   = BoxDeveloperTokenAuth(token=token)
        client = BoxClient(auth)

        # Fund folder mapping — portfolio companies folder ID for each company
        FUND_MAP = {
            "VertoFX":    "54633680859",    # 02 AQF Portfolio Companies (Fund II)
            "Lulalend":   "54633680859",    # 02 AQF Portfolio Companies (Fund II)
            "Cowrywise":  "54633680859",    # 02 AQF Portfolio Companies (Fund II)
            "MaxSoko":    "54633680859",    # 02 AQF Portfolio Companies (Fund II)
            "Yoco":       "3998199207",     # 01 AFIF Portfolio Companies (Fund I)
            "Khazna":     "136720620453",   # 04 Fund III Portfolio Companies
            "SAVA":       "136720620453",   # 04 Fund III Portfolio Companies
            "TWINCO":     "136720620453",   # 04 Fund III Portfolio Companies
            "Enza":       "136720620453",   # 04 Fund III Portfolio Companies
        }

        # Exact match first; fall back to case-insensitive partial match
        portfolio_folder_id = FUND_MAP.get(company_name)
        if not portfolio_folder_id:
            cn_lower = company_name.lower()
            for key, fid in FUND_MAP.items():
                if cn_lower in key.lower() or key.lower() in cn_lower:
                    portfolio_folder_id = fid
                    break
        if not portfolio_folder_id:
            return None, None, None

        # Step 1: Find company folder inside Portfolio Companies
        items = client.folders.get_folder_items(portfolio_folder_id, limit=200).entries
        cn_lower = company_name.lower()
        company_folder = None
        for i in items:
            if i.type.value != "folder":
                continue
            stripped = _strip_prefix(i.name)
            if cn_lower in stripped or stripped in cn_lower:
                company_folder = i
                break

        # Fallback: ask Claude to pick when partial matching fails
        if not company_folder:
            folder_candidates = [i for i in items if i.type.value == "folder"]
            try:
                import anthropic as _anthropic
                _lines = "\n".join(
                    f"{idx}: {f.name}" for idx, f in enumerate(folder_candidates[:20])
                )
                _ac = _anthropic.Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))
                _msg = _ac.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=5,
                    messages=[{
                        "role": "user",
                        "content": (
                            f"Which folder index best matches the company '{company_name}'? "
                            f"Reply with only the integer index.\n{_lines}"
                        ),
                    }],
                )
                _idx = int(_msg.content[0].text.strip())
                if 0 <= _idx < len(folder_candidates):
                    company_folder = folder_candidates[_idx]
            except Exception:
                pass
        if not company_folder:
            st.write(f"DEBUG company folder not found for {company_name!r}")
            return None, None, None

        st.write(f"DEBUG company_folder={company_folder.name!r} [{company_folder.id}]")

        # Step 2: Find both Management Accounts and Board Meetings folders
        sub_items = client.folders.get_folder_items(company_folder.id, limit=200).entries
        sub_folders = [i for i in sub_items if i.type.value == "folder"]

        MGMT_KEYWORDS  = [
            "management account", "financial statement", "monthly financial",
            "financials", "kpi", "reporting", "accounts", "monthly report",
        ]
        BOARD_KEYWORDS = ["board meeting", "board pack", "board deck", "board materials", "board"]

        source_folders = []
        for sf in sub_folders:
            ms = _score_folder(sf, MGMT_KEYWORDS)
            bs = _score_folder(sf, BOARD_KEYWORDS)
            if ms > 0 or bs > 0:
                source_folders.append(sf)
                st.write(f"DEBUG source_folder matched: {sf.name!r} mgmt_score={ms} board_score={bs}")
            else:
                st.write(f"DEBUG subfolder skipped: {sf.name!r}")

        # If nothing matched keywords, fall back to Claude picking the best single folder
        if not source_folders:
            try:
                import anthropic as _anthropic
                _lines = "\n".join(f"{i}: {f.name}" for i, f in enumerate(sub_folders[:10]))
                _ac = _anthropic.Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))
                _msg = _ac.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=5,
                    messages=[{
                        "role": "user",
                        "content": (
                            "Which folder index most likely contains monthly KPI or "
                            "financial reports for a portfolio company? "
                            f"Reply with only the integer index.\n{_lines}"
                        ),
                    }],
                )
                _idx = int(_msg.content[0].text.strip())
                if 0 <= _idx < len(sub_folders):
                    source_folders = [sub_folders[_idx]]
                    st.write(f"DEBUG Claude fallback selected: {sub_folders[_idx].name!r}")
            except Exception:
                pass
        if not source_folders:
            st.write("DEBUG no source folders found — aborting")
            return None, None, None

        # Step 3: Collect all xlsx/pdf candidates from every source folder
        all_candidates = []
        for sf in source_folders:
            all_candidates.extend(_collect_files_from_source(sf))

        st.write(f"DEBUG all_candidates ({len(all_candidates)}):")
        for c in all_candidates:
            st.write(f"  {c['name']!r}  modified_at={c['modified_at']!r}  type={c['file_type']}")

        if not all_candidates:
            return None, None, None

        # Step 4: Pick the single most recently modified file
        target = sorted(all_candidates, key=lambda x: x["modified_at"], reverse=True)[0]
        st.write(f"DEBUG selected target: {target['name']!r}  modified_at={target['modified_at']!r}")

        # Step 5: Download file content as bytes
        response = client.downloads.download_file(target["id"])
        buffer = io.BytesIO()
        if hasattr(response, 'read'):
            buffer.write(response.read())
        elif hasattr(response, 'iter_content'):
            for chunk in response.iter_content(chunk_size=8192):
                buffer.write(chunk)
        else:
            buffer.write(response)
        buffer.seek(0)
        return buffer.read(), target["name"], target["file_type"]

    except Exception as e:
        st.warning(f"Box sync error: {e}")
        return None, None, None


def _extract_kpis_from_pdf(pdf_bytes: bytes, company_name: str) -> tuple[list[dict], list[str]]:
    """
    Extract KPI rows from a PDF board pack or management report using the
    Claude API. Returns (rows, low_confidence_fields).

    Each row maps to kpi_snapshots columns. Fields not found are left as None
    — never guessed. low_confidence_fields lists "period|field" entries the
    model was unsure about, for amber flagging in the preview.
    """
    import base64, json as _json
    import anthropic as _anthropic

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not set")

    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    schema_fields = [
        "period_end_date", "reporting_currency", "fx_rate_to_usd",
        "revenue_usd", "gross_profit_usd", "gross_margin_pct",
        "ebitda_usd", "ebitda_margin_pct", "arr_usd", "mrr_usd",
        "customer_count", "active_clients_count", "net_revenue_retention_pct",
        "gross_churn_rate_pct", "cac_usd", "ltv_usd", "loan_book_gross_usd",
        "npl_rate_pct", "par_30_pct", "par_90_pct", "net_yield_pct",
        "cost_of_risk_pct", "nim_pct", "leverage_ratio", "aum_usd",
        "gmv_usd", "tpv_usd", "unique_borrowers_count",
        "top_3_concentration_pct", "insurance_policies_active",
        "devices_connected", "net_income_usd", "net_margin_pct",
    ]

    prompt = (
        f"You are extracting financial KPIs from a board pack or management "
        f"report for {company_name}, an African fintech company. "
        f"This data will feed an investor benchmarking platform, so accuracy "
        f"is critical.\n\n"
        f"Extract one JSON object per reporting period present in the document. "
        f"Use these exact field names (omit any you cannot find — do NOT guess "
        f"or estimate):\n{schema_fields}\n\n"
        f"RULES:\n"
        f"- period_end_date must be YYYY-MM-DD (last day of the period).\n"
        f"- All *_usd fields in absolute USD (convert if the report uses "
        f"another currency and a rate is given; if you convert, note it).\n"
        f"- All *_pct fields as numbers (e.g. 45.4 for 45.4%, not 0.454).\n"
        f"- If a figure is unclear, ambiguous, or you had to infer it, list it "
        f"in low_confidence as 'PERIOD|field' (e.g. '2026-05-31|gross_margin_pct').\n"
        f"- Never fabricate a number. A missing field is better than a wrong one.\n\n"
        f"Return ONLY this JSON, no other text:\n"
        f'{{"rows": [ {{...period objects...}} ], '
        f'"low_confidence": ["PERIOD|field", ...], '
        f'"source_pages": "brief note on which pages figures came from"}}'
    )

    client = _anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "document",
                 "source": {"type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    )

    raw = msg.content[0].text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    parsed = _json.loads(raw)

    rows = parsed.get("rows", [])
    low_conf = parsed.get("low_confidence", [])

    for r in rows:
        for k, v in list(r.items()):
            if v in ("", None):
                r[k] = None
            elif k.endswith("_count") or k in ("insurance_policies_active", "devices_connected"):
                try:
                    r[k] = int(float(v))
                except Exception:
                    r[k] = None
            elif k not in ("period_end_date", "reporting_currency"):
                try:
                    r[k] = float(v)
                except Exception:
                    r[k] = None

    return rows, low_conf


def render_upload_tab(info: pd.Series, company_id: int) -> None:
    company_name = info["name"]

    st.markdown(
        f"<div style='color:{MUTED};font-size:12px;margin-bottom:16px;line-height:1.7'>"
        f"Upload the latest Excel report for <b style='color:{BLACK}'>{company_name}</b>. "
        f"The parser will extract new periods automatically and show a preview "
        f"before writing anything to the database.</div>",
        unsafe_allow_html=True,
    )

    # ── Box sync button ───────────────────────────────────────────────────────
    box_token = st.secrets.get("BOX_DEV_TOKEN", "")
    if box_token:
        col_box, col_info, _ = st.columns([2, 4, 2])
        with col_box:
            if st.button("↓ Sync from Box", key=f"box_sync_{company_id}"):
                with st.spinner("Fetching latest KPI file from Box…"):
                    box_bytes, box_filename, box_filetype = _box_get_latest_kpi_file(company_name)
                if box_bytes:
                    st.session_state[f"box_file_bytes_{company_id}"]  = box_bytes
                    st.session_state[f"box_file_name_{company_id}"]   = box_filename
                    st.session_state[f"box_is_pdf_{company_id}"]      = (box_filetype == "pdf")
                    st.session_state.pop(f"box_pdf_low_conf_{company_id}", None)
                    st.rerun()
                else:
                    st.error(
                        "No KPI file found in Box for this company. "
                        "Check that the Management Accounts and Board Meetings "
                        "folders exist and contain xlsx or pdf files."
                    )
        with col_info:
            st.markdown(
                f"<div style='font-size:11px;color:#93A3A1;padding-top:8px'>"
                f"Pulls the most recent monthly KPI file from Box automatically.</div>",
                unsafe_allow_html=True,
            )
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Select Excel file",
        type=["xlsx"],
        key=f"uploader_{company_id}",
        label_visibility="collapsed",
    )

    # Box-injected file — treat identically to a manual upload.
    # Use .get() (not .pop()) so bytes survive reruns until the save completes.
    box_bytes = st.session_state.get(f"box_file_bytes_{company_id}")
    box_name  = st.session_state.get(f"box_file_name_{company_id}")
    if box_bytes and not uploaded:
        import io
        uploaded = io.BytesIO(box_bytes)
        uploaded.name = box_name
        uploaded.size = len(box_bytes)

    # Session-state key names
    ss_fkey       = f"upload_fkey_{company_id}"
    ss_parsed     = f"upload_parsed_{company_id}"
    ss_skip       = f"upload_skip_{company_id}"
    ss_saved      = f"upload_saved_{company_id}"
    ss_snap       = f"upload_snap_{company_id}"   # saved-periods snapshot for success display
    ss_commentary = f"upload_commentary_{company_id}"

    if uploaded is None and not st.session_state.get(f"box_file_bytes_{company_id}"):
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

        _write_ts   = st.session_state.get(f"upload_write_ts_{company_id}", "unknown")
        _db_path    = st.session_state.get(f"upload_db_path_{company_id}", DB_PATH)
        _row_count  = st.session_state.get(f"upload_row_count_{company_id}", "?")
        st.markdown(
            f"<div style='background:#E8F5E9;border:1px solid #2E7D32;border-radius:8px;"
            f"padding:12px 18px;font-size:13px;color:#2E7D32;font-weight:600;margin-bottom:6px'>"
            f"✓ {len(snap)} period(s) saved. Charts and benchmarking now reflect the "
            f"updated data.</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='font-size:11px;color:{MUTED};margin-bottom:14px;line-height:1.7'>"
            f"Data last updated: <b>{_write_ts}</b><br>"
            f"Total rows in DB: <b>{_row_count}</b><br>"
            f"DB path: <code style='font-size:10px'>{_db_path}</code>"
            f"</div>",
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
    is_pdf = st.session_state.get(f"box_is_pdf_{company_id}", False)
    if st.session_state.get(ss_fkey) != file_key:
        _spinner_msg = "Extracting KPIs from PDF using AI…" if is_pdf else "Reading and parsing Excel file…"
        with st.spinner(_spinner_msg):
            try:
                file_bytes = uploaded.read() if hasattr(uploaded, "read") else uploaded
                if hasattr(uploaded, "seek"):
                    uploaded.seek(0)

                if is_pdf:
                    all_rows, low_conf = _extract_kpis_from_pdf(file_bytes, company_name)
                    st.session_state[f"box_pdf_low_conf_{company_id}"] = low_conf
                else:
                    all_rows = PARSERS[company_name](file_bytes)

                existing         = _existing_periods(company_id)
                existing_in_file = sum(1 for r in all_rows if r["period_end_date"] in existing)

                st.session_state[ss_fkey]   = file_key
                st.session_state[ss_parsed] = all_rows
                st.session_state[ss_skip]   = existing_in_file
            except Exception as exc:
                st.error(f"Parse error: {exc}")
                return

    cached_rows  = st.session_state.get(ss_parsed, [])
    existing     = _existing_periods(company_id)
    new_rows     = [r for r in cached_rows if r["period_end_date"] not in existing]
    update_rows  = [r for r in cached_rows if r["period_end_date"] in existing]
    skipped      = st.session_state.get(ss_skip, 0)

    if skipped:
        st.markdown(
            f"<div style='color:{MUTED};font-size:12px;margin-bottom:8px'>"
            f"{skipped} period(s) already in database — will be updated with new values.</div>",
            unsafe_allow_html=True,
        )

    if not cached_rows:
        st.markdown(
            f"<div style='background:{WHITE};border:1px solid {BORDER};border-radius:10px;"
            f"padding:28px;text-align:center;color:{MUTED};font-size:13px'>"
            f"No period data found in this file.</div>",
            unsafe_allow_html=True,
        )
        return

    # ── PDF WARNING + AMBER FLAGS ─────────────────────────────────────────────
    if is_pdf:
        _pdf_fname = st.session_state.get(f"box_file_name_{company_id}", "PDF")
        st.markdown(
            f"<div style='background:#FFF3E0;border:1px solid #E65100;border-radius:8px;"
            f"padding:12px 18px;font-size:13px;color:#E65100;margin-bottom:10px'>"
            f"⚠ This data was extracted from a PDF (<b>{_pdf_fname}</b>) using AI. "
            f"PDF extraction can contain errors. Verify every figure against the source "
            f"document before saving. An Excel KPI file is the preferred source where "
            f"available.</div>",
            unsafe_allow_html=True,
        )
        _low_conf = st.session_state.get(f"box_pdf_low_conf_{company_id}", [])
        if _low_conf:
            st.markdown(
                f"<div style='background:#FFFDE7;border:1px solid #F9A825;border-radius:8px;"
                f"padding:10px 16px;font-size:12px;color:#F57F17;margin-bottom:10px'>"
                f"Verify these figures against the source PDF: "
                f"<b>{', '.join(_low_conf)}</b></div>",
                unsafe_allow_html=True,
            )

    # ── PREVIEW TABLE ─────────────────────────────────────────────────────────
    _preview_label = f"{len(new_rows)} new period(s)"
    if update_rows:
        _preview_label += f" + {len(update_rows)} update(s)"
    _preview_label += " — review before saving"
    st.markdown(
        f"<div style='font-size:11px;text-transform:uppercase;letter-spacing:.6px;"
        f"color:{MUTED};font-weight:600;margin-bottom:10px'>{_preview_label}</div>",
        unsafe_allow_html=True,
    )
    if is_pdf:
        _edited_df = st.data_editor(
            _build_preview_df(cached_rows),
            use_container_width=True,
            hide_index=True,
            key=f"pdf_editor_{company_id}_{file_key}",
        )
        # Sync edits back into cached_rows so confirm saves the corrected values
        _col_map = {c: c for c in _edited_df.columns}
        _edited_rows = _edited_df.to_dict(orient="records")
        for orig, edited in zip(cached_rows, _edited_rows):
            for col, val in edited.items():
                if col in orig:
                    orig[col] = None if (val != val or val == "") else val  # NaN → None
        st.session_state[ss_parsed] = cached_rows
    else:
        st.dataframe(_build_preview_df(cached_rows), use_container_width=True, hide_index=True)

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
            f"Saves {len(cached_rows)} period(s) for {company_name}. "
            f"Existing periods will be updated with new values.</div>",
            unsafe_allow_html=True,
        )

    if confirm:
        with st.spinner(f"Saving {len(cached_rows)} period(s) to database…"):
            print(f"[upload confirm] Writing {len(cached_rows)} rows | DB={DB_PATH} | company_id={company_id}")
            for p in cached_rows:
                _upsert_kpi(company_id, p)
            _recompute_growth(company_id)

        # ── Post-write verification ────────────────────────────────────────────
        _vconn = _conn()
        try:
            _vcur = _vconn.cursor()
            _vcur.execute(
                "SELECT COUNT(*) FROM kpi_snapshots WHERE company_id=%s", (company_id,)
            )
            _vcount = _vcur.fetchone()[0]
            _vcur.execute(
                "SELECT period_end_date, revenue_usd, updated_at FROM kpi_snapshots "
                "WHERE company_id=%s ORDER BY period_end_date DESC LIMIT 3",
                (company_id,),
            )
            _vrows = _vcur.fetchall()
            _vcur.close()
        finally:
            _vconn.close()
        _write_ts = _vrows[0][2] if _vrows else "unknown"
        print(f"[upload verify] DB={DB_PATH} company_id={company_id} "
              f"total_rows={_vcount} latest_3={_vrows}")

        with st.spinner("Generating AI commentary…"):
            commentary = _generate_commentary(
                company_name, str(info.get("sector", "")), cached_rows
            )

        st.session_state[ss_snap]       = list(cached_rows)
        st.session_state[ss_commentary] = commentary
        st.session_state[ss_saved]      = True
        st.session_state[f"upload_write_ts_{company_id}"]    = _write_ts
        st.session_state[f"upload_db_path_{company_id}"]     = DB_PATH
        st.session_state[f"upload_row_count_{company_id}"]   = _vcount
        st.session_state.pop(ss_fkey, None)
        st.session_state.pop(f"box_file_bytes_{company_id}", None)
        st.session_state.pop(f"box_file_name_{company_id}", None)
        st.session_state.pop(f"box_is_pdf_{company_id}", None)
        st.session_state.pop(f"box_pdf_low_conf_{company_id}", None)
        for _k in [k for k in st.session_state if k.startswith("_ws_")]:
            del st.session_state[_k]
        st.session_state["_cache_warmed"] = False
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

def fetch_last_affinity_note_for_buyer(buyer_name: str, affinity_api_key: str, company_being_exited: str = "") -> dict | None:
    try:
        import requests, difflib, json as _json, anthropic as _anthropic

        AUTH = ("", affinity_api_key)
        BASE = "https://api.affinity.co"

        # ── Step 1: Find the right org using Claude disambiguation ──────────
        r = requests.get(f"{BASE}/organizations", params={"term": buyer_name}, auth=AUTH, timeout=15)
        r.raise_for_status()
        candidates = r.json().get("organizations", [])
        if not candidates:
            return None

        org_id = None

        if len(candidates) == 1:
            org_id = candidates[0]["id"]
        else:
            # Build candidate list for Claude to disambiguate
            cand_lines = []
            for i, c in enumerate(candidates[:5]):
                name = c.get("name", "")
                domain = (c.get("domain") or c.get("domains") or [""])[0] if isinstance(c.get("domains"), list) else c.get("domain", "")
                desc = (c.get("description") or "")[:120]
                cand_lines.append(f"{i}: name={name} | domain={domain} | description={desc}")

            cache_key = f"_affinity_org_disambig_{buyer_name}"
            if cache_key in st.session_state:
                org_id = st.session_state[cache_key]
            else:
                try:
                    _ac = _anthropic.Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))
                    _msg = _ac.messages.create(
                        model="claude-haiku-4-5",
                        max_tokens=10,
                        messages=[{
                            "role": "user",
                            "content": (
                                f"We are looking for the Affinity CRM record for '{buyer_name}' "
                                f"in the context of African fintech venture capital and exit processes. "
                                f"Candidates:\n" + "\n".join(cand_lines) +
                                f"\nWhich index (0-{len(cand_lines)-1}) best matches '{buyer_name}' "
                                f"as a financial investor or strategic acquirer active in African fintech? "
                                f"Reply with only the integer index. If none match, reply -1."
                            )
                        }]
                    )
                    idx = int(_msg.content[0].text.strip())
                    org_id = candidates[idx]["id"] if 0 <= idx < len(candidates) else None
                except Exception:
                    # Fallback: best fuzzy name match
                    scored = sorted(
                        candidates[:5],
                        key=lambda c: difflib.SequenceMatcher(None, buyer_name.lower(), c.get("name","").lower()).ratio(),
                        reverse=True
                    )
                    best = scored[0]
                    if difflib.SequenceMatcher(None, buyer_name.lower(), best.get("name","").lower()).ratio() >= 0.6:
                        org_id = best["id"]

                st.session_state[cache_key] = org_id

        if not org_id:
            return None

        # ── Step 2: Fetch org-level notes ────────────────────────────────────
        r = requests.get(f"{BASE}/notes", params={"organization_id": org_id}, auth=AUTH, timeout=15)
        r.raise_for_status()
        all_notes = r.json().get("notes", [])

        # ── Step 3: Fetch person-level notes for all persons in this org ─────
        try:
            r2 = requests.get(f"{BASE}/organizations/{org_id}", auth=AUTH, timeout=15)
            person_ids = r2.json().get("person_ids", [])
            for pid in person_ids[:8]:  # cap at 8 to avoid rate limits
                rp = requests.get(f"{BASE}/notes", params={"person_id": pid}, auth=AUTH, timeout=10)
                if rp.ok:
                    all_notes.extend(rp.json().get("notes", []))
        except Exception:
            pass

        # Deduplicate by note id
        seen_ids = set()
        deduped = []
        for n in all_notes:
            nid = n.get("id")
            if nid not in seen_ids:
                seen_ids.add(nid)
                deduped.append(n)

        if not deduped:
            return None

        # ── Step 4: Parse dates and sort DESCENDING — most recent first ──────
        def _note_dt(n):
            raw = n.get("created_at", "")
            if not raw:
                return datetime.min.replace(tzinfo=timezone.utc)
            dt = datetime.fromisoformat(raw)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)

        deduped.sort(key=_note_dt, reverse=True)  # newest first

        # ── Step 5: Stale check on the single most recent note ───────────────
        most_recent_dt = _note_dt(deduped[0])
        date_str = most_recent_dt.strftime("%Y-%m-%d")
        cutoff_90 = datetime.now(tz=timezone.utc) - timedelta(days=90)
        if most_recent_dt < cutoff_90:
            return {"date": date_str, "creator_name": None, "snippet": None, "stale": True}

        # ── Step 6: Exit-relevance keyword scoring across all recent notes ────
        exit_keywords = [
            "acquisition", "acquire", "acquirer", "m&a", "strategic",
            "exit", "secondary", "liquidity", "sell", "sale", "buyer",
            "term sheet", "loi", "due diligence", "valuation", "offer",
            "interest", "mandate", "process", "banker", "transaction",
            "verto", "yoco", "lulalend", "cowrywise", "khazna", "enza",
            "twinco", "maxsoko", "sava",
        ]
        if company_being_exited:
            exit_keywords.append(company_being_exited.lower())

        def _score(n):
            text = (n.get("content") or "").lower()
            return sum(1 for kw in exit_keywords if kw in text)

        # Score all notes, sort by (score DESC, date DESC)
        scored_notes = sorted(deduped, key=lambda n: (_score(n), _note_dt(n).timestamp()), reverse=True)

        # Take top 3 by relevance score; if all score 0, take 3 most recent
        top_notes = scored_notes[:3]

        # ── Step 7: Claude summarization ─────────────────────────────────────
        note_texts = []
        for n in top_notes:
            content = (n.get("content") or "").strip()
            dt_str = _note_dt(n).strftime("%Y-%m-%d")
            note_texts.append(f"[{dt_str}] {content[:800]}")

        combined = "\n---\n".join(note_texts)
        company_context = company_being_exited or "the portfolio company being exited"

        summary = None
        try:
            _ac = _anthropic.Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))
            _msg = _ac.messages.create(
                model="claude-haiku-4-5",
                max_tokens=80,
                messages=[{
                    "role": "user",
                    "content": (
                        f"You are a VC associate reviewing CRM notes about '{buyer_name}' "
                        f"in the context of a potential acquisition of {company_context}, "
                        f"a B2B fintech company.\n\n"
                        f"Notes:\n{combined}\n\n"
                        f"Write ONE sentence (max 25 words) summarizing what was discussed "
                        f"specifically relating to an acquisition, investment, or exit of "
                        f"{company_context}. Include who spoke to whom if mentioned.\n"
                        f"If none of these notes discuss an acquisition, exit, or investment "
                        f"relating to {company_context}, respond with exactly: "
                        f"'No exit-relevant discussion found.'\n"
                        f"Do not summarize LP fundraising, portfolio construction, or "
                        f"general market commentary."
                    )
                }]
            )
            summary = _msg.content[0].text.strip()
        except Exception:
            # Fallback: first 150 chars of most recent note ending at word boundary
            fallback_content = (top_notes[0].get("content") or "").strip()
            if len(fallback_content) > 150:
                fallback_content = fallback_content[:150].rsplit(" ", 1)[0] + "…"
            summary = fallback_content

        if not summary:
            summary = "No exit-relevant discussion found."

        return {
            "date":         date_str,
            "creator_name": None,
            "snippet":      summary,
            "stale":        False,
        }

    except Exception:
        return None


# ── Shared: exit document upload / viewer section ────────────────────────────

def _render_exit_deck_section(company_id: int, company_name: str) -> None:
    """Render the exit document upload/view section at the bottom of any exit tab."""
    import base64
    import streamlit.components.v1 as components

    st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='font-size:11px;font-weight:700;letter-spacing:.8px;"
        f"text-transform:uppercase;color:{MUTED};margin-bottom:12px'>Exit Documents</div>",
        unsafe_allow_html=True,
    )

    view_key = f"viewing_doc_{company_id}"

    # Inline PDF viewer using PDF.js — rendered above the document list
    if st.session_state.get(view_key):
        doc_id = st.session_state[view_key]
        raw = get_exit_document(doc_id)
        if raw:
            b64 = base64.b64encode(raw).decode("utf-8")
            pdfjs_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #2C2C2A; font-family: sans-serif; }}
  #controls {{
    display: flex; align-items: center; gap: 10px;
    padding: 10px 14px; background: #2C2C2A; border-bottom: 1px solid #3a3a38;
  }}
  #controls button {{
    background: #D5FA94; color: #2C2C2A; border: none; border-radius: 5px;
    padding: 5px 14px; font-size: 13px; font-weight: 600; cursor: pointer;
  }}
  #controls button:disabled {{ background: #555; color: #999; cursor: default; }}
  #page-info {{ color: #D4D5CE; font-size: 13px; }}
  #canvas-container {{
    display: flex; justify-content: center;
    padding: 16px; height: 830px; overflow-y: auto; background: #2C2C2A;
  }}
  canvas {{ box-shadow: 0 2px 12px rgba(0,0,0,0.5); border-radius: 4px; }}
</style>
</head>
<body>
<div id="controls">
  <button id="prev" disabled>&#8592; Prev</button>
  <span id="page-info">Page 1 of ?</span>
  <button id="next" disabled>Next &#8594;</button>
</div>
<div id="canvas-container">
  <canvas id="pdf-canvas"></canvas>
</div>
<script>
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const b64 = "{b64}";
  const raw = atob(b64);
  const bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);

  let pdfDoc = null, currentPage = 1, scale = 1.4;
  const canvas = document.getElementById("pdf-canvas");
  const ctx = canvas.getContext("2d");
  const prevBtn = document.getElementById("prev");
  const nextBtn = document.getElementById("next");
  const pageInfo = document.getElementById("page-info");

  function renderPage(num) {{
    pdfDoc.getPage(num).then(page => {{
      const vp = page.getViewport({{ scale }});
      canvas.width = vp.width;
      canvas.height = vp.height;
      page.render({{ canvasContext: ctx, viewport: vp }}).promise.then(() => {{
        pageInfo.textContent = "Page " + num + " of " + pdfDoc.numPages;
        prevBtn.disabled = num <= 1;
        nextBtn.disabled = num >= pdfDoc.numPages;
      }});
    }});
  }}

  pdfjsLib.getDocument({{ data: bytes }}).promise.then(pdf => {{
    pdfDoc = pdf;
    renderPage(1);
  }});

  prevBtn.addEventListener("click", () => {{ if (currentPage > 1) renderPage(--currentPage); }});
  nextBtn.addEventListener("click", () => {{ if (currentPage < pdfDoc.numPages) renderPage(++currentPage); }});
</script>
</body>
</html>"""
            components.html(pdfjs_html, height=880, scrolling=False)
            if st.button("Close viewer", key=f"close_viewer_{company_id}"):
                st.session_state[view_key] = None
                st.rerun()
        else:
            st.session_state[view_key] = None

    docs = list_exit_documents(company_id)
    if docs:
        for doc in docs:
            size_kb = f"{doc['file_size'] // 1024:,} KB" if doc['file_size'] else "—"
            date_str = doc['uploaded_at'][:10] if doc['uploaded_at'] else "—"
            col_name, col_meta, col_read, col_dl, col_del = st.columns([3, 2, 1, 1.2, 0.8])
            with col_name:
                st.markdown(
                    f"<div style='font-size:13px;font-weight:600;color:#2C2C2A;padding-top:6px'>"
                    f"📄 {doc['doc_name']}</div>"
                    + (f"<div style='font-size:11px;color:{MUTED}'>{doc['notes']}</div>"
                       if doc['notes'] else ""),
                    unsafe_allow_html=True,
                )
            with col_meta:
                st.markdown(
                    f"<div style='font-size:12px;color:{MUTED};padding-top:6px'>{size_kb} · {date_str}</div>",
                    unsafe_allow_html=True,
                )
            with col_read:
                if st.button("📖 Read", key=f"read_exitdoc_{doc['id']}"):
                    st.session_state[view_key] = doc['id']
                    st.rerun()
            with col_dl:
                raw = get_exit_document(doc['id'])
                if raw:
                    st.download_button(
                        label="📥 Download",
                        data=raw,
                        file_name=doc['doc_name'] if doc['doc_name'].endswith('.pdf') else doc['doc_name'] + '.pdf',
                        mime="application/pdf",
                        key=f"dl_exitdoc_{doc['id']}",
                    )
            with col_del:
                if st.button("🗑", key=f"del_exitdoc_{doc['id']}", help="Delete this document"):
                    delete_exit_document(doc['id'])
                    st.rerun()
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    else:
        st.markdown(
            f"<div style='font-size:13px;color:{MUTED};margin-bottom:16px'>"
            "No documents uploaded yet.</div>",
            unsafe_allow_html=True,
        )

    with st.expander("Upload exit document", expanded=len(docs) == 0):
        st.markdown(
            f"<div style='font-size:12px;color:{MUTED};margin-bottom:10px'>"
            "Upload exit planning decks, banker materials, or process documents. "
            "PDFs only. Stored in the platform database.</div>",
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader(
            "Select PDF",
            type=["pdf"],
            key=f"exit_doc_uploader_{company_id}",
            label_visibility="collapsed",
        )
        label_input = st.text_input(
            "Document label (optional — defaults to filename)",
            key=f"exit_doc_label_{company_id}",
            placeholder="e.g. Cowrywise Exit Planning v1 — June 2026",
        )
        notes_input = st.text_input(
            "Notes (optional)",
            key=f"exit_doc_notes_{company_id}",
            placeholder="e.g. Prepared by Quona for board meeting",
        )
        if st.button("Save Document", key=f"exit_doc_save_{company_id}", type="primary"):
            if uploaded_file is None:
                st.warning("Please select a PDF file first.")
            else:
                doc_name = label_input.strip() if label_input.strip() else uploaded_file.name
                file_bytes = uploaded_file.read()
                save_exit_document(
                    company_id=company_id,
                    doc_name=doc_name,
                    file_bytes=file_bytes,
                    notes=notes_input.strip(),
                )
                st.success(f"✓ '{doc_name}' saved.")
                st.rerun()


# ── Cowrywise custom exit tab ────────────────────────────────────────────────

def _render_cowrywise_exit_tab() -> None:
    _cw_conn = _conn()
    _cw_row  = _cw_conn.execute("SELECT id FROM companies WHERE name = %s", ("Cowrywise",)).fetchone()
    _cw_conn.close()
    _company_id_cw = _cw_row[0] if _cw_row else 0

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    EMPTY     = "#D4D5CE"

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$50–100M"],
            "Founders comfortable staying; capital-light model sustains growth but limits investor liquidity. Quona secondary most likely near-term outcome.",
            [AMBER, AMBER, EMPTY], "Plan B — Q3 2027", False,
        ),
        (
            "MBO / Majority Acquisition",
            ["$80–150M", "4–8x revenue"],
            "Management buyout or majority acquisition that cleans up the cap table while keeping founders in seat — flagged as viable given founder preference",
            [AMBER, AMBER, EMPTY], "Founder preferred path", False,
        ),
        (
            "Strategic Sale",
            ["$100–200M", "5–10x revenue"],
            "Acquisition by Nigerian bank, pan-African insurer, or strategic fintech — most value-maximising path at $20M+ ARR",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Plan A — Q2 2027", True,
        ),
        (
            "PE Acquisition",
            ["$80–150M", "4–8x revenue"],
            "KKR or IFC-led transaction providing full or partial liquidity — both parties already in active dialogue with the company",
            [GREEN_DOT, AMBER, EMPTY], "Active conversations", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    cowrywise_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'Cowrywise' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not cowrywise_id_row.empty:
        cowrywise_id = int(cowrywise_id_row.iloc[0]["id"])
        ltm_df       = load_ltm_revenue(db_version=_db_global_version())
        _crow        = ltm_df[ltm_df["id"] == cowrywise_id]
        if not _crow.empty and _crow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_crow.iloc[0]["ltm_revenue"])

    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on EM digital wealth platform multiples and active KKR / IFC dialogue. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale",
        "Plan A — Q2 2027", GREEN, BLACK,
        "5–10x Revenue",
        r * 5, r * 7.5, r * 10,
        "#2E7D32",
        "Nigerian digital financial services acquisitions by banks and insurers typically price on "
        "user base and AUM scale — no direct Nigeria wealthtech comp at scale exists",
    )
    _val_row(
        "PE Acquisition",
        "Active conversations", BLUE, "#1565C0",
        "4–8x Revenue",
        r * 4, r * 6, r * 8,
        "#1565C0",
        "KKR Global Impact and IFC both in dialogue. PE buyers at this scale typically apply EBITDA "
        "or revenue multiples; $20M ARR target by end 2026 is the key trigger",
    )
    _val_row(
        "MBO / Majority Acquisition",
        "Founder preferred", "#D4D5CE", BLACK,
        "4–6x Revenue",
        r * 4, r * 5, r * 6,
        BLACK,
        "MBO pricing reflects control premium discount vs strategic sale; founders comfortable "
        "remaining post-transaction",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Plan B", "#D4D5CE", BLACK,
        "3–5x Revenue",
        r * 3, r * 4, r * 5,
        BLACK,
        "Secondary transaction at modest multiple if Plan A and PE routes do not materialise by Q3 2027",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative. No direct Nigeria wealthtech exit comp exists at meaningful scale. "
        f"Ranges anchored to Cowrywise's $20M ARR target (end 2026), comparable EM digital wealth platform "
        f"multiples, and active KKR and IFC dialogue. Actual exit value will depend on ARR achievement, "
        f"AUM scale, and competitive tension in any sale process."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("Moniepoint", "Very High",
         "Unicorn status (Nov 2024, $1B+ valuation); founder-to-founder engagement with Cowrywise already active",
         "Cowrywise's 1M+ user wealth platform extends Moniepoint's SME and consumer financial services into savings and investment"),
        ("Flutterwave", "High",
         "Acquired Mono Technologies (all-stock, 2025); building toward a full-stack African fintech platform",
         "Cowrywise's investment and savings layer would complete Flutterwave's consumer fintech stack alongside payments"),
        ("GTBank / HabariPay", "High",
         "Scaling HabariPay digital financial services; GTBank has one of Nigeria's largest retail customer bases",
         "Cowrywise's SEC-licensed wealth platform and 1M+ users give GTBank instant digital investment distribution"),
        ("Stanbic IBTC", "High",
         "Nigeria's leading wealth and asset management bank; actively digitising investment products",
         "Direct product overlap — Cowrywise's retail digital wealth platform would accelerate Stanbic IBTC's mass-market investment distribution"),
        ("PiggyVest", "Medium",
         "Paid out NGN 835B ($547M) to users in 2024; exploring broader investment products",
         "Merger would create Nigeria's dominant retail savings and investment platform, strengthening exit narrative ahead of a larger strategic sale"),
        ("ABSA", "High",
         "Pan-African banking group; active in Nigeria wealth and asset management",
         "ABSA's pan-African retail banking strategy and Nigeria wealth ambitions align directly with Cowrywise's platform and user base."),
    ]

    global_buyers = [
        ("KKR", "Very High",
         "$686B AUM; Global Impact and Growth Equity strategies actively targeting financial inclusion and EM fintech",
         "Already in active conversation with Cowrywise founders — highest near-term conviction buyer in the universe"),
        ("IFC", "Very High",
         "Active pan-African fintech investor; has been circling Cowrywise for several years",
         "IFC's financial inclusion mandate aligns directly with Cowrywise's mass-market Nigeria wealth platform — DFI financing or equity stake"),
        ("Old Mutual", "High",
         "Expanding digital wealth and insurance distribution across Africa; recently entered Nigerian banking",
         "Cowrywise's SEC-licensed platform and 1M+ users give Old Mutual instant Nigerian retail investment distribution"),
        ("Sanlam", "High",
         "Pan-African insurer with growing Nigerian retail financial services presence via Sanlam Nigeria",
         "Cowrywise's digital-first wealth distribution model complements Sanlam's Nigeria insurance and savings push"),
        ("Franklin Templeton", "Medium",
         "Active in Africa through local fund partnerships; expanding retail investment access across EM",
         "Cowrywise's fund distribution infrastructure could serve as Franklin Templeton's Nigeria retail investment channel"),
    ]

    secondaries_buyers = [
        ("LeapFrog Investments", "Very High",
         "$1B+ funds; purpose-led EM financial services specialist",
         "Perfect thesis fit — top conviction buyer alongside KKR. Financial inclusion mandate maps directly to Cowrywise's mass-market Nigeria wealth platform."),
        ("Actis", "High",
         "EM specialist with strong Africa track record; multiple Nigerian deals",
         "Natural fit at this ticket size. Africa book and financial services focus make Cowrywise a strong candidate."),
        ("DPI", "High",
         "Pan-African PE; multiple fintech deals including FlapKap; existing Quona co-investment relationship",
         "Strong warm intro path via Quona. DPI's fintech portfolio and Nigeria presence make Cowrywise a clear fit."),
        ("AfricInvest", "High",
         "Pan-African PE; financial services experience; active co-investor relationship with Quona",
         "Already in Power Series A target list. Active co-investor relationship provides natural entry point."),
        ("Helios Investment Partners", "Medium",
         "Largest pan-African PE; financial services experience; looking at HR fintech space",
         "Pan-African scale and financial services track record. Lower conviction than DPI and LeapFrog but in scope."),
        ("Warburg Pincus", "Medium",
         "Financial services focus; has done African deals; growth equity bracket fits trajectory",
         "In scope if Cowrywise targets $100M+ ARR. Financial services specialisation is relevant."),
        ("General Atlantic", "Medium",
         "Did Network International; knows MEA financial services; growth equity",
         "Digital wealth angle maps to their portfolio. MEA financial services track record is relevant."),
        ("Verod Capital / Adenia Partners", "Medium",
         "Verod is Nigerian PE; Adenia is pan-African; both credible at this ticket size",
         "Lower priority than DPI and AfricInvest but both are in scope. Verod's Nigeria focus is directly relevant."),
        ("Alphacode", "High",
         "$80M deployed across SA and Nigerian fintechs; latest deal August 2025",
         "Cowrywise is the leading Nigeria wealthtech — natural fit for Alphacode's fintech portfolio and RMI's wealth management ambitions"),
        ("Partech", "High",
         "Closed €280M second Africa fund 2024; one of most active Series A/B investors in Africa 2025",
         "Cowrywise at $20M ARR is exactly the growth-stage asset Partech's second fund targets"),
        ("Norrsken22", "Medium",
         "$205M fund; backed TymeBank and Stitch; Nigeria exposure limited",
         "Cowrywise would add Nigeria wealthtech exposure to a portfolio currently concentrated in SA and payments"),
    ]

    affinity_cache = st.session_state.get("cowrywise_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="cowrywise_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in local_buyers]
                + [g[0] for g in global_buyers]
                + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["cowrywise_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Cowrywise")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Local Buyers", "Global Buyers", "Secondaries Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_cowrywise_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_cowrywise_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_cowrywise_sec_" + name.replace(" ", "").replace("/", "")
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
        "Moniepoint":         "Escalate founder-to-founder conversation to M&A track. Frame Cowrywise's 1M+ user wealth platform as the savings and investment layer Moniepoint's SME ecosystem lacks.",
        "Flutterwave":        "Approach via shared investor network. Frame Cowrywise as completing Flutterwave's consumer fintech stack alongside payments — mirrors their Mono acquisition logic.",
        "GTBank / HabariPay": "Engage HabariPay leadership directly. Cowrywise's SEC licence and 1M+ users gives GTBank instant digital investment distribution without regulatory re-build.",
        "Stanbic IBTC":       "Approach via Quona board network. Direct overlap — Cowrywise's digital wealth platform accelerates Stanbic's mass-market investment distribution.",
        "PiggyVest":          "Initiate consolidation conversation via Cowrywise CEO. Combined platform would create Nigeria's dominant retail savings and investment brand.",
        "KKR":                "Continue active dialogue — highest near-term conviction buyer. Prepare investor materials for Global Impact strategy review.",
        "IFC":                "Re-engage via Quona's IFC relationship. Frame as DFI financing or equity stake aligned with IFC's financial inclusion mandate.",
        "Old Mutual":         "Approach Old Mutual Ventures Africa. Cowrywise's Nigerian platform gives Old Mutual instant retail investment distribution without regulatory build-out.",
        "Sanlam":             "Engage Sanlam Nigeria strategy team. Frame as accelerating their savings and insurance cross-sell via Cowrywise's digital-first distribution.",
        "Franklin Templeton":  "Approach via fund distribution partnership — convert to equity stake conversation once ARR exceeds $20M.",
        "Alphacode":          "Approach via RMI/FNB connections. Frame as bridging Cowrywise to a local bank acquirer — Alphacode's portfolio relationships are the natural path.",
        "Partech":            "Flag for Partech's second Africa fund mandate — Cowrywise at $20M ARR is exactly the growth-stage asset they are targeting.",
        "Norrsken22":         "Approach as a secondary or bridge investment ahead of strategic sale — Cowrywise adds Nigeria wealthtech exposure to a portfolio concentrated in SA.",
    }
    _ALL_BUYERS = [b[0] for b in local_buyers] + [b[0] for b in global_buyers] + [b[0] for b in secondaries_buyers]

    if st.button("Generate Exit Actions for Cowrywise"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get("engage_cowrywise_" + name.replace(" ", "").replace("/", ""), False)
            or st.session_state.get("engage_cowrywise_sec_" + name.replace(" ", "").replace("/", ""), False)
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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

    _render_exit_deck_section(_company_id_cw, "Cowrywise")


# ── VertoFX custom exit tab ──────────────────────────────────────────────────

def _render_vertofx_exit_tab() -> None:
    _vfx_conn = _conn()
    _vfx_row  = _vfx_conn.execute("SELECT id FROM companies WHERE name = %s", ("Verto",)).fetchone()
    _vfx_conn.close()
    _company_id_verto = _vfx_row[0] if _vfx_row else 0

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    RED_DOT   = "#E57373"
    EMPTY     = "#D4D5CE"

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$150–400M", "3–5x revenue"],
            "Continue scaling B2B FX rails and treasury services across Africa and UAE corridors — but investor timeline pressure is growing",
            [AMBER, AMBER, EMPTY], "Unattractive strategically", False,
        ),
        (
            "Strategic Sale to Global PSP",
            ["$150–220M", "Comp median implies ~$362M at 5.8x revenue — treat as ceiling, not base"],
            "Acquisition by a global payments infrastructure player seeking Africa and EM corridor access — most realistic path given consolidation in B2B FX",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — 24–36 months", True,
        ),
        (
            "Strategic Sale to African Bank",
            ["$50–150M", "2–4x revenue"],
            "Acquisition by a pan-African bank seeking to own FX infrastructure and reduce correspondent banking costs",
            [AMBER, AMBER, EMPTY], "Possible — dependent on scale", False,
        ),
        (
            "PE Recap / Acquihire",
            ["$50–100M"],
            "Private equity recapitalisation or acquihire providing investor liquidity while business continues to scale",
            [AMBER, EMPTY, EMPTY], "Last resort", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    vertofx_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'VertoFX' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not vertofx_id_row.empty:
        vertofx_id = int(vertofx_id_row.iloc[0]["id"])
        ltm_df     = load_ltm_revenue(db_version=_db_global_version())
        _vrow      = ltm_df[ltm_df["id"] == vertofx_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on comparable exit multiples and B2B FX transaction benchmarks. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale to Global PSP",
        "Most likely — 24–36 months", GREEN, BLACK,
        "4–8x Revenue",
        r * 4, r * 6, r * 8,
        "#2E7D32",
        "Consistent with Airwallex ($6.2B at ~8x ARR), Wise public comp at 5x revenue, "
        "and Corpay acquisition of GPS Capital Markets (B2B FX, 2024)",
    )
    _val_row(
        "Strategic Sale to African Bank",
        "Possible", BLUE, "#1565C0",
        "2–4x Revenue",
        r * 2, r * 3, r * 4,
        "#1565C0",
        "African bank acquisitions of fintech infrastructure typically price on strategic value not revenue — "
        "Nedbank-iKhokha and Lesaka-Bank Zero as reference points",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Unattractive", "#D4D5CE", BLACK,
        "3–5x Revenue",
        r * 3, r * 4, r * 5,
        BLACK,
        "Secondary transaction at modest multiple; consistent with Payoneer public comp at ~2x "
        "and Wise at ~5x revenue",
    )
    _val_row(
        "PE Recap / Acquihire",
        "Last resort", "#D4D5CE", BLACK,
        "1–2x Revenue",
        r * 1, r * 1.5, r * 2,
        BLACK,
        "Distressed or opportunistic transaction; acquihire value driven by team and licenses not revenue",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative. Comps include Airwallex Series F ($6.2B valuation, 2025), "
        f"Wise public trading (~5x revenue), Payoneer public trading (~2x revenue), "
        f"Corpay acquisition of GPS Capital Markets (B2B FX, 2024), "
        f"and Paystack–Stripe ($200M+, 2020)."
        f"</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='font-size:11px;color:#93A3A1;padding:8px;background:#EFF0EA;"
        "border-radius:6px;margin-top:12px'>"
        "Comp-median multiples reflect peak exit conditions. Base case assumes a strategic "
        "sale at a discount to comp median given current market conditions and single-buyer "
        "dynamics. Use $150–220M as the working base for LP reporting."
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("Standard Bank / Stanbic", "High",
         "Africa's largest bank by assets; actively building trade finance and cross-border payment capabilities across 20+ African markets",
         "Verto's EM FX rails reduce Standard Bank's correspondent banking costs and deepen its SME trade finance offering"),
        ("Access Bank", "High",
         "Expanded to 20+ countries; positioning as Africa's gateway bank for trade and cross-border flows",
         "Verto's multi-currency infrastructure and UAE corridor directly supports Access Bank's pan-African and diaspora payment strategy"),
        ("Ecobank", "Medium",
         "Operates in 35 African countries; focused on deepening intra-African trade payments",
         "Verto's B2B FX and treasury tools would strengthen Ecobank's corporate banking product across its African footprint"),
        ("FirstBank Nigeria", "Medium",
         "Scaling digital corporate banking and trade finance products",
         "Verto's FX infrastructure could power FirstBank's B2B cross-border offering for Nigerian corporates trading across Africa and the UAE"),
    ]

    global_buyers = [
        ("Corpay", "Very High",
         "Acquired GPS Capital Markets (B2B FX treasury, 2024) — directly comparable to Verto's model; scaling Corporate Payments to $2B by 2026",
         "Verto's Africa and EM corridor coverage fills a gap in Corpay's global B2B FX footprint"),
        ("Nium", "Very High",
         "$1.4B valuation; actively expanding real-time payout rails across Africa and the Middle East",
         "Verto's licensed EM infrastructure and $25B+ in annual volume would materially accelerate Nium's Africa expansion"),
        ("Thunes", "High",
         "Raised $150M Series D (2024); expanding cross-border payment corridors in Africa and Asia",
         "Verto's B2B FX rails and treasury tools are highly complementary to Thunes' cross-border payout network"),
        ("Ripple", "Very High",
         "Scaling RippleNet B2B cross-border payments globally; BofA relationship confirmed; active conversations with VertoFX corridor team",
         "Ripple's B2B FX rails and Verto's Africa/EM licensed infrastructure are directly complementary — acquisition accelerates Ripple's emerging market corridor strategy"),
        ("Terrapay", "High",
         "Cross-border B2B payments network operating across 120+ countries; actively consolidating Africa FX capabilities",
         "Terrapay's network expansion into Africa aligns with Verto's licensed corridor infrastructure — acquisition or deep commercial partnership both credible paths"),
        ("Airwallex", "High",
         "Raised $300M Series F at $6.2B valuation (2025); expanding into UAE and EM markets",
         "Verto's Africa corridor expertise and DFSA license complement Airwallex's UAE expansion; acquisition would fast-track African B2B payments"),
        ("Mastercard", "Medium",
         "Invested $200M in MTN MoMo; scaling B2B cross-border and trade finance products in Africa",
         "Verto's multi-currency infrastructure and enterprise client base align with Mastercard's B2B payments and trade finance agenda"),
        ("Wise", "Medium",
         "Scaling Wise Business (B2B) globally; processing $130B+ annually",
         "Verto's EM corridor specialisation and African licensing complement Wise's B2B expansion into frontier markets"),
    ]

    secondaries_buyers = [
        ("Partech", "High",
         "Closed €280M second Africa fund (2024); top Series A/B investor in Africa in 2025",
         "Verto is a de-risked B2B fintech with enterprise clients and growing UAE corridor — fits Partech's fintech mandate"),
        ("Norrsken22", "Medium",
         "$205M fund; backed Stitch and TymeBank in SA payments and banking",
         "Verto deepens their Africa fintech exposure in B2B payments, a segment they don't yet have direct coverage in"),
        ("Blue Earth Capital", "High",
         "Ran GP-led secondary in Moniepoint (Oct 2025); has since inquired directly about VertoFX — confirms active Africa B2B fintech secondaries mandate",
         "Blue Earth ran the Moniepoint process at a comparable stage and ticket size, and has expressed direct interest in Verto. Highest-conviction secondaries buyer in the universe."),
        ("Lightrock", "High",
         "Impact-focused growth equity; $2B+ AUM; active in African fintech",
         "Lightrock's financial inclusion thesis and B2B fintech focus maps directly to VertoFX's cross-border payments corridor for African businesses."),
        ("Apis Partners", "Medium",
         "Specialist financial services growth equity; $1.5B+ AUM; multiple African fintech investments — note: typical check sizes ($50M+) sit above current implied secondary ticket range for VertoFX",
         "Apis has deep fintech expertise in emerging markets and has backed B2B payments and FX players. VertoFX's enterprise corridor business is a strong fit for their mandate."),
    ]

    affinity_cache = st.session_state.get("vertofx_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="vertofx_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in local_buyers]
                + [g[0] for g in global_buyers]
                + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["vertofx_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "VertoFX")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Local Buyers", "Global Buyers", "Secondaries Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_vertofx_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_vertofx_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_vertofx_sec_" + name.replace(" ", "").replace("/", "")
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
        "Standard Bank / Stanbic": "Engage Standard Bank CIB leadership via Quona network. Frame Verto as reducing correspondent banking costs and deepening SME trade finance across 20+ African markets.",
        "Access Bank":             "Approach Access Bank's digital banking team. Verto's multi-currency rails and UAE corridor directly support Access Bank's diaspora and pan-African trade strategy.",
        "Ecobank":                 "Engage via Quona's West Africa network. Frame Verto as deepening intra-African trade payment infrastructure across Ecobank's 35-country footprint.",
        "FirstBank Nigeria":       "Approach via existing Quona-FirstBank relationship. Verto's FX infrastructure could power FirstBank's B2B cross-border offering for Nigerian corporate clients.",
        "Corpay":                  "Approach via investment banking intermediary. Verto's Africa and EM corridor coverage fills the most critical gap in Corpay's global B2B FX footprint.",
        "Nium":                    "Escalate existing corridor partnership conversation to M&A track. Verto's $25B+ annual volume and licensed EM infrastructure materially accelerates Nium's Africa expansion.",
        "Thunes":                  "Approach via shared investor network. Verto's B2B FX treasury tools complement Thunes' cross-border payout network with enterprise-grade supply side.",
        "Ripple":                  "Escalate existing BofA-facilitated conversation to M&A track. Frame Verto's licensed Africa corridors as filling Ripple's most critical emerging-market gap.",
        "Terrapay":                "Approach Terrapay's Corp Dev team via shared network. Verto's $25B+ annual volume and licensed EM infrastructure are directly additive to Terrapay's 120-country network.",
        "Airwallex":               "Approach via Airwallex's UAE expansion team. Verto's DFSA licence and Africa corridor expertise are directly additive to Airwallex's Gulf build-out.",
        "Mastercard":              "Escalate Mastercard partnership conversation to strategic M&A review. Verto's enterprise infrastructure aligns with Mastercard's B2B payments and trade finance agenda.",
        "Wise":                    "Approach Wise Business development team. Verto's EM frontier corridor specialisation plugs the gap in Wise Business's B2B expansion into Africa.",
        "Partech":                 "Flag for Partech's second Africa fund — Verto is exactly the de-risked B2B fintech with enterprise clients they are seeking.",
        "Norrsken22":              "Approach as a secondary or bridge investment. Verto deepens Norrsken's Africa B2B payments exposure in a segment they don't yet cover.",
        "Blue Earth Capital":      "Assess impact narrative strengthening before outreach — Verto's financial inclusion angle needs to be front and centre for Blue Earth.",
    }
    _ALL_BUYERS = [b[0] for b in local_buyers] + [b[0] for b in global_buyers] + [b[0] for b in secondaries_buyers]

    if st.button("Generate Exit Actions for VertoFX"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get("engage_vertofx_" + name.replace(" ", "").replace("/", ""), False)
            or st.session_state.get("engage_vertofx_sec_" + name.replace(" ", "").replace("/", ""), False)
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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

    _render_exit_deck_section(_company_id_verto, "Verto")


# ── Lulalend custom exit tab ─────────────────────────────────────────────────

def _render_lulalend_exit_tab() -> None:
    _ll_conn = _conn()
    _ll_row  = _ll_conn.execute("SELECT id FROM companies WHERE name = %s", ("Lulalend",)).fetchone()
    _ll_conn.close()
    _company_id_lula = _ll_row[0] if _ll_row else 0

    load_comps_detail.clear()
    load_comp_mapping.clear()

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    RED_DOT   = "#E57373"
    EMPTY     = "#D4D5CE"

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$200–315M"],
            "Capital-intensive path requiring fresh equity and debt; investor exit pressure growing",
            [AMBER, AMBER, EMPTY], "Unattractive strategically", False,
        ),
        (
            "Merger with Yoco",
            ["$400–850M"],
            "Full-stack SME digital bank combining Yoco merchants with Lula lending — strong logic but investor misalignment",
            [RED_DOT, EMPTY, EMPTY], "Low feasibility", False,
        ),
        (
            "Strategic Sale Local",
            ["$200–315M", "3–4x revenue"],
            "Sell to FNB, Absa or Vodacom — most realistic path given active bank consolidation in SA SME",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — 12–24 months", True,
        ),
        (
            "Strategic Sale Global",
            ["$250–500M", "6–12x revenue"],
            "Exit to Revolut, Nubank or EM fintech — higher ceiling but limited near-term buyer appetite",
            [AMBER, AMBER, EMPTY], "Execution dependent", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    lulalend_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'Lulalend' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not lulalend_id_row.empty:
        lulalend_id = int(lulalend_id_row.iloc[0]["id"])
        ltm_df      = load_ltm_revenue(db_version=_db_global_version())
        _lrow       = ltm_df[ltm_df["id"] == lulalend_id]
        if not _lrow.empty and _lrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_lrow.iloc[0]["ltm_revenue"])

    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on comparable exit multiples and Bruwer ISP exit analysis. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale Local",
        "Most likely — 12–24 months", GREEN, BLACK,
        "3–4x Revenue",
        r * 3, r * 3.5, r * 4,
        "#2E7D32",
        "Consistent with TymeBank–Retail Capital (~$85–90M at 2.5x) and Absa RfP benchmarks (3.3–3.6x revenue)",
    )
    _val_row(
        "Strategic Sale Global",
        "Execution dependent", BLUE, "#1565C0",
        "6–12x Revenue",
        r * 6, r * 9, r * 12,
        "#1565C0",
        "Consistent with Revolut/Nubank EM fintech benchmarks. Requires profitability and neobank narrative",
    )
    _val_row(
        "Merger with Yoco",
        "Low feasibility", "#D4D5CE", BLACK,
        "Sum of parts",
        400e6, 600e6, 850e6,
        BLACK,
        "Combined value based on sum-of-parts plus ecosystem uplift. Investor alignment required",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Unattractive", "#D4D5CE", BLACK,
        "3.3–3.6x Revenue",
        r * 3.3, r * 3.45, r * 3.6,
        BLACK,
        "Based on Absa RfP implied multiples (6.6–7.2x GP). Capital-intensive without strategic premium",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative and based on comparable transaction multiples. "
        f"TymeBank–Retail Capital (2022), Absa RfP (2025), and Bruwer ISP exit analysis (May 2026) "
        f"used as primary references."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("FNB",      "Very High", "Aggressively scaling SME banking; actively building digital credit capabilities",
         "Lula's SME credit engine and borrower data accelerates FNB's SME strategy by years"),
        ("Absa",     "Very High", "Issued RfP for Lulalend in 2025, implying $200–315M valuation",
         "RfP already issued — most advanced conversation in the buyer universe"),
        ("Vodacom",  "High",      "Active VodaLend partnership with Lula already operational",
         "Acquiring Lula internalises all VodaLend economics and SME lending distribution"),
        ("Capitec",  "Medium",    "Expanding into SME banking and credit products",
         "Lula would fast-track Capitec's SME lending ambitions with proven underwriting models"),
        ("TymeBank", "Medium",    "Acquired Retail Capital (2022) — directly comparable SME lending deal",
         "Potential appetite for a second SME lending acquisition to deepen franchise"),
    ]

    global_buyers = [
        ("Revolut",    "High",   "Exploring South Africa market entry; stated intention publicly",
         "Lula gives Revolut a turnkey SME banking platform for its SA launch"),
        ("Nubank",     "High",   "Invested in TymeBank Series D (2024) — first African exposure",
         "Lula aligns with Nubank's low-cost digital credit DNA and accelerates African expansion"),
        ("Experian",   "Medium", "Acquired Compuscan SA (2019); building SME analytics via ExperiFin (2025)",
         "Lula's lending data enriches Experian's SME credit bureau and scoring models"),
        ("Moniepoint", "Medium", "Raised $250M Series C (2025); pursuing pan-African expansion",
         "Lula provides instant SA foothold aligned with Moniepoint's SME-first model"),
        ("Prosus",     "Medium", "Investments in Remitly, PayU, Moniepoint across Africa",
         "Lula fits Prosus's African fintech thesis and could bundle with PayU SME rails"),
    ]

    secondaries_buyers = [
        ("Blue Earth Capital", "High",   "GP-led secondary in Moniepoint (Oct 2025); dedicated Africa secondaries strategy launched 2024",
         "Actively building African secondaries market — Lula is a de-risked SA lending asset"),
        ("Alphacode",          "High",   "$80M deployed across SA fintechs; latest deal August 2025",
         "RMI/FNB relationships make them a natural bridge to a local bank exit"),
        ("Norrsken22",         "Medium", "Backed TymeBank and Stitch; $205M fund actively deploying",
         "Lula deepens SA fintech exposure with a de-risked revenue-generating lending asset"),
        ("Partech",            "Medium", "Closed €280M second Africa fund in 2024; top Series A/B investor in 2025",
         "Leading SA SME lender at a stage where comparable exits are validating the market"),
    ]

    affinity_cache = st.session_state.get("lulalend_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="lulalend_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in local_buyers]
                + [g[0] for g in global_buyers]
                + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["lulalend_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Lulalend")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Local Buyers", "Global Buyers", "Secondaries Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_lulalend_" + name.replace(" ", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_lulalend_" + name.replace(" ", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_lulalend_sec_" + name.replace(" ", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 4: Next Steps Generator ──────────────────────────────────────
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
        "FirstRand / RMB":   "Engage RMB corporate finance team — RfP already issued. Prepare vendor due diligence pack and term sheet response.",
        "Vodacom":           "Re-engage VodaLend partnership team at M&A level. Acquiring Lula internalises all VodaLend economics — frame as the natural next step.",
        "Capitec":           "Request meeting with Capitec SME banking lead. Lula's underwriting models and 110k+ active borrowers are the fastest path to Capitec SME credit scale.",
        "TymeBank":          "Approach TymeBank CEO directly. Second SME lending acquisition post-Retail Capital is consistent with their full-stack SME strategy.",
        "Revolut":           "Approach Revolut SA launch team. Lula provides a turnkey licensed SME lending platform for Revolut's SA market entry.",
        "Nubank":            "Approach via TymeBank board relationship. Lula's credit DNA and SA footprint align with Nubank's low-cost digital credit model and Africa expansion.",
        "Experian":          "Re-engage Experian ExperiFin team. Lula's lending data enriches Experian's SME credit bureau — convert data partnership to deeper strategic conversation.",
        "Moniepoint":        "Approach Moniepoint expansion team. Lula provides instant SA SME foothold aligned with Moniepoint's pan-African SME-first model.",
        "Prosus":            "Flag for Prosus African fintech thesis — Lula fits their PayU SME rails bundling strategy. Approach via PayU SA relationship.",
        "Blue Earth Capital":"Warm intro via Quona's Blue Earth relationship. Lula is a de-risked SA SME lending asset with clear secondaries market appeal.",
        "Alphacode":         "Approach via RMI/FNB connections — natural bridge to a local bank exit. Alphacode's portfolio relationships are the most efficient path.",
        "Norrsken22":        "Flag as bridge investment ahead of strategic sale — Lula adds de-risked SA fintech exposure to a portfolio thin on lending assets.",
        "Partech":           "Approach Partech's second Africa fund team. Lula at Series C stage with leading SA SME lender position is exactly their mandate.",
    }
    _ALL_BUYERS = [b[0] for b in local_buyers] + [b[0] for b in global_buyers] + [b[0] for b in secondaries_buyers]

    if st.button("Generate Exit Actions for Lulalend"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get("engage_lulalend_" + name.replace(" ", ""), False)
            or st.session_state.get("engage_lulalend_sec_" + name.replace(" ", ""), False)
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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

    _render_exit_deck_section(_company_id_lula, "Lulalend")


# ── Yoco custom exit tab ──────────────────────────────────────────────────────

def _render_yoco_exit_tab() -> None:
    # ── Exit deck section company_id lookup ──────────────────────
    try:
        _yoco_conn = _conn()
        _yoco_cur  = _yoco_conn.cursor()
        _yoco_cur.execute("SELECT id FROM companies WHERE name = %s", ("Yoco",))
        _yoco_row  = _yoco_cur.fetchone()
        _yoco_conn.close()
        _company_id_yoco = _yoco_row[0] if _yoco_row else 1
    except Exception:
        _company_id_yoco = 1

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
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
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

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    _yoco_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'Yoco' LIMIT 1", _conn()
    )
    _yoco_ltm = None
    if not _yoco_id_row.empty:
        _yoco_id = int(_yoco_id_row.iloc[0]["id"])
        _ltm_df  = load_ltm_revenue(db_version=_db_global_version())
        _vrow    = _ltm_df[_ltm_df["id"] == _yoco_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            _yoco_ltm = float(_vrow.iloc[0]["ltm_revenue"])

    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on SA payments M&A comps (iKhokha, Peach Payments) and strategic premium for SARB-licensed merchant base. "
        f"LTM Revenue: {fmt_usd(_yoco_ltm)}</div>",
        unsafe_allow_html=True,
    )

    _HDR_Y = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    _yhcols = st.columns([2, 1, 1, 1, 2])
    for _hc, _lbl in zip(_yhcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with _hc:
            st.markdown(f"<div style='{_HDR_Y}'>{_lbl}</div>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='height:2px;background:{BORDER};margin:6px 0 10px'></div>",
        unsafe_allow_html=True,
    )

    def _yval_row(pathway_name, tag, tag_bg, tag_fg, multiple_lbl, low, base, high, base_color, note):
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
            st.markdown(f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{multiple_lbl}</div>", unsafe_allow_html=True)
        with cols[2]:
            st.markdown(f"<div style='font-size:14px;color:{BLACK};padding-top:6px'>{fmt_usd(low)}</div>", unsafe_allow_html=True)
        with cols[3]:
            st.markdown(f"<div style='font-size:14px;font-weight:700;color:{base_color};padding-top:6px'>{fmt_usd(base)}</div>", unsafe_allow_html=True)
        with cols[4]:
            st.markdown(f"<div style='font-size:14px;color:{MUTED};padding-top:6px'>Up to {fmt_usd(high)}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:11px;color:{MUTED};font-style:italic;margin:4px 0 8px'>{note}</div>", unsafe_allow_html=True)
        st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>", unsafe_allow_html=True)

    _yr = _yoco_ltm or 0
    _yval_row("Strategic Sale Local", "Most likely — 12–24 months", GREEN, BLACK, "2–4x Revenue",
              _yr * 2, _yr * 3, _yr * 4, "#2E7D32",
              "iKhokha acquired by MTN at ~4x revenue ($94M at 4–5x ARR). Peach Payments (Mastercard, 2023) at ~3x. "
              "SA bank/telco deals capped below $400M — Vodacom and Capitec are highest-conviction buyers.")
    _yval_row("SME Bank Build (w/ SAVA)", "Execution dependent", BLUE, "#1565C0", "3–5x Revenue",
              _yr * 3, _yr * 4, _yr * 5, "#1565C0",
              "Yoco + SAVA banking licence creates SA's first full-stack SME challenger bank — commands a premium to a standalone payments exit. "
              "Requires 12–18 months to execute before exit.")
    _yval_row("Strategic Sale Global", "Low feasibility", "#D4D5CE", BLACK, "5–8x Revenue",
              _yr * 5, _yr * 6.5, _yr * 8, BLACK,
              "Stripe, Adyen, or Nubank Africa market-entry acquisition at global SaaS multiples — requires EBITDA breakeven and pan-African narrative first.")
    _yval_row("Remain Independent", "Unattractive near-term", "#D4D5CE", BLACK, "2–3x Revenue",
              _yr * 2, _yr * 2.5, _yr * 3, BLACK,
              "Secondary at modest multiple if strategic sale stalls — avoid unless no strategic process underway.")

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative. Primary comps: iKhokha–MTN ($94M, ~4–5x revenue, 2023), "
        f"Peach Payments–Mastercard (undisclosed, est. ~3x revenue, 2023), TymeBank–Retail Capital ($85–90M, ~2.5x revenue, 2022). "
        f"Yoco's 110k+ active merchant base and approaching EBITDA breakeven are the key near-term value inflection points."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
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

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
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
        ("Sipay",    "Medium", "Entered UAE; South Africa explicitly on radar",
         "Yoco provides instant credibility and merchant scale in SA. Appetite appears genuine given stated plans, though Africa execution track record is unproven."),
        ("Hala",     "Low-Medium", "Expanding beyond Saudi and GCC",
         "Yoco is a natural SA entry vehicle with a ready SME base. Prefers organic growth over acquisition."),
        ("Block",    "Low",    "No Africa activity; focused on US, EU, Asia Pacific",
         "Perfect model fit — Yoco is effectively Square for Africa. No emerging market interest currently."),
        ("PayPal",   "Low",    "Deprioritised Africa and emerging markets",
         "Strong model alignment via iZettle playbook but PayPal has explicitly deprioritised SA. Appetite limited."),
        ("Revolut",  "Low",    "Expanding acquiring and merchant services in new markets",
         "Yoco would give Revolut a turnkey SA entry. Loss-making and Africa far below current priority markets."),
        ("CreditInfo Group", "Low-Medium", "Acquired Experian Uganda and Namibia bureaus; launched SME360",
         "Live POS data would strengthen CreditInfo's SME product in southern markets. M&A capacity limited — likely partnership first."),
    ]

    secondaries_buyers = [
        ("Telkom",                     "High",   "Launched Yep SME marketplace, building merchant payments from scratch",
         "Yoco's 110k merchant network solves exactly what Telkom is trying to build"),
        ("Experian",                   "Medium", "Launched ExperiFin AI credit marketplace for SMEs in 2025",
         "Yoco's SME transaction data complements their credit bureau play"),
        ("Google Africa Investment Fund", "Medium", "Invested in Moniepoint Series C (2025) alongside Visa and IFC",
         "Mirrors their pattern of minority stakes in leading African fintechs"),
        ("Alphacode",                  "High",   "$80M deployed across SA fintechs, latest deal August 2025",
         "RMI/FNB relationships make them a natural bridge to a local bank exit"),
        ("Blue Earth Capital",         "High",   "GP-led secondary in Moniepoint (Oct 2025), dedicated Africa secondaries strategy launched 2024",
         "Actively building the African secondaries market, Yoco is a perfect fit"),
        ("Norrsken22",                 "Medium", "Backed Stitch and TymeBank, $205M fund actively deploying",
         "Deepens SA fintech exposure with a de-risked, revenue-generating asset"),
        ("Partech",                    "Medium", "Closed €280M second Africa fund in 2024, top Series A/B investor in 2025",
         "Leading SA merchant payments platform at a point when comparable exits are validating the market"),
    ]

    affinity_cache = st.session_state.get("yoco_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="yoco_affinity_sync"):
            _api_key = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in local_buyers] + [g[0] for g in global_buyers] + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["yoco_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Yoco")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Local Buyers", "Global Buyers", "Secondaries Buyers"])
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
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_yoco_sec_" + name.replace(" ", "")
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

    _render_exit_deck_section(_company_id_yoco, "Yoco")

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


# ── TWINCO custom exit tab ────────────────────────────────────────────────────

def _render_twinco_exit_tab() -> None:
    _tw_conn = _conn()
    _tw_row  = _tw_conn.execute("SELECT id FROM companies WHERE name = %s", ("TWINCO",)).fetchone()
    _tw_conn.close()
    _company_id_twinco = _tw_row[0] if _tw_row else 0

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    EMPTY     = "#D4D5CE"

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$150–300M", "8–15x ARR"],
            "Continue scaling PO finance platform globally — strong momentum post Series B but investor timeline pressure building",
            [AMBER, AMBER, EMPTY], "Unattractive strategically", False,
        ),
        (
            "Strategic Sale to Global Bank",
            ["$300–600M", "10–20x ARR"],
            "Acquisition by Santander, HSBC or JPMorgan to own the PO finance layer missing from their SCF stack — Santander already leads the securitisation facility",
            [GREEN_DOT, GREEN_DOT, AMBER], "Most likely — 24–48 months", True,
        ),
        (
            "Strategic Sale to Financial Infrastructure Player",
            ["$250–500M", "8–18x ARR"],
            "Acquisition by FIS, SAP or Mastercard to complete their supply chain finance stack — FIS just acquired Demica for $300M, Twinco is the complementary PO layer",
            [GREEN_DOT, GREEN_DOT, AMBER], "High strategic fit", False,
        ),
        (
            "IPO or DFI Full Acquisition",
            ["$200–400M", "7–15x ARR"],
            "FMO-led full acquisition or public listing as Twinco establishes PO finance as a recognised institutional asset class",
            [AMBER, AMBER, EMPTY], "Longer term — 4–5 years", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    _twinco_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name IN ('TWINCO', 'Twinco') LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not _twinco_id_row.empty:
        _twinco_id = int(_twinco_id_row.iloc[0]["id"])
        _ltm_df    = load_ltm_revenue(db_version=_db_global_version())
        _vrow      = _ltm_df[_ltm_df["id"] == _twinco_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on comparable exit multiples and supply chain finance benchmarks. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale to Global Bank",
        "Most likely — 24–48 months", GREEN, BLACK,
        "10–20x ARR",
        r * 10, r * 15, r * 20,
        "#2E7D32",
        "Santander leads €150M securitisation facility — natural path from funder to owner. "
        "Taulia acquired by SAP at ~17x ARR ($24M ARR, ~$400M deal). "
        "Demica acquired by FIS for $300M at ~0.75% of $40B AuA.",
    )
    _val_row(
        "Strategic Sale to Financial Infrastructure Player",
        "High strategic fit", BLUE, "#1565C0",
        "8–18x ARR",
        r * 8, r * 13, r * 18,
        "#1565C0",
        "FIS acquired Demica (Dec 2024, $300M) — Twinco is the PO finance layer Demica lacks. "
        "SAP-Taulia precedent shows infrastructure players pay premium multiples for SCF platforms "
        "with institutional relationships.",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Unattractive", "#D4D5CE", BLACK,
        "8–15x ARR",
        r * 8, r * 11, r * 15,
        BLACK,
        "C2FO valued at $1B on $186M ARR (~5x). Twinco commands premium given zero-loss track "
        "record and unique PO finance positioning.",
    )
    _val_row(
        "IPO or DFI Full Acquisition",
        "Longer term", "#D4D5CE", BLACK,
        "7–15x ARR",
        r * 7, r * 11, r * 15,
        BLACK,
        "FMO (lead Series B investor) has mandate and precedent for full acquisition of "
        "impact-aligned fintechs at scale. IPO path requires broader institutional recognition "
        "of PO finance as asset class.",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative. Primary comps: Taulia–SAP (~$400M, ~17x ARR, 2022), "
        f"Demica–FIS ($300M, ~0.75% of $40B AuA, 2024), C2FO ($1B valuation, ~5x ARR, 2019). "
        f"Twinco's zero-loss track record across $1B+ transactions and first-mover position in "
        f"securitisable PO finance support premium to comp set."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    strategic_banks = [
        ("Banco Santander", "Very High",
         "Leads Twinco's €150M securitisation facility — already the primary funding partner",
         "Natural path from securitisation funder to full owner of the PO finance platform"),
        ("BBVA", "Very High",
         "Provided €50M debt facility to Twinco in 2023 via BBVA Spark",
         "Deep existing relationship; BBVA's trade finance ambitions align directly with Twinco's EM supplier coverage"),
        ("HSBC", "High",
         "World's largest trade finance bank; active partner on Demica's platform before FIS acquisition",
         "Twinco's PO finance capability fills the pre-invoice gap in HSBC's global SCF offering"),
        ("Standard Chartered", "High",
         "Deep EM supply chain finance focus; active Demica platform partner",
         "Twinco's EM supplier network across Latin America, Asia and Africa aligns with Standard Chartered's corridor strategy"),
        ("JPMorgan", "Medium",
         "Invested in Taulia pre-SAP acquisition; major SCF platform operator at scale",
         "Twinco's PO finance layer would complement JPMorgan's existing invoice-stage SCF capabilities"),
    ]

    infrastructure_players = [
        ("FIS", "Very High",
         "Acquired Demica for $300M (Dec 2024) — explicitly stated ambition to become a leader in supply chain finance",
         "Twinco is the PO finance layer Demica does not have — acquisition would complete FIS's end-to-end SCF platform from PO to invoice"),
        ("SAP / Taulia", "High",
         "SAP acquired Taulia in 2022 for ~$400M; Taulia covers invoice and dynamic discounting stage",
         "Twinco covers the pre-invoice PO stage — zero overlap with Taulia, highly complementary acquisition to complete SAP's working capital suite"),
        ("Mastercard", "High",
         "Partnered with Demica to embed SCF before FIS acquisition; scaling B2B trade finance stack globally",
         "Twinco's EM supplier coverage and zero-loss underwriting model would strengthen Mastercard's B2B trade finance product"),
        ("Finastra", "Medium",
         "Leading trade finance software provider actively acquiring SCF capabilities",
         "Twinco's PO finance technology would extend Finastra's trade finance platform into the pre-invoice production cycle"),
        ("Network International", "Medium",
         "Pan-African and MENA payments infrastructure; acquired DPO Group in 2020",
         "Twinco's EM supplier footprint across Africa and Latin America overlaps with Network International's geographic expansion strategy"),
    ]

    secondaries_buyers = [
        ("FMO", "Very High",
         "Leads Twinco's Series B equity round; €12.1B committed portfolio across 85+ countries with impact mandate",
         "Already lead investor — natural path to full acquisition as Twinco scales; FMO has precedent for taking full ownership of impact-aligned fintechs"),
        ("IFC", "High",
         "Active in trade finance gap globally; G20 and UN recognition of SCF platforms as critical infrastructure",
         "Twinco's $1.7T trade finance gap mandate and EM supplier focus aligns directly with IFC's financial inclusion and trade development mission"),
        ("Prosus", "Medium",
         "Invested $79.9M in Mintifi (India SCF platform, 2024); building global SCF portfolio",
         "Twinco fits Prosus's pattern of backing SCF platforms in EM — geographic and product adjacency is strong"),
    ]

    affinity_cache = st.session_state.get("twinco_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="twinco_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in strategic_banks]
                + [g[0] for g in infrastructure_players]
                + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["twinco_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Twinco")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Strategic Banks", "Infrastructure Players", "Secondaries Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(strategic_banks):
            key = "engage_twinco_bank_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(infrastructure_players):
            key = "engage_twinco_infra_" + name.replace(" ", "").replace("/", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_twinco_sec_" + name.replace(" ", "").replace("/", "")
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
        "Banco Santander":     "Escalate securitisation relationship to M&A discussion. Santander already leads the €150M facility — frame as the natural ownership evolution of their PO finance infrastructure bet.",
        "BBVA":                "Re-engage via BBVA Spark. BBVA's €50M debt facility signals conviction — convert to equity conversation via investment banking intermediary.",
        "HSBC":                "Approach HSBC Trade Finance leadership. Twinco's PO finance capability fills the pre-invoice gap HSBC couldn't address via Demica — frame as completing their SCF stack.",
        "Standard Chartered":  "Engage Standard Chartered's EM trade finance team. Twinco's EM supplier network across Latin America, Asia and Africa directly aligns with their corridor strategy.",
        "JPMorgan":            "Approach via investment banking intermediary. JPMorgan's SCF platform at scale needs a PO finance layer — Twinco is the logical bolt-on.",
        "FIS":                 "Most urgent outreach — FIS acquired Demica explicitly to lead SCF. Twinco is the pre-invoice layer Demica lacks. Engage FIS M&A team immediately post-Demica integration.",
        "SAP / Taulia":        "Approach SAP Taulia's partnership team. Twinco has zero overlap with Taulia — position as completing SAP's working capital suite from PO to invoice to payment.",
        "Mastercard":          "Escalate existing partnership conversation to M&A track. Twinco's EM supplier coverage and zero-loss underwriting strengthen Mastercard's B2B trade finance product.",
        "Finastra":            "Flag for Finastra's SCF acquisition strategy. Twinco's PO finance technology extends Finastra into the pre-invoice production cycle.",
        "Network International":"Approach via Brookfield relationship. Twinco's EM supplier footprint overlaps with Network International's Africa and MENA geographic expansion.",
        "FMO":                 "FMO is already lead Series B investor — initiate acquisition conversation directly. FMO has precedent for taking full ownership of impact-aligned fintechs.",
        "IFC":                 "Approach IFC's trade finance team. Twinco's $1.7T trade finance gap mandate and EM supplier focus aligns with IFC's financial inclusion mission.",
        "Prosus":              "Flag for Prosus SCF portfolio team. Twinco fits their Mintifi pattern — geographic and product adjacency is strong for an EM SCF platform roll-up.",
    }
    _ALL_BUYERS = [b[0] for b in strategic_banks] + [b[0] for b in infrastructure_players] + [b[0] for b in secondaries_buyers]

    if st.button("Generate Exit Actions for TWINCO"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get("engage_twinco_bank_" + name.replace(" ", "").replace("/", ""), False)
            or st.session_state.get("engage_twinco_infra_" + name.replace(" ", "").replace("/", ""), False)
            or st.session_state.get("engage_twinco_sec_" + name.replace(" ", "").replace("/", ""), False)
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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

    _render_exit_deck_section(_company_id_twinco, "TWINCO")


# ── MaxSoko custom exit tab ───────────────────────────────────────────────────

_MAXSOKO_TADAWUL_ITEMS = [
    ("tadawul_cma_review",        "CMA (Capital Market Authority) eligibility review initiated"),
    ("tadawul_audited_financials", "Minimum 2 years audited financial statements (IFRS)"),
    ("tadawul_saudi_entity",       "Saudi legal entity established (registered in KSA)"),
    ("tadawul_sama",               "SAMA regulatory standing confirmed"),
    ("tadawul_ksa_revenue",        "KSA revenue exceeds SAR 50M annualised"),
    ("tadawul_profitability",      "Company profitable or clear 12-month path to profitability"),
    ("tadawul_lead_advisor",       "Lead financial advisor / investment bank appointed"),
    ("tadawul_ir",                 "Investor relations function established"),
    ("tadawul_governance",         "Corporate governance structure meets CMA standards"),
    ("tadawul_lockup",             "Lock-up and free float structure agreed with founders"),
]

_MAXSOKO_STRATEGIC_ITEMS = [
    ("strategic_ksa_revenue",      "KSA operations generating meaningful revenue (>20% of total)"),
    ("strategic_embedded_finance", "Embedded finance product live in KSA"),
    ("strategic_gcc_investors",    "At least 2 anchor GCC institutional investors on cap table"),
    ("strategic_saudi_partner",    "Saudi co-investor or strategic partner confirmed"),
    ("strategic_comp_identified",  "Comparable Saudi tech listing comp identified (Jahez used as proxy)"),
    ("strategic_saudi_director",   "Board composition includes Saudi independent director"),
    ("strategic_vision2030",       "Vision 2030 alignment narrative documented"),
    ("strategic_pre_ipo_round",    "Pre-IPO growth equity round closed"),
]


def _render_maxsoko_exit_tab() -> None:
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    EMPTY     = "#D4D5CE"

    # ── Look up company_id and LTM revenue ───────────────────────────────────
    _ms_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'MaxSoko' LIMIT 1", _conn()
    )
    _maxsoko_id = int(_ms_id_row.iloc[0]["id"]) if not _ms_id_row.empty else None
    _company_id_maxsoko = _maxsoko_id or 0
    ltm_revenue = None
    if _maxsoko_id:
        _ltm_df = load_ltm_revenue(db_version=_db_global_version())
        _vrow   = _ltm_df[_ltm_df["id"] == _maxsoko_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$50–100M"],
            "Continue scaling Egypt and KSA operations independently — limited investor liquidity without a clear exit trigger",
            [AMBER, EMPTY, EMPTY], "Unattractive strategically", False,
        ),
        (
            "Strategic Sale to FMCG Conglomerate or Regional Distributor",
            ["$80–150M", "2–4x revenue"],
            "Acquisition by a Saudi or pan-regional FMCG distributor or retailer seeking to own digital B2B distribution infrastructure",
            [AMBER, AMBER, EMPTY], "Possible — 24–36 months", False,
        ),
        (
            "Saudi IPO (Tadawul / Nomu)",
            ["$150–400M", "3–6x revenue"],
            "Founder's stated goal — list on Tadawul Main Market or Nomu parallel market following Saudi revenue scale-up and profitability milestone. Jahez listed at $2.4B on Nomu in 2022 as the first Saudi tech IPO — the benchmark for this path",
            [GREEN_DOT, AMBER, EMPTY], "Founder preferred — 3–5 years", True,
        ),
        (
            "PE Growth Equity Recap",
            ["$80–150M"],
            "Saudi or regional PE firm provides growth capital and liquidity ahead of IPO — bridge path to Tadawul listing",
            [AMBER, AMBER, EMPTY], "Bridge to IPO", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on comparable exit multiples and Saudi / MENA B2B FMCG benchmarks. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Saudi IPO (Tadawul / Nomu)",
        "Founder preferred — 3–5 years", GREEN, BLACK,
        "3–6x Revenue",
        r * 3, r * 4.5, r * 6,
        "#2E7D32",
        "Jahez listed on Nomu (2022) at $2.4B — first Saudi tech IPO. "
        "Udaan (India B2B FMCG, $1.8B valuation) trades at ~3.4x revenue after sector repricing. "
        "Saudi listings command premium to global peers given Vision 2030 tailwinds and domestic liquidity.",
    )
    _val_row(
        "Strategic Sale to FMCG Conglomerate",
        "Possible", BLUE, "#1565C0",
        "2–4x Revenue",
        r * 2, r * 3, r * 4,
        "#1565C0",
        "B2B FMCG marketplace acquisitions have repriced significantly since 2021. "
        "Asset-light models with embedded finance profit drivers command higher multiples than logistics-heavy peers.",
    )
    _val_row(
        "PE Growth Equity Recap",
        "Bridge to IPO", "#D4D5CE", BLACK,
        "2–3x Revenue",
        r * 2, r * 2.5, r * 3,
        BLACK,
        "PE recap at modest multiple to fund KSA scale-up and profitability ahead of Tadawul listing. "
        "Sanabil, STV, and ADQ/DisruptAD are the most relevant regional PE/growth equity investors.",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Unattractive", "#D4D5CE", BLACK,
        "1–2x Revenue",
        r * 1, r * 1.5, r * 2,
        BLACK,
        "Secondary at distressed multiple without strategic buyer process. "
        "Avoid unless IPO and sale processes both stall.",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative. Primary comps: Jahez IPO on Nomu ($2.4B, 2022, Saudi food delivery — "
        f"best available Saudi tech listing proxy), Udaan (India B2B FMCG, $1.8B valuation at ~3.4x revenue, "
        f"IPO prep 2026). Global B2B FMCG marketplace sector has repriced materially since 2021 — "
        f"profitability and asset-light model are now prerequisites for premium valuation."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    regional_buyers = [
        ("Bin Dawood Holding", "High",
         "Saudi grocery retail giant; listed on Tadawul 2020; actively digitising B2B supplier relationships",
         "MaxSoko's B2B FMCG distribution platform fills Bin Dawood's digital supplier network gap and provides Egypt market entry"),
        ("Savola Group", "High",
         "Saudi FMCG conglomerate operating across food manufacturing and retail in MENA; owns Panda Retail and Al Marai stake",
         "MaxSoko's Egypt and KSA B2B distribution data and retailer relationships directly complement Savola's FMCG supply chain"),
        ("Fawry", "Medium",
         "Egypt's leading digital payments platform, listed on EGX; launched Fawry Business with SME invoicing and payroll in 2025",
         "MaxSoko's embedded finance layer and retailer network complement Fawry's SME payments expansion"),
        ("LuLu Group", "Medium",
         "GCC hypermarket giant expanding digital B2B procurement and supplier digitisation",
         "MaxSoko's B2B FMCG marketplace infrastructure could power LuLu's supplier digitisation across Egypt and KSA"),
    ]

    global_strategics = [
        ("AB InBev (BEES)", "High",
         "BEES B2B platform expanded across 10+ African and MENA markets; building trade credit and embedded finance modules",
         "MaxSoko's retailer data and FMCG distribution rails complement BEES's order-to-sellout data ambitions in Egypt and KSA"),
        ("Olam International", "Medium",
         "Singapore-based agri and FMCG conglomerate with deep Africa and MENA distribution; scaling digital trade platforms",
         "MaxSoko's B2B marketplace and embedded finance layer would accelerate Olam's digital distribution ambitions in North Africa"),
        ("Udaan", "Medium",
         "India B2B FMCG marketplace, $1.8B valuation, raising $114M ahead of 2026 IPO; exploring international expansion",
         "MaxSoko is the closest comp to Udaan outside India — a merger or partnership could create a multi-market EM B2B FMCG platform pre-IPO"),
    ]

    pe_growth = [
        ("Sanabil Investments", "Very High",
         "Saudi sovereign VC/PE arm of PIF; actively backing Saudi and MENA tech growth companies pre-IPO",
         "MaxSoko's Saudi IPO ambition aligns directly with Sanabil's mandate to build Tadawul-ready tech companies"),
        ("STV (Saudi Technology Ventures)", "Very High",
         "500M MENA-focused growth tech fund; backed multiple Tadawul IPO candidates including Jahez",
         "STV has a direct playbook for backing MENA tech companies through growth stage to Tadawul listing — ideal bridge investor for MaxSoko"),
        ("ADQ / DisruptAD", "High",
         "Abu Dhabi investment arm; invested in MaxAB pre-Series B alongside British International Investment",
         "Already backed MaxSoko's closest comp (MaxAB) — natural path to back MaxSoko's Saudi expansion and IPO journey"),
        ("Algebra Ventures", "High",
         "Leading Egypt-focused VC/growth fund; backed multiple Egyptian tech exits",
         "Deep Egypt market knowledge and relationships make Algebra a credible bridge investor ahead of MaxSoko's regional scale-up"),
    ]

    affinity_cache = st.session_state.get("maxsoko_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="maxsoko_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in regional_buyers]
                + [g[0] for g in global_strategics]
                + [p[0] for p in pe_growth]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["maxsoko_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "MaxSoko")
                    for bname in all_names
                }
            st.rerun()

    tab_local, tab_global, tab_pe = st.tabs(["Regional Buyers", "Global Strategics", "PE / Growth Equity"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(regional_buyers):
            key = "engage_maxsoko_reg_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_strategics):
            key = "engage_maxsoko_glob_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_pe:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(pe_growth):
            key = "engage_maxsoko_pe_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 4: Next Steps Generator ──────────────────────────────────────
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
        "Bin Dawood Holding":         "Approach via Saudi investment banking intermediary. Frame MaxSoko's B2B FMCG distribution platform as filling Bin Dawood's digital supplier network gap and providing Egypt market entry.",
        "Savola Group":               "Engage Savola strategy team. MaxSoko's Egypt and KSA retailer relationships and distribution data directly complement Savola's FMCG supply chain.",
        "Fawry":                      "Approach via Quona's Egypt network. MaxSoko's embedded finance layer and retailer network are natural extensions of Fawry's SME payments expansion.",
        "LuLu Group":                 "Engage LuLu Group digital procurement team. MaxSoko's B2B marketplace could power LuLu's supplier digitisation across Egypt and KSA.",
        "AB InBev (BEES)":            "Approach via BEES global expansion team. MaxSoko's retailer data and FMCG distribution rails complement BEES's order-to-sellout data ambitions in Egypt and KSA.",
        "Olam International":         "Engage Olam's digital trade platforms team. MaxSoko's B2B marketplace and embedded finance layer accelerates Olam's digital distribution in North Africa.",
        "Udaan":                      "Explore strategic partnership or merger conversation. MaxSoko is the closest EM B2B FMCG comp to Udaan outside India — combined entity creates a multi-market platform pre-IPO.",
        "Sanabil Investments":        "Approach via Saudi investment banking intermediary. MaxSoko's Saudi IPO ambition aligns directly with Sanabil's mandate to build Tadawul-ready tech companies.",
        "STV (Saudi Technology Ventures)": "Re-engage STV — they backed Jahez, the primary Saudi tech IPO comp. Frame MaxSoko as the B2B FMCG equivalent on the Tadawul path.",
        "ADQ / DisruptAD":            "Approach via Abu Dhabi network. ADQ backed MaxAB (MaxSoko's closest comp) — natural path to back MaxSoko's Saudi expansion and IPO journey.",
        "Algebra Ventures":           "Engage via Quona's Egypt network. Algebra's Egypt market knowledge makes them the ideal bridge investor ahead of MaxSoko's regional scale-up.",
    }
    _ALL_BUYERS = [b[0] for b in regional_buyers] + [b[0] for b in global_strategics] + [b[0] for b in pe_growth]

    if st.button("Generate Exit Actions for MaxSoko"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get("engage_maxsoko_reg_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", ""), False)
            or st.session_state.get("engage_maxsoko_glob_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", ""), False)
            or st.session_state.get("engage_maxsoko_pe_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", ""), False)
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Section 5: Saudi IPO Readiness Tracker ───────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:700;color:{BLACK};"
        f"margin:0 0 2px 0;letter-spacing:.3px'>Saudi IPO Readiness Tracker</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Updated quarterly — tracks MaxSoko's progress against Tadawul / Nomu listing prerequisites.</div>",
        unsafe_allow_html=True,
    )

    if _maxsoko_id is None:
        st.warning("Cannot load IPO tracker: MaxSoko company record not found in database.")
        return

    # Load from DB once per session
    _ss_loaded = "maxsoko_ipo_loaded"
    _all_items = _MAXSOKO_TADAWUL_ITEMS + _MAXSOKO_STRATEGIC_ITEMS
    if not st.session_state.get(_ss_loaded):
        _ipo_db = _ipo_readiness_load(_maxsoko_id)
        for _ik, _ in _all_items:
            _d = _ipo_db.get(_ik, {})
            if f"maxsoko_ipo_{_ik}_status" not in st.session_state:
                st.session_state[f"maxsoko_ipo_{_ik}_status"] = _d.get("status", "Not Started")
            if f"maxsoko_ipo_{_ik}_notes" not in st.session_state:
                st.session_state[f"maxsoko_ipo_{_ik}_notes"]  = _d.get("notes", "")
        st.session_state["maxsoko_ipo_db_data"] = _ipo_db
        st.session_state[_ss_loaded] = True

    _ipo_db_data = st.session_state.get("maxsoko_ipo_db_data", {})

    # Progress bar
    _n_complete = sum(
        1 for _ik, _ in _all_items
        if st.session_state.get(f"maxsoko_ipo_{_ik}_status") == "Complete"
    )
    _n_total  = len(_all_items)
    _progress = _n_complete / _n_total if _n_total else 0
    st.progress(_progress, text=f"{_n_complete} / {_n_total} items complete ({_progress * 100:.0f}%)")
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    _STATUS_OPTS   = ["Not Started", "In Progress", "Complete"]
    _STATUS_COLORS = {"Not Started": MUTED, "In Progress": WARN, "Complete": "#2E7D32"}

    def _ipo_checklist(container, items, title):
        with container:
            st.markdown(
                f"<div style='font-size:11px;font-weight:700;color:{MUTED};text-transform:uppercase;"
                f"letter-spacing:.5px;margin-bottom:10px;border-bottom:1px solid {BORDER};padding-bottom:6px'>"
                f"{title}</div>",
                unsafe_allow_html=True,
            )
            for item_key, label in items:
                _ss_status = f"maxsoko_ipo_{item_key}_status"
                _ss_notes  = f"maxsoko_ipo_{item_key}_notes"
                _cur       = st.session_state.get(_ss_status, "Not Started")
                _dot_col   = _STATUS_COLORS.get(_cur, MUTED)
                _updated   = _ipo_db_data.get(item_key, {}).get("updated_at", "")

                st.markdown(
                    f"<div style='font-size:12px;color:{BLACK};margin:10px 0 3px;display:flex;"
                    f"align-items:flex-start;gap:8px'>"
                    f"<span style='display:inline-block;width:8px;height:8px;border-radius:50%;"
                    f"background:{_dot_col};margin-top:3px;flex-shrink:0'></span>"
                    f"<span>{label}</span></div>",
                    unsafe_allow_html=True,
                )
                _c1, _c2 = st.columns([1, 2])
                with _c1:
                    st.selectbox(
                        "", _STATUS_OPTS,
                        key=_ss_status,
                        label_visibility="collapsed",
                    )
                with _c2:
                    st.text_input(
                        "", key=_ss_notes,
                        placeholder="Notes…",
                        label_visibility="collapsed",
                    )
                if _updated:
                    st.markdown(
                        f"<div style='font-size:10px;color:{MUTED};margin:-4px 0 4px'>Updated: {_updated[:16]}</div>",
                        unsafe_allow_html=True,
                    )

    _col_l, _col_r = st.columns(2)
    _ipo_checklist(_col_l, _MAXSOKO_TADAWUL_ITEMS,   "Tadawul / Nomu Requirements")
    _ipo_checklist(_col_r, _MAXSOKO_STRATEGIC_ITEMS, "Strategic Readiness")

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    if st.button("Save Progress", key="maxsoko_ipo_save", type="primary"):
        _updates = {
            _ik: {
                "status": st.session_state.get(f"maxsoko_ipo_{_ik}_status", "Not Started"),
                "notes":  st.session_state.get(f"maxsoko_ipo_{_ik}_notes", ""),
            }
            for _ik, _ in _all_items
        }
        _ipo_readiness_save(_maxsoko_id, _updates)
        st.session_state.pop("maxsoko_ipo_db_data", None)
        st.session_state.pop("maxsoko_ipo_loaded", None)
        st.success("IPO readiness progress saved.")
        st.rerun()

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # Comp Set Reference panel
    st.markdown(
        f"<div style='font-size:12px;font-weight:700;color:{MUTED};text-transform:uppercase;"
        f"letter-spacing:.5px;margin-bottom:10px'>Comp Set Reference</div>",
        unsafe_allow_html=True,
    )
    _COMPS_REF = [
        ("Jahez",  "Listed Nomu January 2022", "IPO valuation $2.4B", "Revenue at listing ~$325M", "~7x revenue", "Saudi food delivery"),
        ("Udaan",  "India B2B FMCG",           "Valuation $1.8B (2024)", "Revenue $530M",          "~3.4x",       "IPO prep 2026"),
    ]
    _comp_cell = f"font-size:12px;color:{MUTED};padding:8px 14px"
    _comp_rows = ""
    for i, (co, evt, val, rev, mult, note) in enumerate(_COMPS_REF):
        _bg = "#E8F5E9" if i == 0 else "#F1F8E9"
        _comp_rows += (
            f"<div style='display:grid;grid-template-columns:0.8fr 1.2fr 1fr 1fr 0.7fr 1fr;"
            f"background:{_bg};border-radius:6px;margin-bottom:4px'>"
            f"<div style='font-size:13px;font-weight:700;color:#2E7D32;padding:8px 14px'>{co}</div>"
            f"<div style='{_comp_cell}'>{evt}</div>"
            f"<div style='{_comp_cell}'>{val}</div>"
            f"<div style='{_comp_cell}'>{rev}</div>"
            f"<div style='font-size:13px;font-weight:600;color:#2E7D32;padding:8px 14px'>{mult}</div>"
            f"<div style='{_comp_cell}'>{note}</div>"
            f"</div>"
        )
    st.markdown(
        f"<div style='background:#F9FBF7;border:1px solid #C8E6C9;border-radius:10px;padding:14px'>"
        + _comp_rows
        + f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    _render_exit_deck_section(_company_id_maxsoko, "MaxSoko")


# ── Khazna custom exit tab ────────────────────────────────────────────────────

def _render_khazna_exit_tab() -> None:
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    EMPTY     = "#D4D5CE"

    # ── Look up company_id and LTM revenue ───────────────────────────────────
    _kh_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'Khazna' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not _kh_id_row.empty:
        _kh_id  = int(_kh_id_row.iloc[0]["id"])
        _ltm_df = load_ltm_revenue(db_version=_db_global_version())
        _vrow   = _ltm_df[_ltm_df["id"] == _kh_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    # ── Section 1: Exit Pathways (collapsed) ─────────────────────────────────
    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Remain Independent — Quona Pursues Secondaries",
            ["$80–150M", "8–15x ARR"],
            "Continue scaling Egypt and KSA digital workforce banking independently — KSA SAMA license and Mudad partnership are the key value inflection points before any exit",
            [AMBER, AMBER, EMPTY], "Unattractive near-term", False,
        ),
        (
            "Full Acquisition by Wagestream",
            ["$100–200M", "10–20x ARR"],
            "Wagestream already holds a stake in Khazna and has built a global EWA portfolio including Refyne (India) and GajiGesa (Indonesia) — full acquisition is the most natural exit path",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — existing shareholder", True,
        ),
        (
            "Strategic Sale to GCC Bank or Payroll Platform",
            ["$100–250M", "10–25x ARR"],
            "Acquisition by a Saudi or UAE bank seeking to own workforce banking infrastructure — Arab National Bank and AlJazira Capital are existing investors and natural consolidators",
            [GREEN_DOT, GREEN_DOT, AMBER], "High strategic fit — 24–36 months", False,
        ),
        (
            "Strategic Sale to Global Payroll / HCM Platform",
            ["$80–180M", "8–18x ARR"],
            "Acquisition by ADP, Workday or SAP SuccessFactors to embed Khazna's EWA and workforce banking into their MENA payroll stack — Mudad integration makes Khazna highly relevant",
            [AMBER, AMBER, EMPTY], "Possible — dependent on KSA scale", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on ARR multiples given Khazna's early revenue stage. "
        f"LTM Revenue / ARR: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Full Acquisition by Wagestream",
        "Most likely", GREEN, BLACK,
        "10–20x ARR",
        r * 10, r * 15, r * 20,
        "#2E7D32",
        "Wagestream acquired stakes in Khazna (Egypt), Refyne (India) and GajiGesa (Indonesia) as part of an EM "
        "EWA consolidation strategy. Full acquisition at premium to minority stake entry is the natural next step. "
        "Payfare acquired by Fiserv at ~0.6x revenue ($147M deal) — strategic acquirers pay ARR multiples for early-stage EWA.",
    )
    _val_row(
        "Strategic Sale to GCC Bank",
        "High strategic fit", BLUE, "#1565C0",
        "10–25x ARR",
        r * 10, r * 17, r * 25,
        "#1565C0",
        "Arab National Bank and AlJazira Capital are existing investors. GCC banks pay strategic premiums for licensed "
        "workforce banking platforms given Vision 2030 workforce digitisation mandate. "
        "KSA SAMA license (expected Q2 2026) is the key valuation trigger.",
    )
    _val_row(
        "Strategic Sale to Global HCM Platform",
        "Possible", "#D4D5CE", BLACK,
        "8–18x ARR",
        r * 8, r * 13, r * 18,
        BLACK,
        "ADP, Workday and SAP SuccessFactors are all acquiring EWA capabilities. Khazna's Mudad partnership "
        "(750K employee pipeline) and MENA payroll infrastructure make it a relevant bolt-on. "
        "DailyPay valued at $1.75B on $235M revenue (~7x) sets the ceiling for mature EWA platforms.",
    )
    _val_row(
        "Remain Independent — Quona Pursues Secondaries",
        "Unattractive near-term", "#D4D5CE", BLACK,
        "8–15x ARR",
        r * 8, r * 11, r * 15,
        BLACK,
        "Secondary at modest ARR multiple ahead of KSA scale-up. "
        "MNT-Halan ($1B+ valuation at ~3x revenue) is the ceiling for what Egypt-origin digital banking can achieve "
        "— Khazna is significantly earlier stage.",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative and based on ARR multiples given Khazna's early revenue stage. "
        f"Primary comps: Payfare–Fiserv ($147M acquisition, ~0.6x revenue on $235M run-rate, 2024), "
        f"DailyPay ($1.75B valuation, ~7x revenue, 2024), MNT-Halan ($1B+ valuation, ~3x revenue, Egypt digital bank). "
        f"Wagestream existing stake is the single most important signal in the acquirer universe. "
        f"KSA SAMA license and Mudad pipeline (750K employees) are the key value inflection points."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("Arab National Bank", "Very High",
         "Existing Khazna investor; Saudi bank actively scaling digital workforce and payroll banking products",
         "Natural path from investor to acquirer — already has deep knowledge of Khazna's business and KSA expansion plans"),
        ("AlJazira Capital", "Very High",
         "Existing Khazna investor; Saudi capital markets and banking group",
         "Existing investor relationship positions AlJazira as a credible bridge to a broader Saudi bank acquisition or IPO process"),
        ("First Abu Dhabi Bank (FAB)", "High",
         "UAE's largest bank; actively acquiring fintech capabilities across MENA workforce and payroll banking",
         "Khazna's Egypt and KSA digital workforce banking platform would extend FAB's payroll banking franchise across two of MENA's largest labour markets"),
        ("Fawry", "Medium",
         "Egypt's largest listed fintech; launched Fawry Business with SME payroll and workforce payment products in 2025",
         "Khazna's payroll-backed lending and EWA capabilities would accelerate Fawry's workforce finance expansion"),
        ("Commercial International Bank (CIB)", "Medium",
         "Egypt's largest private bank; actively digitising SME and corporate payroll services",
         "Khazna's digital workforce banking platform would give CIB instant payroll-backed lending distribution across Egypt's formal sector"),
    ]

    global_buyers = [
        ("Wagestream", "Very High",
         "Already holds equity stake in Khazna; raised £300M debt facility in 2025; built global EWA portfolio including Refyne (India) and GajiGesa (Indonesia)",
         "Most likely acquirer — existing shareholder with strategic intent to consolidate EM EWA platforms into a global workforce banking group"),
        ("ADP", "High",
         "Partnered with Payfare for EWA in Canada (2024); integrating EWA into ADP's HCM platform globally; 1M+ employer relationships",
         "Khazna's MENA payroll and EWA infrastructure would extend ADP's workforce banking capabilities into Egypt and KSA — two underserved high-growth markets"),
        ("Workday", "High",
         "Partnered with DailyPay for EWA integration into Workday platform; expanding HCM coverage across MENA enterprise clients",
         "Khazna's Mudad partnership and KSA payroll infrastructure would give Workday a turnkey EWA solution for its growing Saudi enterprise client base"),
        ("Fiserv", "Medium",
         "Acquired Payfare for $147M (Dec 2024) — directly comparable EWA and digital banking platform",
         "Fiserv is actively building a global EWA portfolio post-Payfare; Khazna is the logical MENA addition to complement its gig economy workforce banking expansion"),
        ("Network International", "Medium",
         "Pan-Africa and MENA payments infrastructure; deepening Egypt and KSA fintech relationships post-acquisition by Brookfield",
         "Khazna's workforce banking and payroll-linked payments would complement Network International's merchant and corporate payments stack in MENA"),
    ]

    secondaries_buyers = [
        ("IFC", "Very High",
         "Active pan-MENA fintech investor; has been circling Khazna for several years; co-led MNT-Halan Series E (2024)",
         "IFC's financial inclusion mandate aligns directly with Khazna's unbanked workforce mission — DFI financing or secondary stake ahead of KSA scale-up"),
        ("Apis Partners", "High",
         "Led MNT-Halan Series E (2024); specialist financial services growth investor across Africa and MENA",
         "Apis has deep conviction in Egypt digital banking — Khazna is the natural complement to MNT-Halan in their MENA portfolio"),
        ("Partech", "Medium",
         "Closed €280M second Africa fund 2024; active MENA fintech investor",
         "Khazna at $9.4M ARR with clear KSA scale-up path fits Partech's growth-stage thesis"),
    ]

    affinity_cache = st.session_state.get("khazna_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="khazna_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in local_buyers]
                + [g[0] for g in global_buyers]
                + [s[0] for s in secondaries_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["khazna_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Khazna")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage Q3?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_sec = st.tabs(["Local Buyers", "Global Buyers", "Secondaries Buyers"])
    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_khazna_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_khazna_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)
    with tab_sec:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(secondaries_buyers):
            key = "engage_khazna_sec_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 4: Next Steps Generator ──────────────────────────────────────
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
        "Arab National Bank":              "Schedule bilateral with Arab National Bank strategy team. Frame as natural progression from investor to acquirer given existing deep knowledge of KSA expansion plans.",
        "AlJazira Capital":                "Engage AlJazira Capital M&A team directly. Position as bridge to a broader Saudi bank acquisition or structured IPO process leveraging their existing stake.",
        "First Abu Dhabi Bank (FAB)":      "Approach FAB via Quona MENA network. Frame as extending FAB's payroll banking franchise across Egypt and KSA — two of MENA's largest labour markets.",
        "Fawry":                           "Engage Fawry Business leadership. Frame as acquiring EWA and payroll-backed lending capability to accelerate Fawry's workforce finance expansion beyond payments.",
        "Commercial International Bank (CIB)": "Approach CIB digital banking team. Frame as instant payroll-backed lending distribution across Egypt's formal sector without building from scratch.",
        "Wagestream":                      "Escalate relationship to M&A track — existing stake makes this the most natural conversation. Approach Wagestream CEO directly via Quona connection.",
        "ADP":                             "Engage ADP HCM MENA leadership. Frame as a turnkey EWA and workforce banking solution for ADP's growing Saudi enterprise client base via Mudad partnership.",
        "Workday":                         "Approach Workday MENA strategy team. Frame as enabling Workday to offer EWA to SA enterprise clients using Khazna's Mudad-integrated KSA payroll infrastructure.",
        "Fiserv":                          "Approach Fiserv corporate development post-Payfare. Frame as the MENA addition to their global EWA portfolio — directly comparable to the Payfare acquisition logic.",
        "Network International":           "Engage Network International strategy post-Brookfield acquisition. Frame as complementing their MENA merchant payments stack with workforce banking and payroll-linked payments.",
        "IFC":                             "Approach IFC via existing relationship. Frame as financial inclusion secondary stake ahead of KSA scale-up — DFI mandate aligns directly with Khazna's unbanked workforce mission.",
        "Apis Partners":                   "Engage Apis via their Egypt digital banking conviction from MNT-Halan Series E. Frame Khazna as the natural MENA complement to Halan in their portfolio.",
        "Partech":                         "Flag for outreach at Series B. Khazna at $9.4M ARR with KSA scale-up path fits Partech's growth-stage Africa thesis — approach post SAMA licence.",
    }

    _ALL_BUYERS = (
        [b[0] for b in local_buyers]
        + [b[0] for b in global_buyers]
        + [b[0] for b in secondaries_buyers]
    )

    if st.button("Generate Exit Actions for Khazna"):
        ticked = []
        for name in _ALL_BUYERS:
            k_local = "engage_khazna_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            k_sec   = "engage_khazna_sec_" + name.replace(" ", "").replace("/", "").replace("(", "").replace(")", "")
            if st.session_state.get(k_local, False) or st.session_state.get(k_sec, False):
                ticked.append(name)
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory conversation with {name} via Quona network.")
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


# ── Enza custom exit tab ──────────────────────────────────────────────────────

def _render_enza_exit_tab() -> None:
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    RED_DOT   = "#E57373"
    EMPTY     = "#D4D5CE"

    # ── Look up company_id and LTM revenue ───────────────────────────────────
    _ez_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'Enza' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not _ez_id_row.empty:
        _ez_id  = int(_ez_id_row.iloc[0]["id"])
        _ltm_df = load_ltm_revenue(db_version=_db_global_version())
        _vrow   = _ltm_df[_ltm_df["id"] == _ez_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Strategic Sale to Global Issuer-Processor",
            ["$80–200M", "8–15x ARR"],
            "Acquisition by a global card issuing or PaaS platform seeking African bank client distribution and local scheme connectivity — the most direct fit given Enza's product.",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — 36–48 months", True,
        ),
        (
            "Acqui-hire by Global Payments Network",
            ["$50–150M"],
            "Acquisition by Mastercard, Visa, or Stripe primarily for the team, Africa bank relationships, and scheme connectivity built across Egypt, Nigeria, and South Africa.",
            [AMBER, GREEN_DOT, AMBER], "Possible — dependent on scale", False,
        ),
        (
            "Strategic Sale to Pan-African Bank",
            ["$40–100M", "5–10x ARR"],
            "Acquisition by a pan-African bank group seeking to own issuing infrastructure rather than buy it as a service.",
            [AMBER, AMBER, EMPTY], "Lower likelihood — long sales cycle", False,
        ),
        (
            "Remain Independent — Series A and Beyond",
            ["TBD at Series A valuation"],
            "Continue scaling bank and fintech client base, raise Series A, and build toward a larger exit — requires demonstrating unit economics at scale first.",
            [AMBER, AMBER, EMPTY], "Current trajectory", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on ARR multiples for card issuing PaaS in Africa and EM. "
        f"LTM Revenue / ARR: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR_V = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols_v = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols_v, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR_V}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale to Global Issuer-Processor",
        "Most likely", GREEN, BLACK,
        "8–15x ARR",
        r * 8, r * 11, r * 15,
        "#2E7D32",
        "Paymentology and Rapyd are the most credible acquirers. Comparable: Network International acquired by Brookfield at ~4x revenue (2023). "
        "Card issuing PaaS with live African bank clients and local scheme connectivity commands a strategic premium above pure revenue multiples.",
    )
    _val_row(
        "Acqui-hire by Global Payments Network",
        "Possible", BLUE, "#1565C0",
        "6–12x ARR",
        r * 6, r * 9, r * 12,
        "#1565C0",
        "Mastercard and Stripe have both paid significant premiums for Africa distribution and team quality. "
        "Stripe acquired Paystack for $200M+ with limited revenue — team, bank relationships, and scheme connectivity drive value here.",
    )
    _val_row(
        "Strategic Sale to Pan-African Bank",
        "Lower likelihood", "#D4D5CE", BLACK,
        "5–10x ARR",
        r * 5, r * 7, r * 10,
        BLACK,
        "Standard Bank and Access Bank are the most credible strategic bank buyers. "
        "Bank acquisitions of issuing infrastructure typically price below financial investor rounds — regulatory friction and longer sales cycles compress multiples.",
    )
    _val_row(
        "Remain Independent — Series A and Beyond",
        "Current trajectory", "#D4D5CE", BLACK,
        "10–20x ARR",
        r * 10, r * 15, r * 20,
        BLACK,
        "Series A round at ARR multiple consistent with African fintech comps (Paystack Series A at ~$8M ARR, Interswitch at ~6x revenue pre-Visa stake). "
        "Enza's next fundraise will be priced on ARR growth trajectory and bank client count, not exit multiples.",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative and based on ARR multiples for card issuing PaaS platforms in Africa and EM. "
        f"Primary comps: Network International acquired by Brookfield (~4x revenue, 2023), Paystack acquired by Stripe ($200M+, 2020), "
        f"Tutuka merged with Paymentology (undisclosed, 2022). "
        f"Mastercard partnership and founding team's Network International background are the strongest valuation signals."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
                f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>")

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    strategic_buyers = [
        ("Paymentology (SaltPay Group)", "Very High",
         "Dominant Africa/EM issuer-processor; actively scaling post-Tutuka merger",
         "Enza's African bank client base and local scheme connectivity (NIBSS, PayShap, InstaPay) directly fills Paymentology's distribution gap in Nigeria, South Africa, and Egypt."),
        ("Network International (Brookfield)", "Very High",
         "Enza founders' former employer; taken private by Brookfield 2023",
         "Founding team spent years building Network's Africa acceptance business — reacquiring their issuing infrastructure play is a highly logical bolt-on."),
        ("Mastercard", "High",
         "Active Enza partner; $200M MTN MoMo stake (2023)",
         "Mastercard already uses Enza as its fintech-to-card-issuance bridge in Africa. Partnership-to-acquisition is a well-worn Mastercard playbook across the continent."),
        ("Visa / Visa Direct", "High",
         "Active Africa M&A thesis; Interswitch minority stake (2019)",
         "Visa needs bank-facing issuing infrastructure in Egypt, Nigeria, and South Africa to deepen scheme penetration beyond acquirer relationships."),
        ("Stripe", "High",
         "Acquired Paystack $200M+ for Africa access (2020)",
         "Stripe's Africa stack is strong on the merchant/acquiring side via Paystack but thin on card issuing and bank-facing PaaS — Enza fills that gap directly."),
        ("Rapyd", "Medium",
         "Acquired PayU GPO $610M (2023); building global FaaS stack",
         "Rapyd's issuing capability in Africa is limited. Enza adds the bank-side infrastructure layer Rapyd needs to offer full issuing-plus-acceptance in key African markets."),
        ("Nuvei", "Medium",
         "Taken private by Advent International 2024; active EM expansion",
         "Nuvei's EM issuing footprint is thin. Enza would accelerate their Africa bank client penetration with an existing live platform."),
    ]

    bank_buyers = [
        ("Standard Bank", "Medium",
         "Largest African bank by assets; active fintech M&A",
         "Owning Enza's issuing PaaS would let Standard Bank offer white-label card infrastructure to its African correspondent bank network at scale."),
        ("Access Bank", "Medium",
         "Aggressive pan-African expansion; 18+ country footprint",
         "Access Bank's pan-African ambition needs infrastructure to match. Enza's multi-market issuing rails across Nigeria, Egypt, and South Africa align directly."),
        ("Absa Group", "Low-Medium",
         "Digitising retail and SME banking across Africa",
         "Enza's bank-facing issuing infrastructure could accelerate Absa's digital banking product rollout for partner fintechs across its African footprint."),
    ]

    affinity_cache = st.session_state.get("enza_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="enza_affinity_sync"):
            _api_key  = st.secrets.get("AFFINITY_API_KEY", "")
            all_names = list(dict.fromkeys(
                [b[0] for b in strategic_buyers]
                + [b[0] for b in bank_buyers]
            ))
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["enza_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Enza")
                    for bname in all_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols  = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_global, tab_local = st.tabs(["Global Strategic Buyers", "Pan-African Bank Buyers"])

    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(strategic_buyers):
            key = "engage_enza_" + name.replace(" ", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(bank_buyers):
            key = "engage_enza_bank_" + name.replace(" ", "")
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
        "Paymentology (SaltPay Group)":        "Approach via shared investor network. Frame as African bank distribution acquisition — Enza already operates where Paymentology is thin.",
        "Network International (Brookfield)":  "Warm intro via founding team relationship — direct outreach from Hany Fekry or Hamish Houston is the right channel.",
        "Mastercard":                          "Escalate existing partnership conversation to M&A track. Propose strategic review with Mastercard Africa leadership.",
        "Visa / Visa Direct":                  "Approach via Visa's Africa fintech investment team. Frame as bank-side issuing infrastructure to complement Visa's acquirer relationships.",
        "Stripe":                              "Approach via investment banking intermediary. Frame as the issuing-side complement to Paystack's acquiring-side Africa play.",
        "Rapyd":                               "Flag for outreach once Enza reaches 20M+ monthly transactions. Rapyd appetite for Africa issuing will grow as PayU integration settles.",
        "Standard Bank":                       "Engage via Quona board network. Frame as white-label issuing infrastructure for Standard Bank's correspondent banking partners.",
        "Access Bank":                         "Approach via Quona Africa network. Align with Access Bank's pan-African digital infrastructure buildout narrative.",
    }

    _ALL_BUYERS = [b[0] for b in strategic_buyers] + [b[0] for b in bank_buyers]

    if st.button("Generate Exit Actions for Enza"):
        ticked = [
            name for name in _ALL_BUYERS
            if st.session_state.get(
                "engage_enza_" + name.replace(" ", "").replace("(", "").replace(")", ""),
                st.session_state.get("engage_enza_bank_" + name.replace(" ", ""), False)
            )
        ]
        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory conversation with {name} via Quona network.")
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



# ── SAVA custom exit tab ──────────────────────────────────────────────────────

def _render_sava_exit_tab() -> None:
    AMBER     = "#FFC107"
    GREEN_DOT = "#D5FA94"
    RED_DOT   = "#E57373"
    EMPTY     = "#D4D5CE"

    # ── Look up company_id and LTM revenue ───────────────────────────────────
    _sv_id_row = pd.read_sql_query(
        "SELECT id FROM companies WHERE name = 'SAVA' LIMIT 1", _conn()
    )
    ltm_revenue = None
    if not _sv_id_row.empty:
        _sv_id  = int(_sv_id_row.iloc[0]["id"])
        _ltm_df = load_ltm_revenue(db_version=_db_global_version())
        _vrow   = _ltm_df[_ltm_df["id"] == _sv_id]
        if not _vrow.empty and _vrow.iloc[0]["ltm_revenue"] is not None:
            ltm_revenue = float(_vrow.iloc[0]["ltm_revenue"])

    def _pathway_card(title, valuation, description, feasibility_dots, tag, highlight=False):
        border_extra = "border-left:3px solid #D5FA94;" if highlight else ""
        dots_html = "".join(
            f"<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            f"background:{d};margin-right:3px'></span>"
            for d in feasibility_dots
        )
        rev_line = (
            f"<div style='font-size:12px;color:{MUTED};margin-top:2px'>{valuation[1]}</div>"
            if len(valuation) > 1 else ""
        )
        return f"""
<div style='background:#FFFFFF;border:1px solid #D4D5CE;{border_extra}border-radius:8px;
     padding:16px;height:100%'>
  <div style='font-size:14px;font-weight:700;color:#2C2C2A;margin-bottom:4px'>{title}</div>
  <div style='font-size:10px;color:#93A3A1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px'>Valuation</div>
  <div style='font-size:13px;color:#2C2C2A'>{valuation[0]}</div>
  {rev_line}
  <div style='font-size:12px;color:#93A3A1;font-style:italic;margin:6px 0 8px'>{description}</div>
  <div style='margin:4px 0 8px'>{dots_html}</div>
  <span style='font-size:11px;font-weight:600;color:{MUTED};background:#EFF0EA;
    border-radius:4px;padding:2px 7px'>{tag}</span>
</div>"""

    pathways = [
        (
            "Strategic Sale — Local (Bank or Telco)",
            ["$40–80M", "10–20x revenue"],
            "Acquisition by a SA bank or telco seeking the SARB banking licence as a fast-track into SME banking. Most realistic near-term path — licence scarcity creates premium above pure financial multiples.",
            [GREEN_DOT, GREEN_DOT, GREEN_DOT], "Most likely — 18–36 months", True,
        ),
        (
            "Consolidation with Yoco or Lula",
            ["$30–60M", "based on contribution"],
            "Merger with Yoco (payments) or Lula (lending) to form a full-stack SA SME challenger — Sava contributes the banking licence, spend management rails, and credit infrastructure.",
            [AMBER, GREEN_DOT, AMBER], "Possible — alignment difficult", False,
        ),
        (
            "Strategic Sale — Global (SaaS or Fintech)",
            ["$50–100M", "10–20x revenue"],
            "Acquisition by a global SaaS or fintech player using Sava as a licence-backed Africa entry point — Xero and Sage are most credible given existing SA presence and product adjacency.",
            [AMBER, AMBER, EMPTY], "Longer horizon — low near-term probability", False,
        ),
        (
            "Remain Independent — Raise Series A",
            ["$20–40M valuation"],
            "Continue scaling the SME spend management and banking platform independently, leveraging the licence for partnerships with banks and telcos while building toward a stronger exit story.",
            [AMBER, AMBER, EMPTY], "Current trajectory", False,
        ),
    ]

    with st.expander("Exit Pathways — click to expand", expanded=False):
        row1, row2 = st.columns(2), st.columns(2)
        for idx, (title, val, desc, dots, tag, highlight) in enumerate(pathways):
            col = row1[idx] if idx < 2 else row2[idx - 2]
            with col:
                st.markdown(_pathway_card(title, val, desc, dots, tag, highlight), unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 2: Implied Valuation Range ───────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 6px 0;letter-spacing:.3px'>Implied Valuation Range</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:12px;color:{MUTED};margin-bottom:16px'>"
        f"Based on revenue multiples consistent with SA fintech comps and SARB licence premium. "
        f"LTM Revenue: {fmt_usd(ltm_revenue)}</div>",
        unsafe_allow_html=True,
    )

    _HDR_V = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px"
    )
    hcols_v = st.columns([2, 1, 1, 1, 2])
    for hc, lbl in zip(hcols_v, ["Pathway", "Multiple", "Low Case", "Base Case", "High Case"]):
        with hc:
            st.markdown(f"<div style='{_HDR_V}'>{lbl}</div>", unsafe_allow_html=True)
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

    r = ltm_revenue or 0
    _val_row(
        "Strategic Sale — Local (Bank or Telco)",
        "Most likely", GREEN, BLACK,
        "10–20x revenue",
        r * 10, r * 15, r * 20,
        "#2E7D32",
        "Capitec (post-Walletdoc) and Vodacom (VodaPay) are the most credible local acquirers. "
        "SARB banking licence creates significant scarcity premium — comparable: TymeBank valued at $1.5B (2024) with SARB licence as key value driver. "
        "SA bank M&A has historically priced at 1.5–3x book value; Sava's licence and tech stack shift the framing to revenue multiples.",
    )
    _val_row(
        "Consolidation with Yoco or Lula",
        "Possible", BLUE, "#1565C0",
        "based on contribution",
        r * 8, r * 12, r * 18,
        "#1565C0",
        "Consolidation valuation would be driven by Sava's relative contribution to the combined entity. "
        "Yoco at ~$400M valuation and Lula at ~$100M (estimated) — Sava's licence adds structural value beyond revenue multiple. "
        "Comparable: Yoco acquired PayFast parent DPO for undisclosed sum to add acquiring infrastructure.",
    )
    _val_row(
        "Strategic Sale — Global (SaaS or Fintech)",
        "Longer horizon", "#D4D5CE", BLACK,
        "10–20x revenue",
        r * 10, r * 14, r * 20,
        BLACK,
        "Xero and Sage are the most credible global strategic buyers given existing SA SME presence. "
        "Xero acquired Syft Analytics in 2024 (undisclosed). Sage acquired Brightpearl at ~5x ARR. "
        "A global acquirer pays for the SA SME distribution and SARB licence, not the current revenue run-rate.",
    )
    _val_row(
        "Remain Independent — Raise Series A",
        "Current trajectory", "#D4D5CE", BLACK,
        "8–15x revenue",
        r * 8, r * 11, r * 15,
        BLACK,
        "Series A round would price on ARR growth trajectory and SARB licence optionality. "
        "SA challenger bank comps: Spot Money, Payflex at 8–12x ARR at seed/Series A. "
        "SARB licence is the primary valuation anchor — no comparable SA fintech has raised without one at this stage.",
    )

    st.markdown(
        f"<div style='background:{BG};border-radius:8px;padding:12px 16px;"
        f"font-size:11px;color:{MUTED};margin-top:8px'>"
        f"Valuation ranges are indicative and based on revenue multiples consistent with SA fintech comps. "
        f"Primary comps: TymeBank ($1.5B valuation, SARB licence, 2024), Yoco (~$400M valuation, 2022), "
        f"Capitec acquisition of Walletdoc (R400M, 2025). "
        f"SARB banking licence is the single most important value driver — scarcity premium applies above all revenue multiples."
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Section 3: Acquirer Universe ─────────────────────────────────────────
    st.markdown(
        f"<div style='font-size:13px;font-weight:500;color:{MUTED};"
        f"margin:20px 0 12px 0;letter-spacing:.3px'>Acquirer Universe — Prioritized</div>",
        unsafe_allow_html=True,
    )

    FIT_COLORS = {
        "Very High":  ("#D5FA94", "#2C2C2A"),
        "High":       ("#C5E5FF", "#1565C0"),
        "Medium":     ("#D4D5CE", "#2C2C2A"),
        "Low-Medium": ("#FFCDD2", "#B71C1C"),
        "Low":        ("#FFCDD2", "#B71C1C"),
    }

    def _fit_badge(fit):
        bg, fg = FIT_COLORS.get(fit, ("#D4D5CE", "#2C2C2A"))
        return (
            f"<span style='background:{bg};color:{fg};font-size:11px;font-weight:600;"
            f"border-radius:4px;padding:2px 7px;margin-left:6px'>{fit}</span>"
        )

    def _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=0, affinity_override=None):
        row_bg = "#EFF0EA" if row_idx % 2 == 0 else "#FFFFFF"
        with st.container():
            st.markdown(
                f"<div style='background:{row_bg};border-radius:6px;padding:6px 4px 2px'>",
                unsafe_allow_html=True,
            )
            cols = st.columns([2, 2, 3, 1, 3])
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
                if affinity_override is not None:
                    st.markdown(
                        f"<div style='font-size:12px;color:{MUTED};padding-top:8px'>{affinity_override}</div>",
                        unsafe_allow_html=True,
                    )
                elif affinity_cache is None:
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
                        _snip = note.get("snippet") or ""
                        _snip_style = "font-style:italic;" if _snip == "No exit-relevant discussion found." else ""
                        st.markdown(
                            f"<div style='font-size:12px;color:#2E7D32;font-weight:600;padding-top:4px'>{note['date']}</div>"
                            f"<div style='font-size:11px;color:{MUTED};{_snip_style}'>{_snip}</div>",
                            unsafe_allow_html=True,
                        )
            st.markdown("</div>", unsafe_allow_html=True)

    local_buyers = [
        ("Capitec", "Very High",
         "Acquired Walletdoc R400M (Dec 2025); partnered with Stub for SME accounting",
         "Sava's SARB licence and SME spend management rails complete Capitec's business banking buildout — they have payments (Walletdoc) but lack a licensed SME banking and credit infrastructure layer."),
        ("Vodacom", "Very High",
         "R34B fintech revenue FY2025; VodaPay SME strategy active",
         "Sava's banking licence is worth more to Vodacom than to any SA bank — it gives them a regulated SME banking shortcut that would otherwise take years and hundreds of millions to build organically."),
        ("Old Mutual (OM Bank)", "High",
         "OM Bank launched 2025; targeting 2.5–3M customers by 2028",
         "OM Bank is building fast but lacks SME-specific spend management and BaaS rails. Sava's licence and platform would accelerate their business banking expansion ahead of the 2028 profitability target."),
        ("MTN", "High",
         "MoMo SA traction weak vs other markets; SME banking gap clear",
         "MTN's SA fintech story is underdeveloped relative to its other markets. Sava's licence and SME-first platform gives MTN a credible SA business banking entry point alongside MoMo consumer services."),
        ("FNB / FirstRand", "Medium",
         "Scaling SME payments and digital business services",
         "FNB already serves a large SA SME base but has limited spend management and challenger brand capability. Sava's tech stack could run as an FNB white-label or challenger brand for digitally-native SMEs."),
        ("TymeBank", "Medium",
         "Acquired Retail Capital (2022); unicorn status ($1.5B) Dec 2024",
         "Tyme is building the full SME financial OS — credit (Retail Capital) plus banking. Sava's spend management and SARB licence adds the missing expense management and regulatory depth layer."),
        ("Absa", "Medium",
         "Digitising retail and SME banking across Africa",
         "Absa's SME banking product is lagging Capitec and Tyme. Sava's challenger brand positioning and tech rails could give Absa a faster path to SME-native banking without rebuilding from scratch."),
    ]

    global_buyers = [
        ("Xero", "High",
         "Acquired Cape Town-based Syft Analytics (2024)",
         "Xero already acquires SA fintechs and integrates with Yoco. Adding Sava's banking licence and spend management layer would let Xero offer SA SMEs a complete accounting-plus-banking stack — a natural product extension."),
        ("Sage", "High",
         "Dominant SA SME software player; adding payments and banking globally",
         "Sage has deep SA SME roots and is building out financial services beyond accounting. Sava gives Sage a licence-backed route to offer banking, payments, and credit alongside its existing SA software suite."),
        ("Stripe", "Low-Medium",
         "Africa presence via Paystack (Nigeria); limited SA footprint",
         "Stripe's SA story is thin. Sava's SARB licence and SME banking rails could accelerate their SA expansion, but appetite for a second Africa market-entry acquisition remains uncertain."),
        ("Adyen", "Low-Medium",
         "Scaling enterprise acquiring globally; limited SA SME focus",
         "Adyen's SA operations are enterprise-focused. Sava would require a meaningful strategic pivot toward SME banking — possible but not a near-term priority."),
    ]

    consolidation_buyers = [
        ("Yoco", "High",
         "Scaling Yoco Capital; approaching EBITDA breakeven",
         "Sava contributes the SARB banking licence and spend management rails that Yoco needs to complete its neobank buildout — combined entity becomes SA's first full-stack SME financial OS."),
        ("Lulalend (Lula)", "High",
         "Strategic process active; Series C or trade sale being explored",
         "Lula's credit platform plus Sava's banking licence and spend management creates a complete SME bank. Logic is strong but alignment remains difficult given each company's independent investor timelines."),
    ]

    all_buyer_names = (
        [b[0] for b in local_buyers]
        + [b[0] for b in global_buyers]
        + [b[0] for b in consolidation_buyers]
    )

    affinity_cache = st.session_state.get("sava_affinity_data")
    _, _sync_btn_col = st.columns([6, 1])
    with _sync_btn_col:
        if st.button("Sync Affinity", key="sava_affinity_sync"):
            _api_key = st.secrets.get("AFFINITY_API_KEY", "")
            with st.spinner("Fetching Affinity data for all buyers…"):
                st.session_state["sava_affinity_data"] = {
                    bname: fetch_last_affinity_note_for_buyer(bname, _api_key, "Sava")
                    for bname in all_buyer_names
                }
            st.rerun()

    _HDR_STYLE = (
        f"font-size:10px;font-weight:700;color:#93A3A1;"
        f"text-transform:uppercase;letter-spacing:.5px;padding-bottom:4px"
    )

    def _header_row():
        hcols = st.columns([2, 2, 3, 1, 3])
        labels = ["Buyer / Fit", "Recent Activity", "Strategic Rationale", "Re-engage?", "Last Affinity Contact"]
        for hc, lbl in zip(hcols, labels):
            with hc:
                st.markdown(f"<div style='{_HDR_STYLE}'>{lbl}</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:2px;background:#EFF0EA;margin-bottom:8px'></div>", unsafe_allow_html=True)

    tab_local, tab_global, tab_consolidation = st.tabs(["Local Buyers", "Global Buyers", "Consolidation"])

    with tab_local:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(local_buyers):
            key = "engage_sava_local_" + name.replace(" ", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    with tab_global:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(global_buyers):
            key = "engage_sava_global_" + name.replace(" ", "").replace("(", "").replace(")", "")
            _buyer_row(name, fit, activity, rationale, key, affinity_cache, row_idx=idx)

    with tab_consolidation:
        _header_row()
        for idx, (name, fit, activity, rationale) in enumerate(consolidation_buyers):
            key = "engage_sava_consol_" + name.replace(" ", "").replace("(", "").replace(")", "")
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
        "Capitec":              "Request meeting with Capitec Business Banking leadership. Frame Sava as the SME spend management and licensed banking layer that completes their post-Walletdoc buildout.",
        "Vodacom":              "Approach via VodaPay strategy team. Position Sava's SARB licence as a regulated SME banking shortcut that accelerates Vodacom's SA business banking ambitions by 3–5 years.",
        "Old Mutual (OM Bank)": "Engage OM Bank CEO and strategy team directly. Frame as accelerating their 2028 profitability target by adding SME business banking without building from scratch.",
        "MTN":                  "Approach via MTN SA Fintech team. Position as the missing SA SME banking chapter in MTN's African fintech story — complements MoMo consumer services with a regulated business offering.",
        "FNB / FirstRand":      "Engage via Quona's FirstRand network. Frame as a challenger brand play — Sava runs as a separate SME-native brand within FirstRand's ecosystem.",
        "TymeBank":             "Approach TymeBank CEO directly. Frame as the spend management and SARB licence layer that completes the Tyme SME financial OS post-Retail Capital.",
        "Absa":                 "Engage Absa Digital and Strategy team. Frame as a faster path to SME challenger banking than internal build given Capitec and Tyme pressure.",
        "Xero":                 "Warm intro via Quona or Syft relationship — Xero is already acquiring SA fintechs. Frame as adding banking and spend management to Xero's SA accounting stack.",
        "Sage":                 "Engage Sage SA leadership. Frame as completing Sage's SA SME financial OS — accounting, payroll, banking, and credit from one platform.",
        "Yoco":                 "Initiate consolidation conversation via board. Frame as the fastest path to a neobank narrative — Sava's licence removes Yoco's biggest structural gap.",
        "Lulalend (Lula)":      "Engage via Quona board relationship. Frame as creating SA's first full-stack SME bank — Lula credit plus Sava licence and spend management.",
    }

    if st.button("Generate Exit Actions for SAVA"):
        ticked = []
        for name in all_buyer_names:
            for prefix in ["engage_sava_local_", "engage_sava_global_", "engage_sava_consol_"]:
                key = prefix + name.replace(" ", "").replace("(", "").replace(")", "")
                if st.session_state.get(key, False):
                    ticked.append(name)
                    break

        st.markdown("#### Strategic Acquisition Outreach")
        if ticked:
            for name in ticked:
                action = _BUYER_ACTIONS.get(name, f"Schedule introductory strategic conversation with {name} via Quona network.")
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


# ── Exit Tracking tab ─────────────────────────────────────────────────────────

def render_exit_tab(info: pd.Series, company_id: int) -> None:
    company_name = info["name"]
    sector       = str(info.get("sector", "")).lower()
    _today       = datetime.utcnow()
    cur_q        = f"Q{(_today.month - 1) // 3 + 1} {_today.year}"

    # ── Enza custom exit tab ───────────────────────────────────────────────────
    if company_name == "Enza":
        return _render_enza_exit_tab()

    # ── SAVA custom exit tab ────────────────────────────────────────────────────
    if company_name == "SAVA":
        return _render_sava_exit_tab()

    # ── Cowrywise custom exit tab ──────────────────────────────────────────────
    if company_name == "Cowrywise":
        return _render_cowrywise_exit_tab()

    # ── Yoco custom exit tab ───────────────────────────────────────────────────
    if company_name == "Yoco":
        return _render_yoco_exit_tab()

    # ── Lulalend custom exit tab ───────────────────────────────────────────────
    if company_name == "Lulalend":
        return _render_lulalend_exit_tab()

    # ── VertoFX custom exit tab ────────────────────────────────────────────────
    if company_name in ("VertoFX", "Verto FX"):
        return _render_vertofx_exit_tab()

    # ── TWINCO custom exit tab ─────────────────────────────────────────────────
    if company_name in ("TWINCO", "Twinco"):
        return _render_twinco_exit_tab()

    # ── MaxSoko custom exit tab ────────────────────────────────────────────────
    if company_name == "MaxSoko":
        return _render_maxsoko_exit_tab()

    # ── Khazna custom exit tab ─────────────────────────────────────────────────
    if company_name == "Khazna":
        return _render_khazna_exit_tab()

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

_warm_cache()

# ── Authentication gate ───────────────────────────────────────────────────────
def _render_login() -> None:
    st.markdown("""
        <style>
        #MainMenu, footer, header {visibility: hidden;}
        [data-testid="stToolbar"] {display: none;}
        </style>
    """, unsafe_allow_html=True)
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)
        st.markdown("""
        <div style='background:#2C2C2A;border-radius:16px;padding:48px;
                    text-align:center;margin-bottom:32px'>
          <div style='font-size:32px;font-weight:800;color:#D5FA94;
                      letter-spacing:-0.5px;margin-bottom:8px'>Quona Capital</div>
          <div style='font-size:15px;color:rgba(255,255,255,0.55)'>
                      Portfolio Intelligence Platform</div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login"):
            password = st.text_input("Access code", type="password", placeholder="Enter access code")
            submit   = st.form_submit_button("Enter", use_container_width=True)
        _stored = (
            st.secrets.get("APP_PASSWORD") or
            st.secrets.get("auth", {}).get("APP_PASSWORD", "")
        )
        if submit:
            if password == _stored:
                st.session_state["_auth"] = True
                st.rerun()
            else:
                st.error("Incorrect access code.")
        st.markdown(
            "<div style='text-align:center;font-size:12px;color:#93A3A1;"
            "margin-top:16px'>Restricted to Quona Capital team</div>",
            unsafe_allow_html=True,
        )

if not st.session_state.get("_auth"):
    _render_login()
    st.stop()

# ── Persistent header ─────────────────────────────────────────────────────────
_hdr_left, _hdr_right = st.columns([9, 1])
with _hdr_left:
    st.markdown(f"""
    <div style="background:{BLACK};border-radius:12px;padding:14px 28px;
                margin-bottom:20px;display:flex;align-items:baseline;gap:12px;">
      <span style="font-size:22px;font-weight:800;color:{GREEN};
                   letter-spacing:-0.5px;">Quona Capital</span>
      <span style="font-size:13px;color:rgba(255,255,255,0.55);">
                   Portfolio Intelligence</span>
    </div>
    """, unsafe_allow_html=True)
with _hdr_right:
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    if st.button("Sign out", key="header_logout"):
        st.session_state["_auth"] = False
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# HOME PAGE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "home":

    _gver     = _db_global_version()
    companies = load_companies(db_version=_gver)
    growth    = load_revenue_growth(db_version=_gver)
    ltm       = load_ltm_revenue(db_version=_gver)
    all_rev   = load_all_revenue(db_version=_gver)
    vol       = load_ltm_volume(db_version=_gver)

    companies = companies.merge(growth, on="id", how="left")
    companies = companies.merge(ltm,    on="id", how="left")
    companies = companies.merge(vol,    on="id", how="left")

    flags = compute_data_quality_flags(companies, ltm, all_rev)

    # ── Summary KPIs ──────────────────────────────────────────────────────────
    ltm_gm_col = companies["ltm_gross_margin_pct"].combine_first(
        companies["gross_margin_pct"]
    )
    ltm_em_col = companies["ltm_ebitda_margin_pct"].combine_first(
        companies["ebitda_margin_pct"]
    )
    n_companies  = len(companies)
    combined_rev = fmt_usd(companies["ltm_revenue"].sum())
    avg_gm       = fmt_pct(ltm_gm_col.mean())
    _avg_em_num  = ltm_em_col.mean()
    avg_em       = fmt_pct(_avg_em_num)
    _em_color    = "#2E7D32" if (not _is_null(_avg_em_num) and float(_avg_em_num) > 0) else "#C62828"
    st.markdown(f"""
<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:0;
            background:white;border:1px solid #D4D5CE;border-radius:10px;
            margin-bottom:20px;overflow:hidden">
  <div style="padding:16px 24px;border-right:1px solid #D4D5CE">
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;
                text-transform:uppercase;color:#93A3A1;margin-bottom:4px">Portfolio Companies</div>
    <div style="font-size:28px;font-weight:800;color:#2C2C2A">{n_companies}</div>
  </div>
  <div style="padding:16px 24px;border-right:1px solid #D4D5CE">
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;
                text-transform:uppercase;color:#93A3A1;margin-bottom:4px">Combined LTM Revenue</div>
    <div style="font-size:28px;font-weight:800;color:#2C2C2A">{combined_rev}</div>
  </div>
  <div style="padding:16px 24px;border-right:1px solid #D4D5CE">
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;
                text-transform:uppercase;color:#93A3A1;margin-bottom:4px">Avg Gross Margin</div>
    <div style="font-size:28px;font-weight:800;color:#2C2C2A">{avg_gm}</div>
  </div>
  <div style="padding:16px 24px">
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;
                text-transform:uppercase;color:#93A3A1;margin-bottom:4px">Avg EBITDA Margin</div>
    <div style="font-size:28px;font-weight:800;color:{_em_color}">{avg_em}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Filter bar ────────────────────────────────────────────────────────────
    all_sectors = sorted(companies["sector"].dropna().unique().tolist())
    sector_options = ["All"] + [sector_label(s) for s in all_sectors if s and sector_label(s) and sector_label(s).lower() != "none"]

    _fund_options = ["All Funds"] + sorted([f for f in companies["fund"].dropna().unique() if f and str(f).lower() != "none"])
    st.markdown("<div style='font-size:11px;font-weight:700;color:#93A3A1;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px'>Fund</div>", unsafe_allow_html=True)
    selected_fund = st.radio(
        "Filter by fund",
        options=_fund_options,
        index=0,
        horizontal=True,
        label_visibility="collapsed",
    )

    filter_col, sort_col = st.columns([4, 1])
    with filter_col:
        st.markdown("<div style='font-size:11px;font-weight:700;color:#93A3A1;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px'>Sector</div>", unsafe_allow_html=True)
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

    EXIT_READY_PRIORITY = ["Yoco", "Cowrywise", "Lulalend", "VertoFX", "MaxSoko"]

    filtered = filtered.copy()
    filtered["_sort_primary"] = filtered["name"].apply(lambda n: 0 if n in EXIT_READY_PRIORITY else 1)
    filtered["_sort_secondary"] = filtered["name"].apply(
        lambda n: EXIT_READY_PRIORITY.index(n) if n in EXIT_READY_PRIORITY else 999
    )
    filtered = filtered.sort_values(["_sort_primary", "_sort_secondary", "name"]).drop(
        columns=["_sort_primary", "_sort_secondary"]
    )

    n_showing = len(filtered)
    st.markdown(
        f"<div style='font-size:13px;font-weight:600;color:#2C2C2A;letter-spacing:.04em;"
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
            ltm_tpv = row.get("ltm_tpv_usd")
            ltm_gmv = row.get("ltm_gmv_usd")
            if not _is_null(ltm_tpv):
                tpv, tpv_lbl = ltm_tpv, "ANNUAL TPV"
            elif not _is_null(ltm_gmv):
                tpv, tpv_lbl = ltm_gmv, "ANNUAL GMV"
            else:
                tpv, tpv_lbl = None, "ANNUAL TPV"
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
        country    = row.get("hq_country") or ""
        ltm_val    = row.get("ltm_revenue")
        ltm_lbl    = row.get("ltm_label") or ""
        pt         = row.get("period_type", "monthly")
        period_lbl = fmt_period_label(row.get("period_end_date"), pt)
        asof       = as_of(row.get("period_end_date"))

        # LTM revenue
        rev_str = fmt_usd(ltm_val)
        if ltm_lbl == "LTM":
            rev_sub   = ""
            rev_label = "LTM REVENUE"
        elif ltm_lbl == "ARR (est.)":
            rev_sub   = "ARR · est."
            rev_label = "ARR REVENUE (EST.)"
        else:
            rev_sub   = ""
            rev_label = "REVENUE"
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

        SECTOR_BORDER = {
            "payments":         "#D5FA94",
            "lending":          "#C5E5FF",
            "wealth_management":"#C5E5FF",
            "insurtech":        "#F5C36A",
        }
        sector_key   = str(row.get("sector", "")).lower()
        sector_color = SECTOR_BORDER.get(sector_key, "#D4D5CE")

        _MLBL = "font-size:11px;font-weight:600;color:#2C2C2A;letter-spacing:.06em;text-transform:uppercase;margin-bottom:2px"
        _MVAL = "font-size:18px;font-weight:800"
        _MSUB = f"font-size:10px;color:{MUTED};margin-top:1px"
        _MPAD = "padding:8px 0"

        with col:
            with st.container(border=True):
                # ── Sector color bar ──────────────────────────────────────
                st.markdown(
                    f"<div style='height:4px;background:{sector_color};"
                    f"border-radius:4px 4px 0 0;margin:-8px -8px 12px -8px'></div>",
                    unsafe_allow_html=True,
                )

                # ── Header: name + fund tag + sector tag ─────────────────
                fund_val  = str(row.get("fund", "") or "")
                FUND_COLORS = {
                    "Fund I":   ("#D5FA94", "#2C2C2A"),
                    "Fund II":  ("#C5E5FF", "#2C2C2A"),
                    "Fund III": ("#EFF0EA", "#93A3A1"),
                }
                fund_bg, fund_fg = FUND_COLORS.get(fund_val, (None, None))
                fund_tag_html = (
                    f"<span style='background:{fund_bg};color:{fund_fg};border-radius:99px;"
                    f"padding:3px 10px;font-size:10px;font-weight:700;letter-spacing:.04em;"
                    f"white-space:nowrap;display:inline-block;margin-top:4px'>{fund_val}</span>"
                    if fund_bg else ""
                )

                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;"
                    f"align-items:flex-start;margin-bottom:2px'>"
                    f"<div>"
                    f"<div style='font-size:19px;font-weight:800;color:{BLACK};"
                    f"letter-spacing:-0.3px;line-height:1.2'>{name}</div>"
                    f"<div style='font-size:12px;font-weight:600;color:{BLACK};"
                    f"text-transform:uppercase;letter-spacing:.06em;margin-top:2px'>"
                    f"{country}</div>"
                    f"{fund_tag_html}"
                    f"</div>"
                    f"<span style='background:{BLUE};color:{BLACK};border-radius:99px;"
                    f"padding:4px 12px;font-size:11px;font-weight:700;"
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
                        f"<div style='{_MPAD}'>"
                        f"<div style='{_MLBL}'>{rev_label}</div>"
                        f"<div style='{_MVAL};color:{BLACK}'>{rev_str}</div>"
                        f"<div style='{_MSUB}'>{rev_sub}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                with m2:
                    st.markdown(
                        f"<div style='{_MPAD}'>"
                        f"<div style='{_MLBL}'>Gross Margin</div>"
                        f"<div style='{_MVAL};color:{gm_color}'>{gm_str}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                with m3:
                    st.markdown(
                        f"<div style='{_MPAD}'>"
                        f"<div style='{_MLBL}'>Rev Growth</div>"
                        f"<div style='{_MVAL};color:{gcol}'>{gtxt}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                with m4:
                    st.markdown(
                        f"<div style='{_MPAD}'>"
                        f"<div style='{_MLBL}'>{lbl3}</div>"
                        f"<div style='{_MVAL};color:{col3}'>{val3}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )

                st.markdown(f"<hr style='margin:8px 0;border-color:{BORDER}'>",
                            unsafe_allow_html=True)

                # ── Footer: as of date + view button ─────────────────────
                st.markdown(
                    f"<div style='font-size:11px;color:#93A3A1;font-weight:600;"
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
    _cver   = _kpi_db_version(company_id)
    _gver   = _db_global_version()
    print(f"[company_detail] DB={DB_PATH} company_id={company_id} cver={_cver!r}")
    info = load_company_info(company_id, db_version=_cver)
    kpis = load_kpis(company_id, db_version=_cver)

    # LTM for this company
    ltm_df  = load_ltm_revenue(db_version=_gver)
    ltm_row = ltm_df[ltm_df["id"] == company_id]
    ltm_val = float(ltm_row.iloc[0]["ltm_revenue"]) if not ltm_row.empty and not _is_null(ltm_row.iloc[0]["ltm_revenue"]) else None
    _raw_ltm_lbl = ltm_row.iloc[0].get("ltm_label") if not ltm_row.empty else None
    ltm_lbl = str(_raw_ltm_lbl) if not _is_null(_raw_ltm_lbl) and str(_raw_ltm_lbl).lower() not in ("none", "") else "—"

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
            <span style="color:{MUTED};font-size:13px">{info.get('hq_country') or ''} {founded}</span>
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

    try:
        _valid_dates = kpis["period_end_date"].dropna()
        date_range = (
            f"{_valid_dates.min().strftime('%b %Y')} – "
            f"{_valid_dates.max().strftime('%b %Y')}"
        ) if len(_valid_dates) >= 1 else "—"
    except Exception:
        date_range = "—"

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
    _tab_names  = [t for t in ["Performance", "Benchmarking", "Exit Tracking",
                                "Upload Data" if _has_upload else None] if t is not None]
    _tabs       = st.tabs(_tab_names)
    tab_perf    = _tabs[0]
    tab_bench   = _tabs[1]
    tab_exit    = _tabs[2]
    tab_upload  = _tabs[3] if _has_upload else None

    with tab_perf:
        # ── Last-updated stamp (reads directly from DB, not cache) ─────────────
        _last_ts = _kpi_last_updated(company_id)
        st.markdown(
            f"<div style='text-align:right;font-size:11px;color:{MUTED};"
            f"margin-bottom:4px'>Data last updated: <b>{_last_ts}</b></div>",
            unsafe_allow_html=True,
        )

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
                    col.metric(lbl, val_str if val_str is not None else "—")

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
