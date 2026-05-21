"""
Seed portfolio company data into the benchmarking database.

Step 1 — inserts one row per company into `companies`.
Step 2 — reads each company's Excel file and inserts KPI rows into `kpi_snapshots`.

Run:
    python scripts/seed_portfolio.py

Each company has its own reader function because every Excel file has a
different layout.  Files that cannot be read are skipped with a warning.
MaxSoko's Summary sheet contains only cross-sheet formula references whose
cached values are absent; it is skipped automatically.
"""

import calendar
import re
import sqlite3
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")

ROOT      = Path(__file__).parent.parent
DB_PATH   = ROOT / "benchmarking.db"
DATA_DIR  = ROOT / "data"

# Fixed FX rates  (quote convention: ZAR/NGN → divide; GBP/EUR → multiply)
FX = {"ZAR": 18.5, "NGN": 1500.0, "GBP": 1.27, "EUR": 1.08, "USD": 1.0}

# ─── helpers ─────────────────────────────────────────────────────────────────

def to_float(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "").replace(" ", ""))
    except (ValueError, TypeError):
        return None

def to_int(v):
    f = to_float(v)
    return int(round(f)) if f is not None else None

def pct(v):
    """Convert a decimal fraction (0.xx) to a percentage (xx.x)."""
    f = to_float(v)
    return round(f * 100, 4) if f is not None else None

def to_usd(value, currency):
    """Convert a local-currency amount to USD using the fixed rates."""
    f = to_float(value)
    if f is None:
        return None
    if currency in ("ZAR", "NGN"):
        return f / FX[currency]
    if currency in ("GBP", "EUR"):
        return f * FX[currency]
    return f  # USD

def month_end_from_dt(dt):
    """Return ISO 8601 month-end string from a date or datetime."""
    if isinstance(dt, (datetime, date)):
        y, m = dt.year, dt.month
        last = calendar.monthrange(y, m)[1]
        return date(y, m, last).isoformat()
    return None

def month_end_from_str(s):
    """Parse text dates like 'Feb 2017', 'January 2020', \"Apr'24\"."""
    if not isinstance(s, str):
        return None
    s = s.strip()
    m = re.match(r"(\w{3,9})'(\d{2})$", s)
    if m:
        try:
            dt = datetime.strptime(f"{m.group(1)} 20{m.group(2)}", "%b %Y")
            last = calendar.monthrange(dt.year, dt.month)[1]
            return date(dt.year, dt.month, last).isoformat()
        except ValueError:
            pass
    for fmt in ("%B  %Y", "%B %Y", "%b %Y"):
        try:
            dt = datetime.strptime(s, fmt)
            last = calendar.monthrange(dt.year, dt.month)[1]
            return date(dt.year, dt.month, last).isoformat()
        except ValueError:
            pass
    return None

def year_end(v):
    """Convert a year value (int/float/str) to Dec 31 ISO string."""
    try:
        return date(int(float(str(v))), 12, 31).isoformat()
    except (ValueError, TypeError):
        return None

def load_rows(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = {i + 1: list(row) for i, row in enumerate(ws.iter_rows(values_only=True))}
    wb.close()
    return rows

def upsert_kpi(conn, company_id, period, reporting_currency, **kwargs):
    if not period:
        return
    # Skip rows where every metric value is None or zero — template/budget artefacts
    metric_vals = [v for k, v in kwargs.items() if k not in ("fx_rate_to_usd",)]
    if not any(v not in (None, 0, 0.0) for v in metric_vals):
        return
    cols = ["company_id", "period_end_date", "reporting_currency"] + list(kwargs.keys())
    vals = [company_id, period, reporting_currency] + list(kwargs.values())
    ph   = ",".join("?" * len(cols))
    upd  = ",".join(f"{k}=excluded.{k}" for k in kwargs)
    sql  = (
        f"INSERT INTO kpi_snapshots ({','.join(cols)}) VALUES ({ph})"
        + (f" ON CONFLICT(company_id, period_end_date) DO UPDATE SET {upd}" if upd else
           " ON CONFLICT(company_id, period_end_date) DO NOTHING")
    )
    conn.execute(sql, vals)

def get_id(conn, name):
    row = conn.execute("SELECT id FROM companies WHERE name=?", (name,)).fetchone()
    if not row:
        raise ValueError(f"Company not found: {name}")
    return row[0]

# ─── Step 1: companies ───────────────────────────────────────────────────────

COMPANIES = [
    dict(name="Cowrywise",  sector="wealth_management", sub_sector="savings_and_investment",
         hq_country="NG", business_model="b2c",  reporting_currency="NGN", founded_year=2017),
    dict(name="Yoco",       sector="payments",          sub_sector="merchant_acquiring",
         hq_country="ZA", business_model="b2b",  reporting_currency="ZAR", founded_year=2015),
    dict(name="Verto",      sector="payments",          sub_sector="cross_border_fx",
         hq_country="NG", business_model="b2b",  reporting_currency="USD", founded_year=2019),
    dict(name="Enza",       sector="payments",          sub_sector="card_issuing_paas",
         hq_country="KE", business_model="b2b",  reporting_currency="USD", founded_year=2020),
    dict(name="Lulalend",   sector="lending",           sub_sector="sme_lending",
         hq_country="ZA", business_model="b2b",  reporting_currency="ZAR", founded_year=2014),
    dict(name="Khazna",     sector="lending",           sub_sector="consumer_lending",
         hq_country="EG", business_model="b2c",  reporting_currency="USD", founded_year=2021),
    dict(name="TWINCO",     sector="lending",           sub_sector="supply_chain_finance",
         hq_country="ES", business_model="b2b",  reporting_currency="EUR", founded_year=2019),
    dict(name="MaxSoko",    sector="marketplace",       sub_sector="ecommerce_embedded_finance",
         hq_country="EG", business_model="b2b",  reporting_currency="USD", founded_year=2015),
    dict(name="SAVA",       sector="payments",          sub_sector="card_issuing_baas",
         hq_country="ZA", business_model="b2b",  reporting_currency="ZAR", founded_year=2022),
    dict(name="AllLife",    sector="insurtech",         sub_sector="life_insurance",
         hq_country="ZA", business_model="b2c",  reporting_currency="ZAR", founded_year=2004),
    dict(name="OCTA",       sector="saas",              sub_sector="invoice_ar_automation",
         hq_country="NG", business_model="b2b",  reporting_currency="USD", founded_year=2023),
    dict(name="Eseye",      sector="iot_infrastructure",sub_sector="managed_connectivity",
         hq_country="GB", business_model="b2b",  reporting_currency="GBP", founded_year=2007),
    dict(name="POWER",      sector="lending",           sub_sector="earned_wage_access",
         hq_country="US", business_model="b2b",  reporting_currency="USD", founded_year=2020),
]

def seed_companies(conn):
    inserted = 0
    for c in COMPANIES:
        conn.execute("""
            INSERT OR IGNORE INTO companies
                (name, type, sector, sub_sector, hq_country, founded_year,
                 business_model, reporting_currency)
            VALUES (?,?,?,?,?,?,?,?)
        """, (c["name"], "portfolio", c["sector"], c["sub_sector"],
              c["hq_country"], c["founded_year"],
              c["business_model"], c["reporting_currency"]))
        inserted += conn.execute(
            "SELECT changes()"
        ).fetchone()[0]
    conn.commit()
    print(f"  companies: {inserted} row(s) inserted")

# ─── Step 2: KPI readers ─────────────────────────────────────────────────────

def seed_cowrywise(conn):
    """Monthly, Jan 2020–Mar 2026. USD revenue/EBITDA already provided."""
    cid  = get_id(conn, "Cowrywise")
    rows = load_rows(DATA_DIR / "Cowrywise KPIs_Mar 2026.xlsx", "Cowrywise")
    # Row 6: date headers (col E = idx 4 onwards)
    # Row 15: AUM in USD; Row 48/49: Revenue/EBITDA USD
    # Row 50/51: Gross/EBITDA margin as decimals
    # Row 56/57: Total/Active customers; Row 59: Churn rate decimal
    date_r  = rows.get(6,  [])
    aum_r   = rows.get(15, [])
    rev_r   = rows.get(48, [])
    ebt_r   = rows.get(49, [])
    gm_r    = rows.get(50, [])
    em_r    = rows.get(51, [])
    cust_r  = rows.get(56, [])
    act_r   = rows.get(57, [])
    churn_r = rows.get(59, [])

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(4, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "NGN",
                   fx_rate_to_usd=FX["NGN"],
                   aum_usd=to_float(g(aum_r, col)),
                   revenue_usd=to_float(g(rev_r, col)),
                   ebitda_usd=to_float(g(ebt_r, col)),
                   gross_margin_pct=pct(g(gm_r, col)),
                   ebitda_margin_pct=pct(g(em_r, col)),
                   customer_count=to_int(g(cust_r, col)),
                   active_clients_count=to_int(g(act_r, col)),
                   gross_churn_rate_pct=pct(g(churn_r, col)))
        n += 1
    conn.commit()
    return n


def seed_yoco(conn):
    """Monthly, Jan 2017–Feb 2020. ZAR → USD at 18.5."""
    cid  = get_id(conn, "Yoco")
    rows = load_rows(DATA_DIR / "Yoco Quona KPIs_02112026.xlsx", "Copy of KPIsQuona")
    # Row 1: text date headers (col D = idx 3 onwards)
    # Row 13: active_merchants; Row 16: GMV ZAR; Row 17: revenue ZAR; Row 18: gross margin ZAR
    date_r = rows.get(1,  [])
    act_r  = rows.get(13, [])
    gmv_r  = rows.get(16, [])
    rev_r  = rows.get(17, [])
    gp_r   = rows.get(18, [])

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(3, len(date_r)):
        period = month_end_from_str(str(g(date_r, col) or ""))
        if not period:
            continue
        rev = to_usd(g(rev_r, col), "ZAR")
        gp  = to_usd(g(gp_r,  col), "ZAR")
        gm_pct = round(gp / rev * 100, 4) if rev and gp is not None and rev != 0 else None
        upsert_kpi(conn, cid, period, "ZAR",
                   fx_rate_to_usd=FX["ZAR"],
                   gmv_usd=to_usd(g(gmv_r, col), "ZAR"),
                   revenue_usd=rev,
                   gross_profit_usd=gp,
                   gross_margin_pct=gm_pct,
                   active_clients_count=to_int(g(act_r, col)))
        n += 1
    conn.commit()
    return n


def seed_verto(conn):
    """Monthly from Mar 2024. All USD."""
    cid  = get_id(conn, "Verto")
    rows = load_rows(DATA_DIR / "Verto Monthly KPIs - 2026-01_Board.xlsx", "KPIs")
    # Row 2: date headers (col D = idx 3 onwards)
    date_r = rows.get(2,  [])
    act_r  = rows.get(15, [])   # 1-month active clients
    gmv_r  = rows.get(22, [])   # Total Processed Volume USD
    rev_r  = rows.get(40, [])   # Total Revenue USD
    gp_r   = rows.get(57, [])   # Gross Profit USD
    gm_r   = rows.get(58, [])   # Gross Margin decimal
    ebt_r  = rows.get(59, [])   # EBITDA USD

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(3, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "USD",
                   fx_rate_to_usd=1.0,
                   gmv_usd=to_float(g(gmv_r, col)),
                   revenue_usd=to_float(g(rev_r, col)),
                   gross_profit_usd=to_float(g(gp_r, col)),
                   gross_margin_pct=pct(g(gm_r, col)),
                   ebitda_usd=to_float(g(ebt_r, col)),
                   active_clients_count=to_int(g(act_r, col)))
        n += 1
    conn.commit()
    return n


def seed_enza(conn):
    """Monthly from Jan 2025. USD."""
    cid  = get_id(conn, "Enza")
    rows = load_rows(DATA_DIR / "Copy of Enza KPI Template_Quona.xlsx", "KPIs")
    # Row 2: date headers (col E = idx 4 onwards)
    date_r = rows.get(2,  [])
    rev_r  = rows.get(7,  [])   # Total Revenues USD
    ebt_r  = rows.get(15, [])   # EBITDA USD
    cust_r = rows.get(26, [])   # Number of Clients (Banks/Institutions)

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(4, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "USD",
                   fx_rate_to_usd=1.0,
                   revenue_usd=to_float(g(rev_r, col)),
                   ebitda_usd=to_float(g(ebt_r, col)),
                   customer_count=to_int(g(cust_r, col)))
        n += 1
    conn.commit()
    return n


def seed_lulalend(conn):
    """Quarterly (reverse chronological). ZAR → USD at 18.5."""
    cid  = get_id(conn, "Lulalend")
    rows = load_rows(
        DATA_DIR / "Lulalend 12. Investor Man Acc - December 2025 - Quona.xlsx", "KPI's")
    # Row 5: quarter labels like "December 2025" (col D = idx 3 onwards)
    date_r  = rows.get(5,  [])
    rev_r   = rows.get(7,  [])   # Credit Revenue ZAR
    ebt_r   = rows.get(8,  [])   # EBITDA ZAR
    loan_r  = rows.get(13, [])   # Net Loan Portfolio ZAR
    par30_r = rows.get(19, [])   # Par 30 + Restructured decimal
    cust_r  = rows.get(22, [])   # Total businesses funded (unique)
    act_r   = rows.get(23, [])   # Total active clients

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(3, len(date_r)):
        period = month_end_from_str(str(g(date_r, col) or ""))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "ZAR",
                   fx_rate_to_usd=FX["ZAR"],
                   revenue_usd=to_usd(g(rev_r,  col), "ZAR"),
                   ebitda_usd=to_usd(g(ebt_r,  col), "ZAR"),
                   loan_book_gross_usd=to_usd(g(loan_r, col), "ZAR"),
                   par_30_pct=pct(g(par30_r, col)),
                   customer_count=to_int(g(cust_r, col)),
                   active_clients_count=to_int(g(act_r, col)))
        n += 1
    conn.commit()
    return n


def seed_khazna(conn):
    """Monthly Jan 2025–Mar 2026. USD."""
    cid  = get_id(conn, "Khazna")
    rows = load_rows(
        DATA_DIR / "Khazna Consolidated Mgmt Accts + KPIs Mar26 (2).xlsx", "KPIs")
    # Row 1: 'KPIs' label in col A, then dates from col B (idx 1 onwards)
    date_r  = rows.get(1,  [])
    reg_r   = rows.get(3,  [])   # Registered Users
    act_r   = rows.get(4,  [])   # Active Users
    loan_r  = rows.get(11, [])   # Outstanding loan portfolio USD
    par30_r = rows.get(16, [])   # PAR 30 decimal
    rev_r   = rows.get(20, [])   # Monthly Revenue USD
    arr_r   = rows.get(24, [])   # ARR USD
    ebt_r   = rows.get(25, [])   # Operating Profit (EBITDA proxy)

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(1, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "USD",
                   fx_rate_to_usd=1.0,
                   customer_count=to_int(g(reg_r,  col)),
                   active_clients_count=to_int(g(act_r,  col)),
                   loan_book_gross_usd=to_float(g(loan_r, col)),
                   par_30_pct=pct(g(par30_r, col)),
                   revenue_usd=to_float(g(rev_r, col)),
                   arr_usd=to_float(g(arr_r, col)),
                   ebitda_usd=to_float(g(ebt_r, col)))
        n += 1
    conn.commit()
    return n


def seed_twinco(conn):
    """Monthly from Jan 2025. GMV/revenue/loan in USD; EBITDA in EUR → USD."""
    cid  = get_id(conn, "TWINCO")
    rows = load_rows(DATA_DIR / "26.02.28 REP-TWINCO KPI FEB-26.xlsx", "KPIs")
    # Row 1: dates from col E (idx 4 onwards)
    date_r  = rows.get(1,  [])
    buyer_r = rows.get(5,  [])   # Active Anchor Buyers (customer_count)
    sup_r   = rows.get(7,  [])   # Transacting Suppliers (active_clients_count)
    gmv_r   = rows.get(18, [])   # Total Processed Volume USD
    loan_r  = rows.get(34, [])   # Loans outstanding USD
    par30_r = rows.get(36, [])   # PAR30 absolute USD (compute %)
    conc_r  = rows.get(41, [])   # Top 3 buyers concentration decimal
    rev_r   = rows.get(52, [])   # Revenue Accrued USD
    ebt_r   = rows.get(57, [])   # EBITDA EUR

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(4, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        loan     = to_float(g(loan_r,  col))
        par30abs = to_float(g(par30_r, col))
        par30pct = round(par30abs / loan * 100, 4) if (loan and par30abs is not None and loan != 0) else None
        ebt_eur  = to_float(g(ebt_r, col))
        ebt_usd  = (ebt_eur * FX["EUR"]) if ebt_eur is not None else None
        upsert_kpi(conn, cid, period, "EUR",
                   fx_rate_to_usd=FX["EUR"],
                   customer_count=to_int(g(buyer_r, col)),
                   active_clients_count=to_int(g(sup_r, col)),
                   gmv_usd=to_float(g(gmv_r, col)),
                   loan_book_gross_usd=loan,
                   par_30_pct=par30pct,
                   top_3_concentration_pct=pct(g(conc_r, col)),
                   revenue_usd=to_float(g(rev_r, col)),
                   ebitda_usd=ebt_usd)
        n += 1
    conn.commit()
    return n


def seed_maxsoko(conn):
    """MaxSoko's Excel file uses cross-sheet formula references with no cached
    values — openpyxl returns None for every data cell.  Skipped."""
    return None


def seed_sava(conn):
    """Monthly from Feb 2025. Col A = date, col C = revenue USD."""
    cid  = get_id(conn, "SAVA")
    rows = load_rows(DATA_DIR / "SAVA Revenue (Q1 2026).xlsx", "Revenue To Date")

    n = 0
    for row in rows.values():
        a = row[0] if len(row) > 0 else None
        if not isinstance(a, (datetime, date)):
            continue  # skip headers, year labels, totals
        period  = month_end_from_dt(a)
        rev_usd = to_float(row[2]) if len(row) > 2 else None
        if not period:
            continue
        upsert_kpi(conn, cid, period, "ZAR",
                   fx_rate_to_usd=FX["ZAR"],
                   revenue_usd=rev_usd)
        n += 1
    conn.commit()
    return n


def seed_alllife(conn):
    """Annual (years as columns). Values in ZAR 000s → × 1 000 / 18.5."""
    cid  = get_id(conn, "AllLife")
    rows = load_rows(DATA_DIR / "AllLife Data_Quona 2025.xlsx", "Data Template")
    # Row 3: year headers (col D = idx 3 onwards)
    # Row 4: Ongoing Revenue (ZAR 000s); Row 8: EBITDA (ZAR 000s)
    # Row 26: Active Customers with policies (raw count)
    date_r     = rows.get(3,  [])
    rev_r      = rows.get(4,  [])
    ebt_r      = rows.get(8,  [])
    policies_r = rows.get(26, [])

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(3, len(date_r)):
        period = year_end(g(date_r, col))
        if not period:
            continue
        rev_zar = to_float(g(rev_r, col))
        ebt_zar = to_float(g(ebt_r, col))
        upsert_kpi(conn, cid, period, "ZAR",
                   fx_rate_to_usd=FX["ZAR"],
                   revenue_usd=(rev_zar * 1000 / FX["ZAR"]) if rev_zar is not None else None,
                   ebitda_usd=(ebt_zar * 1000 / FX["ZAR"]) if ebt_zar is not None else None,
                   insurance_policies_active=to_int(g(policies_r, col)))
        n += 1
    conn.commit()
    return n


def seed_octa(conn):
    """Monthly Apr 2024–Jan 2026. USD."""
    cid  = get_id(conn, "OCTA")
    rows = load_rows(DATA_DIR / "Metrics - OCTA_Dec 2025.xlsx", "Sheet1")
    # Row 5: text date headers "Apr'24" etc. (col C = idx 2 onwards)
    date_r = rows.get(5,  [])
    cust_r = rows.get(6,  [])   # Cumulative customers onboarded
    tpv_r  = rows.get(14, [])   # TPV USD
    arr_r  = rows.get(16, [])   # ARR USD

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(2, len(date_r)):
        period = month_end_from_str(str(g(date_r, col) or ""))
        if not period:
            continue
        upsert_kpi(conn, cid, period, "USD",
                   fx_rate_to_usd=1.0,
                   customer_count=to_int(g(cust_r, col)),
                   tpv_usd=to_float(g(tpv_r, col)),
                   arr_usd=to_float(g(arr_r, col)))
        n += 1
    conn.commit()
    return n


def seed_eseye(conn):
    """Monthly from Jan 2018. GBP → USD at 1.27. Only 'Actual' columns."""
    cid  = get_id(conn, "Eseye")
    rows = load_rows(DATA_DIR / "Eseye Mgmt_accounts_2026- 02 final.xlsx", "KPIs")
    # Row 2: dates (col C = idx 2 onwards); Row 3: 'Actual' / 'Budget' labels
    date_r   = rows.get(2,  [])
    flag_r   = rows.get(3,  [])
    rev_r    = rows.get(8,  [])   # Total revenue GBP
    gp_r     = rows.get(10, [])   # Gross profit GBP
    ebt_r    = rows.get(14, [])   # EBITDA before bonus GBP
    cust_r   = rows.get(21, [])   # No of customers
    devic_r  = rows.get(22, [])   # No of devices

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(2, len(date_r)):
        if str(g(flag_r, col) or "").strip().lower() != "actual":
            continue
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        rev_usd = to_usd(g(rev_r,   col), "GBP")
        gp_usd  = to_usd(g(gp_r,   col), "GBP")
        ebt_usd = to_usd(g(ebt_r,  col), "GBP")
        gm_pct  = round(gp_usd  / rev_usd * 100, 4) if rev_usd and gp_usd  is not None and rev_usd != 0 else None
        em_pct  = round(ebt_usd / rev_usd * 100, 4) if rev_usd and ebt_usd is not None and rev_usd != 0 else None
        upsert_kpi(conn, cid, period, "GBP",
                   fx_rate_to_usd=FX["GBP"],
                   revenue_usd=rev_usd,
                   gross_profit_usd=gp_usd,
                   gross_margin_pct=gm_pct,
                   ebitda_usd=ebt_usd,
                   ebitda_margin_pct=em_pct,
                   customer_count=to_int(g(cust_r,  col)),
                   devices_connected=to_int(g(devic_r, col)))
        n += 1
    conn.commit()
    return n


def seed_power(conn):
    """Monthly from Jan 2025. USD P&L."""
    cid  = get_id(conn, "POWER")
    rows = load_rows(
        DATA_DIR / "POWER - 09 2025 Mgmt Accounts - Consolidated FS.xlsx", "IS_12MoM")
    # Row 8: date headers (col C = idx 2 onwards)
    # Row 20: NET INCOME (total revenues); Row 26: GROSS MARGINS; Row 34: EBITDA
    date_r = rows.get(8,  [])
    rev_r  = rows.get(20, [])
    gp_r   = rows.get(26, [])
    ebt_r  = rows.get(34, [])

    def g(row, col): return row[col] if col < len(row) else None

    n = 0
    for col in range(2, len(date_r)):
        period = month_end_from_dt(g(date_r, col))
        if not period:
            continue
        rev = to_float(g(rev_r, col))
        gp  = to_float(g(gp_r,  col))
        ebt = to_float(g(ebt_r, col))
        gm_pct = round(gp  / rev * 100, 4) if rev and gp  is not None and rev != 0 else None
        em_pct = round(ebt / rev * 100, 4) if rev and ebt is not None and rev != 0 else None
        upsert_kpi(conn, cid, period, "USD",
                   fx_rate_to_usd=1.0,
                   revenue_usd=rev,
                   gross_profit_usd=gp,
                   gross_margin_pct=gm_pct,
                   ebitda_usd=ebt,
                   ebitda_margin_pct=em_pct)
        n += 1
    conn.commit()
    return n


# ─── orchestrator ────────────────────────────────────────────────────────────

SEEDERS = [
    ("Cowrywise", seed_cowrywise),
    ("Yoco",      seed_yoco),
    ("Verto",     seed_verto),
    ("Enza",      seed_enza),
    ("Lulalend",  seed_lulalend),
    ("Khazna",    seed_khazna),
    ("TWINCO",    seed_twinco),
    ("MaxSoko",   seed_maxsoko),
    ("SAVA",      seed_sava),
    ("AllLife",   seed_alllife),
    ("OCTA",      seed_octa),
    ("Eseye",     seed_eseye),
    ("POWER",     seed_power),
]

def main():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")

    print("\n--- Step 1: seeding companies ---")
    seed_companies(conn)

    print("\n--- Step 2: seeding kpi_snapshots ---")
    total = 0
    for name, fn in SEEDERS:
        try:
            n = fn(conn)
            if n is None:
                print(f"  {name:12s}  SKIPPED — Excel file has no cached formula values")
            else:
                print(f"  {name:12s}  {n:4d} rows inserted")
                total += n
        except Exception as exc:
            print(f"  {name:12s}  WARNING: {exc}")

    conn.close()
    print(f"\n  Total kpi_snapshots rows inserted: {total}")

if __name__ == "__main__":
    main()
