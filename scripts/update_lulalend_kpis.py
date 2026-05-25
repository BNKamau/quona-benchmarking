"""
Update Lulalend (company_id=5) KPI data in benchmarking.db:

1. Add par_90_pct and unique_borrowers_count columns to kpi_snapshots
2. Recalculate all Lulalend USD figures from FX 18.5 → 18.3
   (revenue_usd, ebitda_usd, loan_book_gross_usd, gross_profit_usd)
3. Set Q4 2025 lending KPIs:
   - gross_margin_pct = 51.3
   - par_90_pct = 5.4
   - unique_borrowers_count = 24921  (cumulative unique SMEs funded)

Run from project root:  python scripts/update_lulalend_kpis.py
"""

import sqlite3
from datetime import datetime

DB_PATH = "benchmarking.db"
FX_OLD = 18.5
FX_NEW = 18.3
FX_FACTOR = FX_OLD / FX_NEW  # ~1.01093  —  multiply USD values to restate at 18.3

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
now = datetime.utcnow().isoformat()

# ── 1. Add new columns (no-op if already present) ──────────────────────────────
existing_cols = {c["name"] for c in conn.execute("PRAGMA table_info(kpi_snapshots)")}
NEW_COLS = [
    ("par_90_pct",             "ALTER TABLE kpi_snapshots ADD COLUMN par_90_pct REAL"),
    ("unique_borrowers_count", "ALTER TABLE kpi_snapshots ADD COLUMN unique_borrowers_count INTEGER"),
]
for col_name, ddl in NEW_COLS:
    if col_name not in existing_cols:
        conn.execute(ddl)
        print(f"  Added column: {col_name}")
    else:
        print(f"  Column already exists: {col_name}")

# ── 2. FX correction on all Lulalend rows ─────────────────────────────────────
# Original data was loaded at ZAR/18.5; restate to ZAR/18.3
conn.execute(f"""
    UPDATE kpi_snapshots
    SET
        revenue_usd         = ROUND(revenue_usd         * {FX_FACTOR}, 2),
        ebitda_usd          = ROUND(ebitda_usd           * {FX_FACTOR}, 2),
        loan_book_gross_usd = ROUND(loan_book_gross_usd  * {FX_FACTOR}, 2),
        gross_profit_usd    = CASE WHEN gross_profit_usd IS NOT NULL
                                   THEN ROUND(gross_profit_usd * {FX_FACTOR}, 2)
                                   ELSE NULL END,
        fx_rate_to_usd      = {FX_NEW},
        updated_at          = '{now}'
    WHERE company_id = 5
""")
n_updated = conn.execute(
    "SELECT changes()"
).fetchone()[0]
print(f"\nFX correction applied to {n_updated} Lulalend rows ({FX_OLD} -> {FX_NEW})")

# ── 3. Q4 2025 lending KPI additions ──────────────────────────────────────────
conn.execute(f"""
    UPDATE kpi_snapshots
    SET
        gross_margin_pct        = 51.3,
        par_90_pct              = 5.4,
        unique_borrowers_count  = 24921,
        updated_at              = '{now}'
    WHERE company_id = 5
      AND period_end_date = '2025-12-31'
""")
print("Q4 2025: gross_margin_pct=51.3, par_90_pct=5.4, unique_borrowers_count=24921")

conn.commit()

# ── Verify ─────────────────────────────────────────────────────────────────────
print("\n=== Verification: last 4 quarters ===")
rows = conn.execute("""
    SELECT period_end_date,
           revenue_usd, ebitda_usd, loan_book_gross_usd,
           gross_margin_pct, par_30_pct, par_90_pct,
           net_yield_pct, active_clients_count, unique_borrowers_count,
           fx_rate_to_usd
    FROM kpi_snapshots
    WHERE company_id = 5
    ORDER BY period_end_date DESC
    LIMIT 4
""").fetchall()
for r in rows:
    print(dict(r))

# LTM totals
ltm_rev = conn.execute("""
    SELECT SUM(revenue_usd), SUM(ebitda_usd)
    FROM (
        SELECT revenue_usd, ebitda_usd
        FROM kpi_snapshots
        WHERE company_id = 5 AND revenue_usd IS NOT NULL
        ORDER BY period_end_date DESC
        LIMIT 4
    )
""").fetchone()
ltm_rev_usd   = ltm_rev[0]
ltm_ebitda    = ltm_rev[1]
ltm_em        = ltm_ebitda / ltm_rev_usd * 100 if ltm_rev_usd else None
print(f"\n=== LTM totals (last 4 quarters) ===")
print(f"  LTM Revenue:      ${ltm_rev_usd/1e6:.1f}M")
print(f"  LTM EBITDA:       ${ltm_ebitda/1e6:.2f}M")
print(f"  LTM EBITDA Margin: {ltm_em:.1f}%")

conn.close()
print("\nDone.")
