# Quona Portfolio Intelligence

Internal Streamlit dashboard for monitoring Quona portfolio companies. Stores monthly/quarterly KPI snapshots in a local SQLite database and renders a portfolio home page plus per-company detail pages.

**Run the app:**

```bash
streamlit run app.py
```

---

## Repository layout

```
quona-benchmarking/
├── app.py                        # Streamlit dashboard (single-file)
├── benchmarking.db               # SQLite database (not committed)
├── migrations/
│   ├── 001_initial_schema.sql    # Core tables
│   ├── 002_add_portfolio_fields.sql  # Sector-specific KPI columns
│   └── 003_update_companies_schema.sql  # sub_sector + widened business_model
├── scripts/
│   ├── migrate.py                # Migration runner
│   ├── seed_portfolio.py         # Initial data load from Excel files
│   └── fix_company_data.py       # Supplementary fixes (Yoco, Lulalend, MaxSoko, OCTA)
└── data/                         # Source Excel files (not committed)
```

---

## Running migrations

```bash
python scripts/migrate.py                   # default: benchmarking.db in project root
python scripts/migrate.py path/to/other.db  # custom path
```

The runner records each applied migration in `schema_migrations` and skips anything already applied. Add new migrations as `migrations/004_description.sql`.

---

## Database schema

### `schema_migrations`
Tracks applied migrations. Never edit manually.

| Column | Type | Notes |
|---|---|---|
| `version` | INTEGER PK | Migration number |
| `description` | TEXT | Short label |
| `applied_at` | TEXT | UTC timestamp |

---

### `companies`
Master registry. Both portfolio companies and exit comps live here.

| Column | Type | Notes |
|---|---|---|
| `id` | INTEGER PK | Auto-increment |
| `name` | TEXT | Company name |
| `type` | TEXT | `'portfolio'` or `'exit_comp'` |
| `sector` | TEXT | e.g. `payments`, `lending`, `insurtech`, `marketplace`, `saas`, `iot_infrastructure`, `wealth_management` |
| `sub_sector` | TEXT | Added in migration 003 (e.g. `sme_lending`, `card_issuing_paas`) |
| `hq_country` | TEXT | ISO 3166-1 alpha-2 |
| `founded_year` | INTEGER | |
| `business_model` | TEXT | `b2b`, `b2c`, `b2b2c`, `lending`, `saas`, `marketplace`, `payments`, `other` |
| `reporting_currency` | TEXT | Primary source currency before USD conversion (e.g. `ZAR`, `NGN`, `GBP`) |
| `notes` | TEXT | |
| `created_at` / `updated_at` | TEXT | Auto-managed by trigger |

---

### `portfolio_metadata`
One row per portfolio company. All fields except `company_id` are nullable.

| Column | Type | Notes |
|---|---|---|
| `company_id` | INTEGER FK | References `companies.id` |
| `investment_date` | TEXT | ISO 8601 |
| `stage_at_investment` | TEXT | `seed`, `series_a`, etc. |
| `current_stage` | TEXT | |
| `ownership_pct` | REAL | |
| `fund` | TEXT | e.g. `Quona Fund III` |
| `notes` | TEXT | |

---

### `kpi_snapshots`
One row per company per period (monthly, quarterly, or annual). Every metric column is nullable.

| Column | Type | Notes |
|---|---|---|
| `company_id` | INTEGER FK | References `companies.id` |
| `period_end_date` | TEXT | ISO 8601 month/quarter/year end |
| `reporting_currency` | TEXT | Default `USD` |
| `fx_rate_to_usd` | REAL | Spot rate at period end |
| `revenue_usd` | REAL | |
| `gross_profit_usd` | REAL | |
| `gross_margin_pct` | REAL | 0–100 scale |
| `ebitda_usd` | REAL | |
| `ebitda_margin_pct` | REAL | Can be negative; also derived in-dashboard from ebitda/revenue |
| `arr_usd` | REAL | NULL for non-SaaS |
| `mrr_usd` | REAL | |
| `customer_count` | INTEGER | |
| `net_revenue_retention_pct` | REAL | NRR |
| `gross_churn_rate_pct` | REAL | |
| `cac_usd` | REAL | |
| `ltv_usd` | REAL | |
| `loan_book_gross_usd` | REAL | Lending only |
| `npl_rate_pct` | REAL | Non-performing loans as % of book |
| `net_yield_pct` | REAL | Lending only |
| `cost_of_risk_pct` | REAL | Lending only |
| `nim_pct` | REAL | Net interest margin |
| `leverage_ratio` | REAL | Debt / equity |
| `aum_usd` | REAL | Assets Under Management (Cowrywise) |
| `gmv_usd` | REAL | Gross Merchandise / Payment Volume |
| `active_clients_count` | INTEGER | Active merchants / clients / users |
| `par_30_pct` | REAL | Portfolio at Risk 30d as % of loan book |
| `top_3_concentration_pct` | REAL | Revenue concentration of top 3 customers |
| `insurance_policies_active` | INTEGER | Active insurance policies (AllLife) |
| `tpv_usd` | REAL | Total Payment Volume (OCTA) |
| `devices_connected` | INTEGER | Connected IoT devices (Eseye) |

Unique constraint on `(company_id, period_end_date)`. `updated_at` maintained by trigger.

---

### `exit_comps`
Financials and multiples at exit. One row per company.

| Column | Type | Notes |
|---|---|---|
| `company_id` | INTEGER FK | References `companies.id` |
| `exit_date` | TEXT | |
| `exit_type` | TEXT | `ipo`, `acquisition`, `merger`, `secondary` |
| `enterprise_value_usd` | REAL | |
| `revenue_ttm_usd` | REAL | Trailing twelve months |
| `ebitda_ttm_usd` | REAL | |
| `arr_at_exit_usd` | REAL | NULL for non-SaaS |
| `ev_revenue_multiple` | REAL | |
| `ev_ebitda_multiple` | REAL | |
| `ev_arr_multiple` | REAL | NULL for non-SaaS |

---

### `funding_stage_snapshots`
What a company looked like at each raise. Primarily for exit comps.

| Column | Type | Notes |
|---|---|---|
| `company_id` | INTEGER FK | |
| `stage` | TEXT | `seed`, `series_a`, `series_b`, `series_c`, `series_d_plus`, `growth` |
| `snapshot_date` | TEXT | |
| `raise_amount_usd` | REAL | |
| `post_money_valuation_usd` | REAL | |
| `revenue_ttm_usd` | REAL | |
| `arr_usd` | REAL | |
| `gross_margin_pct` | REAL | |
| `revenue_multiple_at_raise` | REAL | |
| `employee_count` | INTEGER | |

Unique constraint on `(company_id, stage)`.

---

### `benchmarks`
Computed percentile aggregates. Cleared and rewritten each benchmarking run — do not edit manually.

| Column | Type | Notes |
|---|---|---|
| `cohort_label` | TEXT | e.g. `Series B, lending, EMEA, 2020-2024` |
| `metric` | TEXT | e.g. `gross_margin_pct` |
| `p25` / `p50` / `p75` / `p90` | REAL | Percentile values |
| `sample_size` | INTEGER | |

---

## Migrations applied

| # | File | What it does |
|---|---|---|
| 001 | `001_initial_schema.sql` | Creates all six core tables plus `updated_at` triggers |
| 002 | `002_add_portfolio_fields.sql` | Adds `aum_usd`, `gmv_usd`, `active_clients_count`, `par_30_pct`, `top_3_concentration_pct`, `insurance_policies_active`, `tpv_usd`, `devices_connected` to `kpi_snapshots`; adds `reporting_currency` to `companies` |
| 003 | `003_update_companies_schema.sql` | Adds `sub_sector` column; widens `business_model` CHECK to include `b2b`, `b2c`, `b2b2c`; recreates `companies` table via safe copy-rename |

---

## Company registry

All 13 portfolio companies as seeded. The `id` column is stable and referenced throughout.

| id | Name | Sector | Sub-sector | HQ | Model | Source currency |
|---|---|---|---|---|---|---|
| 1 | Cowrywise | wealth_management | savings_and_investment | NG | b2c | NGN |
| 2 | Yoco | payments | merchant_acquiring | ZA | b2b | ZAR |
| 3 | Verto | payments | cross_border_fx | NG | b2b | USD |
| 4 | Enza | payments | card_issuing_paas | KE | b2b | USD |
| 5 | Lulalend | lending | sme_lending | ZA | b2b | ZAR |
| 6 | Khazna | lending | consumer_lending | EG | b2c | USD |
| 7 | TWINCO | lending | supply_chain_finance | ES | b2b | EUR |
| 8 | MaxSoko | marketplace | ecommerce_embedded_finance | EG | b2b | EGP |
| 9 | SAVA | payments | card_issuing_baas | ZA | b2b | ZAR |
| 10 | AllLife | insurtech | life_insurance | ZA | b2c | ZAR |
| 11 | OCTA | saas | invoice_ar_automation | NG | b2b | USD |
| 12 | Eseye | iot_infrastructure | managed_connectivity | GB | b2b | GBP |
| 13 | POWER | lending | earned_wage_access | US | b2b | USD |

**Fixed FX rates used for conversion** (applied at seeding time, not live):

| Currency | Rate | Direction |
|---|---|---|
| ZAR | 18.5 | ÷ 18.5 to get USD |
| NGN | 1500.0 | ÷ 1500 to get USD |
| GBP | 1.27 | × 1.27 to get USD |
| EUR | 1.08 | × 1.08 to get USD |

---

## Data pipeline

### Step 1 — Apply migrations

```bash
python scripts/migrate.py
```

### Step 2 — Seed portfolio companies

```bash
python scripts/seed_portfolio.py
```

Inserts all 13 companies into `companies` (idempotent via `INSERT OR IGNORE`), then reads each company's Excel file and upserts rows into `kpi_snapshots`.

**Per-company reader summary:**

| Company | File | Sheet | Period | Key rows read |
|---|---|---|---|---|
| Cowrywise | `Cowrywise KPIs_Mar 2026.xlsx` | Cowrywise | Monthly Jan 2020–Mar 2026 | AUM, Revenue, EBITDA, Gross/EBITDA margin, customers, churn |
| Yoco | `Yoco Quona KPIs_02112026.xlsx` | Copy of KPIsQuona | Monthly Jan 2017–Feb 2020 | Active merchants, GMV, Revenue, Gross profit (ZAR→USD) |
| Verto | `Verto Monthly KPIs - 2026-01_Board.xlsx` | KPIs | Monthly Feb 2024–Jan 2026 | Active clients, GMV, Revenue, Gross profit, EBITDA |
| Enza | `Copy of Enza KPI Template_Quona.xlsx` | KPIs | Monthly Jan 2025–Mar 2026 | Revenue, EBITDA, customer count |
| Lulalend | `Lulalend 12. Investor Man Acc - December 2025 - Quona.xlsx` | KPI's | Quarterly Dec 2022–Dec 2025 | Revenue, EBITDA, loan book, PAR30, customers (ZAR→USD) |
| Khazna | `Khazna Consolidated Mgmt Accts + KPIs Mar26 (2).xlsx` | KPIs | Monthly Jan 2025–Mar 2026 | Customers, active users, loan book, PAR30, revenue, ARR, EBITDA |
| TWINCO | `26.02.28 REP-TWINCO KPI FEB-26.xlsx` | KPIs | Monthly Jan 2025–Feb 2026 | Buyers, suppliers, GMV, loan book, PAR30, concentration, revenue, EBITDA (EUR→USD) |
| MaxSoko | *(skipped — no cached formula values)* | — | — | Summary/Consolidated sheets return `#VALUE!` with `data_only=True`; seeded via `fix_company_data.py` instead |
| SAVA | `SAVA Revenue (Q1 2026).xlsx` | Revenue To Date | Monthly Feb 2025–Mar 2026 | Revenue USD (pre-converted in source) |
| AllLife | `AllLife Data_Quona 2025.xlsx` | Data Template | Annual 2018–2025 | Revenue, EBITDA (ZAR 000s → USD), active policies |
| OCTA | `Metrics - OCTA_Dec 2025.xlsx` | Sheet1 | Monthly Apr 2024–Jan 2026 | Customers, TPV, ARR — note: this file is superseded; see manual fixes below |
| Eseye | `Eseye Mgmt_accounts_2026- 02 final.xlsx` | KPIs | Monthly Jan 2018–Feb 2026 (Actual only) | Revenue, Gross profit, EBITDA, customers, devices (GBP→USD) |
| POWER | `POWER - 09 2025 Mgmt Accounts - Consolidated FS.xlsx` | IS_12MoM | Monthly Jan 2025–Sep 2025 | Revenue, Gross profit, EBITDA |

### Step 3 — Supplementary fixes

```bash
python scripts/fix_company_data.py
```

Fills gaps in the initial seed. Run after `seed_portfolio.py`:

| Company | What it adds |
|---|---|
| Yoco | Re-reads `KPIQuona Export Doc` sheet for Jan 2023–Dec 2025: updates revenue, gross profit, gross margin, GMV, customer count, active merchants |
| Lulalend | Adds `net_yield_pct` (Average Annualized Interest Rate) from row 18 of KPI's sheet |
| Khazna | No-op (all rows already complete) |
| AllLife | No-op (gross margin not in source) |
| OCTA | Seeds `revenue_usd` from `arr_usd` for pre-2026 rows (later superseded — see manual fixes) |
| MaxSoko | Reads `Redash EG` sheet (hardcoded values, no cross-sheet formulas): seeds 36 months (Jan 2023–Dec 2025) of revenue, GMV, customer count, gross profit |

---

## Manual data corrections

The following fixes were applied interactively via ad-hoc Python scripts (not persisted as rerunnable scripts). They are reflected in the current `benchmarking.db`.

### MaxSoko — gross margin recalculation

**Problem:** `fix_company_data.py` seeded gross margin using only Front Margin (row 14 of Redash EG), which turned negative from March 2024 onward when COGS exceeded NMV.

**Fix:** Recomputed gross margin as `(Front Margin + Back Margin) / NMV excl VAT × 100` for all 36 rows. Back Margin (row 15 — supplier rebates) is always large and positive, making the combined gross margin ~3–10% across the full history.

Result: gross margin now ranges 3.8%–9.5%, average 5.9% — consistent with a marketplace take-rate model.

### SAVA — March 2026 revenue correction

**Problem:** March 2026 revenue was missing or incorrect.

**Fix:** Derived from the `Q1 Revenue Breakdown` sheet of `SAVA Revenue (Q1 2026).xlsx`. Q1 total = ZAR 2,456,112.50 = $150,681.75. Jan = $21,387.20, Feb = $53,310.21, so Mar = $75,984.34.

Updated `period_end_date = '2026-03-31'` from the incorrect value to $75,984.34.

### OCTA — revenue data replaced with P&L actuals

**Problem:** The initial seed propagated `arr_usd` (contracted ARR milestones from `Metrics - OCTA_Dec 2025.xlsx`) into `revenue_usd`. These were commitment-based figures ($36K–$1.55M ARR), not actual monthly P&L revenue.

**Fix:** Nulled `revenue_usd` for all 21 pre-2026 rows (no reliable monthly actuals exist for that period). Seeded Q1 2026 actuals from `P&L and BS Q1 2026.xlsx` (PL sheet, Net Revenue row):

| Period | Revenue USD | Source |
|---|---|---|
| 2026-01-31 | $66,998.04 | P&L actuals |
| 2026-02-28 | $67,221.04 | P&L actuals |
| 2026-03-31 | $67,783.04 | P&L actuals |

LTM shown on dashboard = Q1 total × 4 = ~$808K (ARR est., 3 of 12 periods).

### Khazna — gross margin seeded from P&L Consolidated

**Problem:** `gross_profit_usd` and `gross_margin_pct` were NULL for all 15 rows.

**Fix:** Read `PL - Consolidated` sheet of `Khazna Consolidated Mgmt Accts + KPIs Mar26 (2).xlsx`:
- Row 12: Gross Margin (absolute USD)
- Row 13: % of Gross Margin

Applied:
- Jan–Dec 2025 (12 rows): `gross_margin_pct = 54.31%` (FY2025 blended rate); `gross_profit_usd = revenue_usd × 0.5431`
- Jan 2026: GP = $385,421.87 / margin = 50.75%
- Feb 2026: GP = $436,676.76 / margin = 57.31%
- Mar 2026: GP = $372,815.63 / margin = 57.64%

### OCTA — gross margin seeded from P&L

**Problem:** `gross_profit_usd` and `gross_margin_pct` were NULL for all rows.

**Fix:** Read `GROSS PROFIT I` row from `P&L and BS Q1 2026.xlsx` for Q1 2026:

| Period | Gross Profit USD | Gross Margin % |
|---|---|---|
| 2026-01-31 | $63,517.61 | 94.81% |
| 2026-02-28 | $64,666.64 | 96.20% |
| 2026-03-31 | $65,883.65 | 97.20% |

---

## Current data coverage (as of 2026-05-06)

| Company | KPI rows | Period range | Rev rows | GM rows | Avg GM% | EBITDA rows |
|---|---|---|---|---|---|---|
| Cowrywise | 75 | Jan 2020–Mar 2026 | 75 | 75 | 81.2% | 75 |
| Yoco | 71 | Jan 2017–Dec 2025 | 70 | 70 | 43.8% | 0 |
| Verto | 24 | Feb 2024–Jan 2026 | 24 | 24 | 79.2% | 24 |
| Enza | 15 | Jan 2025–Mar 2026 | 15 | 0 | — | 15 |
| Lulalend | 13 | Dec 2022–Dec 2025 | 13 | 0 | — | 13 |
| Khazna | 15 | Jan 2025–Mar 2026 | 15 | 15 | 54.5% | 15 |
| TWINCO | 14 | Jan 2025–Feb 2026 | 14 | 0 | — | 14 |
| MaxSoko | 36 | Jan 2023–Dec 2025 | 36 | 36 | 5.9% | 0 |
| SAVA | 14 | Feb 2025–Mar 2026 | 14 | 0 | — | 0 |
| AllLife | 8 | Dec 2018–Dec 2025 | 8 | 0 | — | 8 |
| OCTA | 24 | Apr 2024–Mar 2026 | 3 | 3 | 96.1% | 0 |
| Eseye | 98 | Jan 2018–Feb 2026 | 98 | 98 | 49.6% | 98 |
| POWER | 9 | Jan 2025–Sep 2025 | 9 | 9 | 85.4% | 9 |

**Gross margin not available** for Enza, Lulalend, TWINCO, SAVA, AllLife — source files do not contain gross profit line items.

---

## Dashboard (`app.py`)

Single-file Streamlit app. No multi-page directory — routing is handled via `st.session_state.page`.

### Home page

**Summary bar:** 4 `st.metric` tiles — portfolio company count, combined LTM Revenue, avg Gross Margin, avg EBITDA Margin.

**Company table:** One row per company with columns:
- Company name (clickable button → detail page) + inline data quality badge
- Sector chip + HQ country
- LTM Revenue with basis label (`LTM · 12 mo.` / `ARR est. · N of 12`)
- Gross Margin % (green if >50%)
- EBITDA Margin % (green if positive, red if negative)
- Revenue Growth % (period-over-period, last two data points)
- As of date (most recent period)

**LTM Revenue calculation** (`load_ltm_revenue()`):
- Period type is inferred from the gap between the two most recent rows: ≤45 days = monthly (need 12), ≤135 days = quarterly (need 4), >135 days = annual (need 1)
- If the company has ≥ N periods: LTM = sum of the N most recent
- If fewer: ARR est. = sum(available) × (N / n), labelled `ARR (est.)`

**Data quality flags** (`compute_data_quality_flags()`):
- `DATA STALE` — latest period is more than 6 months before 2026-05-05
- `CHECK: NEGATIVE MARGIN` — gross margin < 0% at latest period
- `CHECK: UNUSUALLY HIGH MARGIN` — gross margin > 95% at latest period
- `CHECK: EXTREME BURN` — EBITDA margin < −200%
- `DATA INCOMPLETE` — fewer than 6 months equivalent of revenue history
- `CHECK: REVENUE VOLATILITY` — any consecutive period change > 80%

Flags appear inline as small amber badges below the company name, and in full in a **DATA QUALITY FLAGS** table below the company list.

**Methodology note** displayed below the table explains LTM and Revenue Growth calculations.

**LTM summary expander** — full `st.dataframe` with all companies, LTM figure, basis, margins, and flags.

### Detail page

Accessed by clicking a company name. Back button restores home page.

**Header card:** Company initial avatar, name, sector chip, HQ country, founded year.

**6-metric summary row:**
1. LTM Revenue (or ARR est.)
2. Revenue (latest period)
3. Gross Margin (latest)
4. EBITDA Margin (latest)
5. Customers / Active Clients (latest, whichever is populated)
6. History span (period count + date range)

**Chart sections** (shown only if ≥2 data points available):
- Financial Performance: Revenue (USD), Gross Margin %, EBITDA Margin %, Customer / Active Client count
- Lending & Credit Metrics (if present): Gross Loan Book, PAR 30%, NPL Rate, Net Yield, NIM
- Assets Under Management (if present): AUM chart for Cowrywise
- Sector Metrics: GMV, TPV, ARR, NRR — shown if not already in Lending section

All charts use Plotly `Scatter` with `lines+markers`, Quona brand colours (black line, green fill), unified hover.

**Raw data table** — expandable `st.dataframe` showing all non-empty columns in reverse chronological order.

### Styling

Brand palette constants at the top of `app.py`:

| Constant | Hex | Usage |
|---|---|---|
| `GREEN` | `#D5FA94` | Primary accent (buttons, avatars) |
| `BLACK` | `#2C2C2A` | Text, chart lines |
| `BLUE` | `#C5E5FF` | Sector chips |
| `BG` | `#EFF0EA` | Page background |
| `WHITE` | `#FFFFFF` | Card backgrounds |
| `BORDER` | `#DDE0D8` | Card/chart borders |
| `MUTED` | `#888884` | Secondary labels |
| `WARN` | `#E65100` | Data quality flag text |
| `WARN_BG` | `#FFF3E0` | Data quality flag background |

---

## Key decisions and known gaps

**MaxSoko Excel limitation:** The Summary and Consolidated sheets use cross-sheet formula references that are not cached in the file. `openpyxl` with `data_only=True` returns `None` for every formula cell. Only the `Redash EG` sheet has hardcoded values and is usable.

**OCTA pre-2026 revenue:** No monthly P&L actuals exist before 2026. The `arr_usd` contract values ($36K–$1.55M ARR) stored in pre-2026 rows are not monthly revenue figures and are intentionally excluded from `revenue_usd`. OCTA's LTM is therefore an ARR estimate from Q1 2026 × 4 = ~$808K.

**Gross margin gaps:** Enza, Lulalend, TWINCO, SAVA, AllLife do not provide gross profit line items in their source files. The `gross_margin_pct` column is NULL for these companies across all periods.

**Yoco EBITDA:** EBITDA data was not available in the Yoco Excel files as of the last data pull.

**Fixed FX rates:** All non-USD conversion uses static rates (ZAR 18.5, NGN 1500, GBP 1.27, EUR 1.08) baked in at seeding time. The `fx_rate_to_usd` column stores the rate used. Updating rates requires re-running the seed/fix scripts with revised `FX` constants.

**No exit comps seeded yet:** The `exit_comps`, `funding_stage_snapshots`, and `benchmarks` tables exist in the schema but are empty. Benchmarking against exit comps is the next major feature to build.

---

## Next steps

1. **Exit comps data** — populate `exit_comps` and `funding_stage_snapshots` with comparable exits; build the benchmarking engine that writes percentile cohorts to `benchmarks`.

2. **Update OCTA history** — if pre-2026 monthly P&L actuals become available, seed them and remove the Q1-only limitation.

3. **Live FX rates** — replace the fixed `FX` dict with a rate lookup at period end date (e.g. from an open exchange rate API) stored in `fx_rate_to_usd`.

4. **Refresh script** — a single entrypoint that re-runs the full pipeline (migrate → seed → fix → gross margin patch) idempotently so any analyst can refresh from updated Excel files.

5. **Gross margin for remaining companies** — Enza, Lulalend, TWINCO, SAVA, AllLife. Would require sourcing management accounts that include COGS detail.

6. **Yoco EBITDA** — the source file has a complete P&L in a different sheet; extract and seed.
