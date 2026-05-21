-- Migration 001: Initial schema
-- Apply with: python scripts/migrate.py

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;

-- Tracks which migrations have been applied. Never edit this table manually.
CREATE TABLE IF NOT EXISTS schema_migrations (
    version     INTEGER PRIMARY KEY,
    description TEXT    NOT NULL,
    applied_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- ─────────────────────────────────────────────
-- Master company registry
-- Both portfolio companies and exit comps live here.
-- type = 'portfolio' | 'exit_comp'
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS companies (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    name           TEXT    NOT NULL,
    type           TEXT    NOT NULL CHECK (type IN ('portfolio', 'exit_comp')),
    sector         TEXT,               -- e.g. 'payments', 'lending', 'insurtech', 'b2b_saas'
    hq_country     TEXT,               -- ISO 3166-1 alpha-2, e.g. 'NG', 'ZA', 'IN'
    founded_year   INTEGER,
    business_model TEXT    CHECK (business_model IN ('lending', 'saas', 'marketplace', 'payments', 'other')),
    notes          TEXT,
    created_at     TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at     TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE TRIGGER companies_updated_at
AFTER UPDATE ON companies
FOR EACH ROW BEGIN
    UPDATE companies SET updated_at = datetime('now') WHERE id = OLD.id;
END;


-- ─────────────────────────────────────────────
-- Portfolio-specific metadata
-- One row per portfolio company. All fields except company_id are nullable
-- so a company can be added before deal terms are confirmed.
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS portfolio_metadata (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id            INTEGER NOT NULL UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
    investment_date       TEXT,       -- ISO 8601: '2022-09-01'
    stage_at_investment   TEXT,       -- 'seed', 'series_a', 'series_b', etc.
    current_stage         TEXT,
    ownership_pct         REAL,
    fund                  TEXT,       -- e.g. 'Quona Fund III'
    notes                 TEXT,
    created_at            TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at            TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE TRIGGER portfolio_metadata_updated_at
AFTER UPDATE ON portfolio_metadata
FOR EACH ROW BEGIN
    UPDATE portfolio_metadata SET updated_at = datetime('now') WHERE id = OLD.id;
END;


-- ─────────────────────────────────────────────
-- Quarterly KPI snapshots (long/narrow format)
-- One row per company per quarter. Every metric column is nullable —
-- a row with only company_id + period_end_date is valid and useful.
-- Lending-specific columns are NULL for non-lending companies.
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS kpi_snapshots (
    id                         INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id                 INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    period_end_date            TEXT    NOT NULL,    -- ISO 8601: '2024-03-31'
    reporting_currency         TEXT    NOT NULL DEFAULT 'USD',
    fx_rate_to_usd             REAL,               -- spot rate at period end

    -- Core P&L
    revenue_usd                REAL,
    gross_profit_usd           REAL,
    gross_margin_pct           REAL,               -- 0–100 scale, e.g. 62.5
    ebitda_usd                 REAL,
    ebitda_margin_pct          REAL,               -- can be negative

    -- Growth & customer metrics
    arr_usd                    REAL,               -- NULL for non-SaaS
    mrr_usd                    REAL,
    customer_count             INTEGER,
    net_revenue_retention_pct  REAL,               -- NRR, e.g. 115.0
    gross_churn_rate_pct       REAL,
    cac_usd                    REAL,
    ltv_usd                    REAL,

    -- Lending-specific (NULL for non-lending businesses)
    loan_book_gross_usd        REAL,
    npl_rate_pct               REAL,               -- non-performing loans as % of book
    net_yield_pct              REAL,
    cost_of_risk_pct           REAL,
    nim_pct                    REAL,               -- net interest margin
    leverage_ratio             REAL,               -- debt / equity

    notes                      TEXT,
    created_at                 TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at                 TEXT    NOT NULL DEFAULT (datetime('now')),

    UNIQUE (company_id, period_end_date)
);

CREATE TRIGGER kpi_snapshots_updated_at
AFTER UPDATE ON kpi_snapshots
FOR EACH ROW BEGIN
    UPDATE kpi_snapshots SET updated_at = datetime('now') WHERE id = OLD.id;
END;


-- ─────────────────────────────────────────────
-- Exit comp data
-- Financials and multiples at the moment of exit.
-- One row per company (a company only exits once).
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS exit_comps (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id            INTEGER NOT NULL UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
    exit_date             TEXT,
    exit_type             TEXT CHECK (exit_type IN ('ipo', 'acquisition', 'merger', 'secondary')),
    acquirer_name         TEXT,
    acquirer_type         TEXT CHECK (acquirer_type IN ('strategic', 'pe', 'public_market')),

    -- Financials at exit (trailing twelve months)
    enterprise_value_usd  REAL,
    revenue_ttm_usd       REAL,
    ebitda_ttm_usd        REAL,
    arr_at_exit_usd       REAL,       -- NULL for non-SaaS

    -- Exit multiples (can be derived, but stored for convenience)
    ev_revenue_multiple   REAL,
    ev_ebitda_multiple    REAL,
    ev_arr_multiple       REAL,       -- NULL for non-SaaS

    notes                 TEXT,
    created_at            TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at            TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE TRIGGER exit_comps_updated_at
AFTER UPDATE ON exit_comps
FOR EACH ROW BEGIN
    UPDATE exit_comps SET updated_at = datetime('now') WHERE id = OLD.id;
END;


-- ─────────────────────────────────────────────
-- Funding stage snapshots
-- What a company looked like at each raise — primarily for exit comps
-- so you can ask "what did this company look like at Series B?"
-- Also useful for portfolio companies once they have raised follow-ons.
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS funding_stage_snapshots (
    id                         INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id                 INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    stage                      TEXT    NOT NULL CHECK (stage IN (
                                   'seed', 'series_a', 'series_b',
                                   'series_c', 'series_d_plus', 'growth'
                               )),
    snapshot_date              TEXT,
    raise_amount_usd           REAL,
    post_money_valuation_usd   REAL,
    revenue_ttm_usd            REAL,
    arr_usd                    REAL,
    gross_margin_pct           REAL,
    revenue_multiple_at_raise  REAL,
    employee_count             INTEGER,
    notes                      TEXT,
    created_at                 TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at                 TEXT    NOT NULL DEFAULT (datetime('now')),

    UNIQUE (company_id, stage)
);

CREATE TRIGGER funding_stage_snapshots_updated_at
AFTER UPDATE ON funding_stage_snapshots
FOR EACH ROW BEGIN
    UPDATE funding_stage_snapshots SET updated_at = datetime('now') WHERE id = OLD.id;
END;


-- ─────────────────────────────────────────────
-- Computed benchmark aggregates
-- Written by the benchmarking script, not edited manually.
-- Cleared and recomputed each run.
-- ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS benchmarks (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    cohort_label     TEXT    NOT NULL,  -- e.g. 'Series B, lending, EMEA, 2020-2024'
    metric           TEXT    NOT NULL,  -- e.g. 'gross_margin_pct'
    p25              REAL,
    p50              REAL,
    p75              REAL,
    p90              REAL,
    sample_size      INTEGER,
    last_computed_at TEXT    NOT NULL DEFAULT (datetime('now')),

    UNIQUE (cohort_label, metric)
);

-- Record this migration as applied
INSERT INTO schema_migrations (version, description) VALUES (1, 'initial_schema');
