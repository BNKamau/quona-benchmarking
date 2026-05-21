-- Migration 003: Update companies schema
-- Adds sub_sector column and widens the business_model CHECK constraint
-- to accept 'b2b' and 'b2c' in addition to the original type values.
-- SQLite does not support ALTER TABLE … DROP/MODIFY CONSTRAINT,
-- so we recreate the table (safe here — no data yet).

PRAGMA foreign_keys = OFF;

CREATE TABLE companies_new (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    name           TEXT    NOT NULL,
    type           TEXT    NOT NULL CHECK (type IN ('portfolio', 'exit_comp')),
    sector         TEXT,
    sub_sector     TEXT,
    hq_country     TEXT,
    founded_year   INTEGER,
    business_model TEXT    CHECK (business_model IN (
                       'b2b', 'b2c', 'b2b2c',
                       'lending', 'saas', 'marketplace', 'payments', 'other'
                   )),
    reporting_currency TEXT,
    notes          TEXT,
    created_at     TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at     TEXT    NOT NULL DEFAULT (datetime('now'))
);

INSERT INTO companies_new
    SELECT id, name, type, sector, NULL, hq_country,
           founded_year, business_model, reporting_currency,
           notes, created_at, updated_at
    FROM companies;

DROP TABLE companies;
ALTER TABLE companies_new RENAME TO companies;

CREATE TRIGGER companies_updated_at
AFTER UPDATE ON companies
FOR EACH ROW BEGIN
    UPDATE companies SET updated_at = datetime('now') WHERE id = OLD.id;
END;

PRAGMA foreign_keys = ON;

INSERT INTO schema_migrations (version, description) VALUES (3, '003_update_companies_schema');
