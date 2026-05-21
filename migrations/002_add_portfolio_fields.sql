-- Migration 002: Add portfolio-specific fields
-- Adds metric columns surfaced during the 12-company portfolio review.
-- Apply with: python scripts/migrate.py

-- kpi_snapshots additions
ALTER TABLE kpi_snapshots ADD COLUMN aum_usd                  REAL;    -- Assets Under Management (e.g. Cowrywise)
ALTER TABLE kpi_snapshots ADD COLUMN gmv_usd                  REAL;    -- Gross Merchandise/Payment Volume (e.g. Yoco, Verto, Enza, TWINCO, MaxSoko, OCTA)
ALTER TABLE kpi_snapshots ADD COLUMN active_clients_count      INTEGER; -- Active clients/users/merchants distinct from customer_count (e.g. Yoco, Verto, Khazna, MaxSoko)
ALTER TABLE kpi_snapshots ADD COLUMN par_30_pct               REAL;    -- Portfolio at Risk 30d as % of loan book (e.g. Lulalend, Khazna, TWINCO)
ALTER TABLE kpi_snapshots ADD COLUMN top_3_concentration_pct  REAL;    -- Revenue/volume concentration of top 3 customers % (e.g. TWINCO)
ALTER TABLE kpi_snapshots ADD COLUMN insurance_policies_active INTEGER; -- Active insurance policies outstanding (e.g. AllLife)
ALTER TABLE kpi_snapshots ADD COLUMN tpv_usd                  REAL;    -- Total Payment Volume where TPV differs from GMV (e.g. OCTA)
ALTER TABLE kpi_snapshots ADD COLUMN devices_connected         INTEGER; -- Connected IoT devices (e.g. Eseye)

-- companies addition
ALTER TABLE companies ADD COLUMN reporting_currency TEXT; -- Primary reporting currency before USD conversion (e.g. 'ZAR', 'NGN', 'GBP', 'USD', 'EUR')

-- Record this migration as applied
INSERT INTO schema_migrations (version, description) VALUES (2, '002_add_portfolio_fields');
