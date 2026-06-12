import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

COMPS = [
    {
        "comp_id": "QC-Demica",
        "company_name": "Demica",
        "sub_sector": "supply_chain_finance",
        "geography": "Europe", "exit_type": "Acquisition", "exit_year": 2024,
        "acquirer_exchange": "FIS", "revenue_at_exit_usd_m": None, "gross_margin_pct": None,
        "ebitda_margin_pct": None, "ev_revenue_multiple": None,
        "key_narrative_drivers": "Supply chain finance platform. $40B AuA, 40% CAGR. FIS acquisition validates institutional appetite for SCF platforms at $300M.",
        "data_confidence": "high", "data_source": "Public announcement Dec 2024", "is_clean_exit": 1,
    },
    {
        "comp_id": "QC-Taulia",
        "company_name": "Taulia",
        "sub_sector": "supply_chain_finance",
        "geography": "Global", "exit_type": "Acquisition", "exit_year": 2022,
        "acquirer_exchange": "SAP", "revenue_at_exit_usd_m": None, "gross_margin_pct": None,
        "ebitda_margin_pct": None, "ev_revenue_multiple": 17.0,
        "key_narrative_drivers": "$24M ARR at exit, ~17x ARR multiple, $500B+ processed annually. SAP acquisition for working capital and SCF capabilities.",
        "data_confidence": "high", "data_source": "Public announcement Mar 2022", "is_clean_exit": 1,
    },
    {
        "comp_id": "QC-C2FO",
        "company_name": "C2FO",
        "sub_sector": "supply_chain_finance",
        "geography": "Global", "exit_type": "Private", "exit_year": None,
        "acquirer_exchange": None, "revenue_at_exit_usd_m": 186.0, "gross_margin_pct": None,
        "ebitda_margin_pct": None, "ev_revenue_multiple": 5.0,
        "key_narrative_drivers": "$186M ARR (2025), ~5x ARR at last $1B valuation (2019). Dynamic discounting platform — lower bound for SCF multiples.",
        "data_confidence": "medium", "data_source": "PitchBook / public sources", "is_clean_exit": 0,
    },
    {
        "comp_id": "QC-Greensill",
        "company_name": "Greensill",
        "sub_sector": "supply_chain_finance",
        "geography": "Global", "exit_type": "Collapsed", "exit_year": 2021,
        "acquirer_exchange": None, "revenue_at_exit_usd_m": None, "gross_margin_pct": None,
        "ebitda_margin_pct": None, "ev_revenue_multiple": None,
        "key_narrative_drivers": "Cautionary — fraud and concentration risk. Peak $1.7B valuation before collapse. Key lesson: counterparty concentration and off-balance-sheet risk.",
        "data_confidence": "low", "data_source": "Public — collapsed Mar 2021", "is_clean_exit": 0,
    },
    {
        "comp_id": "QC-Stenn",
        "company_name": "Stenn",
        "sub_sector": "supply_chain_finance",
        "geography": "Global", "exit_type": "Collapsed", "exit_year": 2024,
        "acquirer_exchange": None, "revenue_at_exit_usd_m": None, "gross_margin_pct": None,
        "ebitda_margin_pct": None, "ev_revenue_multiple": None,
        "key_narrative_drivers": "Cautionary — HSBC fraud allegations Dec 2024. Peak $900M valuation. Key lesson: fraud risk in invoice finance; importance of audit controls.",
        "data_confidence": "low", "data_source": "Public — administration Dec 2024", "is_clean_exit": 0,
    },
]

MAPPINGS = {
    "QC-Demica":    (5, "Most direct SCF platform comp — FIS acquired for $300M. Twinco is the PO finance layer Demica lacks."),
    "QC-Taulia":    (5, "Best ARR multiple benchmark — 17x ARR at SAP acquisition. Sets the ceiling for Twinco's valuation range."),
    "QC-C2FO":      (3, "Dynamic discounting platform at 5x ARR — lower bound for SCF multiples. $1B valuation benchmark."),
    "QC-Greensill": (2, "Cautionary comp — concentration risk and fraud. Demonstrates what to avoid."),
    "QC-Stenn":     (2, "Cautionary comp — fraud allegations in invoice finance. Highlights importance of audit and controls."),
}

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

for comp in COMPS:
    comp_id = comp["comp_id"]
    cur.execute("DELETE FROM exit_comps WHERE comp_id = %s", (comp_id,))
    cur.execute("""
        INSERT INTO exit_comps (comp_id, company_name, sub_sector, geography, exit_type, exit_year,
        acquirer_exchange, revenue_at_exit_usd_m, gross_margin_pct, ebitda_margin_pct,
        ev_revenue_multiple, key_narrative_drivers, data_confidence, data_source, is_clean_exit)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, (comp_id, comp["company_name"], comp["sub_sector"], comp["geography"],
          comp["exit_type"], comp["exit_year"], comp["acquirer_exchange"],
          comp["revenue_at_exit_usd_m"], comp["gross_margin_pct"], comp["ebitda_margin_pct"],
          comp["ev_revenue_multiple"], comp["key_narrative_drivers"],
          comp["data_confidence"], comp["data_source"], comp["is_clean_exit"]))

    relevance, rationale = MAPPINGS[comp_id]
    cur.execute("DELETE FROM portfolio_comp_mapping WHERE portfolio_company = %s AND comp_id = %s",
                ("TWINCO", comp_id))
    cur.execute("""INSERT INTO portfolio_comp_mapping (portfolio_company, comp_id, relevance_score, mapping_rationale)
                   VALUES (%s, %s, %s, %s)""", ("TWINCO", comp_id, relevance, rationale))
    print(f"Inserted and mapped {comp['company_name']} (comp_id={comp_id}, relevance={relevance})")

conn.commit()
cur.close()
conn.close()
print("\nDone.")
