import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

# Show current state
cur.execute("""
    SELECT e.company_name, e.is_clean_exit, e.exit_type, m.relevance_score
    FROM portfolio_comp_mapping m
    JOIN exit_comps e ON e.comp_id = m.comp_id
    WHERE m.portfolio_company = 'Lulalend'
    ORDER BY e.is_clean_exit DESC, m.relevance_score DESC
""")
print("Current Lulalend comps:")
for r in cur.fetchall():
    print(f"  {r[0]}: clean={r[1]}, type={r[2]}, relevance={r[3]}")

# Remove all pre-exit comps (is_clean_exit = 0 or NULL)
cur.execute("""
    DELETE FROM portfolio_comp_mapping
    WHERE portfolio_company = 'Lulalend'
    AND comp_id IN (
        SELECT comp_id FROM exit_comps WHERE is_clean_exit = 0 OR is_clean_exit IS NULL
    )
""")
print(f"\nRemoved {cur.rowcount} pre-exit comps from Lulalend mapping")
conn.commit()

# Verify
cur.execute("""
    SELECT e.company_name, e.is_clean_exit, m.relevance_score
    FROM portfolio_comp_mapping m
    JOIN exit_comps e ON e.comp_id = m.comp_id
    WHERE m.portfolio_company = 'Lulalend'
    ORDER BY m.relevance_score DESC
""")
print("\nRemaining Lulalend comps:")
for r in cur.fetchall():
    print(f"  {r[0]}: clean={r[1]}, relevance={r[2]}")

cur.close()
conn.close()
print("\nDone.")
