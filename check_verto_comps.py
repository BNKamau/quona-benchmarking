import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

# portfolio_comp_mapping uses comp_id (text), exit_comps PK is comp_id (text)
cur.execute("""
    SELECT e.company_name, e.is_clean_exit, e.exit_type, m.relevance_score
    FROM portfolio_comp_mapping m
    JOIN exit_comps e ON e.comp_id = m.comp_id
    WHERE m.portfolio_company IN ('Verto', 'VertoFX')
    ORDER BY e.is_clean_exit DESC, m.relevance_score DESC
""")
rows = cur.fetchall()
print("VertoFX comp mappings:")
for r in rows:
    print(f"  company={r[0]}, is_clean_exit={r[1]}, exit_type={r[2]}, relevance={r[3]}")

cur.close()
conn.close()
