import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

# Remove pre-exit comps from VertoFX mapping
# Note: portfolio_company = 'Verto FX' (with space), comp_id is text (not integer id)
REMOVE = ["Flutterwave", "SumUp", "CloudWalk"]

for name in REMOVE:
    cur.execute("""
        DELETE FROM portfolio_comp_mapping
        WHERE portfolio_company IN ('Verto', 'VertoFX', 'Verto FX')
        AND comp_id = (SELECT comp_id FROM exit_comps WHERE company_name = %s)
    """, (name,))
    print(f"Removed {name} from VertoFX mapping (rows affected: {cur.rowcount})")

conn.commit()

# Verify remaining Verto comps
cur.execute("""
    SELECT e.company_name, e.is_clean_exit, m.relevance_score
    FROM portfolio_comp_mapping m
    JOIN exit_comps e ON e.comp_id = m.comp_id
    WHERE m.portfolio_company IN ('Verto', 'VertoFX', 'Verto FX')
    ORDER BY e.is_clean_exit DESC, m.relevance_score DESC
""")
print("\nRemaining VertoFX comps:")
for r in cur.fetchall():
    print(f"  {r[0]}: clean={r[1]}, relevance={r[2]}")

cur.close()
conn.close()
print("\nDone.")
