import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

# 1. See exactly what's in the table
cur.execute("SELECT id, name, sector, sub_sector, hq_country, founded_year, fund FROM companies ORDER BY name")
rows = cur.fetchall()
print("Current companies table:")
for r in rows:
    print(r)

# 2. Fix any column storing the string "None" — replace with NULL
cur.execute("UPDATE companies SET fund = NULL WHERE fund = 'None' OR fund = 'none'")
cur.execute("UPDATE companies SET sector = NULL WHERE sector = 'None' OR sector = 'none'")
cur.execute("UPDATE companies SET sub_sector = NULL WHERE sub_sector = 'None' OR sub_sector = 'none'")
cur.execute("UPDATE companies SET hq_country = NULL WHERE hq_country = 'None' OR hq_country = 'none'")
conn.commit()
print("\nNone strings replaced with NULL")

# 3. Now set correct fund values
FUND_MAP = {
    "Yoco":      "Fund I",
    "Cowrywise": "Fund II",
    "Lulalend":  "Fund II",
    "Verto":     "Fund II",
    "VertoFX":   "Fund II",
    "MaxSoko":   "Fund II",
    "Khazna":    "Fund III",
    "Enza":      "Fund III",
    "SAVA":      "Fund III",
    "TWINCO":    "Fund III",
}
for company, fund in FUND_MAP.items():
    cur.execute("UPDATE companies SET fund = %s WHERE name = %s", (fund, company))
conn.commit()
print("Fund values set")

# 4. Verify final state
cur.execute("SELECT name, sector, hq_country, fund FROM companies ORDER BY name")
print("\nFinal state:")
for r in cur.fetchall():
    print(f"  {r[0]}: sector={r[1]}, country={r[2]}, fund={r[3]}")

cur.close()
conn.close()
print("\nDone.")
