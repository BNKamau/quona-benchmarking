import psycopg2

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

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

conn = psycopg2.connect(SUPABASE_URL)
cur = conn.cursor()

cur.execute("SELECT name, fund FROM companies ORDER BY name")
print("Current fund values:")
for row in cur.fetchall():
    print(f"  {row[0]}: {row[1]}")

for company, fund in FUND_MAP.items():
    cur.execute("UPDATE companies SET fund = %s WHERE name = %s", (fund, company))
    print(f"Updated {company} -> {fund}")

conn.commit()

cur.execute("SELECT name, fund FROM companies ORDER BY name")
print("\nAfter update:")
for row in cur.fetchall():
    print(f"  {row[0]}: {row[1]}")

cur.close()
conn.close()
print("\nDone.")
