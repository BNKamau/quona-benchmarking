import sys
sys.path.insert(0, '.')

TURSO_BENCHMARKING_URL   = "libsql://quona-benchmarking-bnkamau.aws-eu-west-1.turso.io"
TURSO_BENCHMARKING_TOKEN = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODA1NzUzMDMsImlkIjoiMDE5ZTkyOGMtZWYwMS03YmY3LTlmMzAtNjI3YmM5NmVmZTZhIiwicmlkIjoiYTJmOGVlZTQtNGExYS00MjVjLWFjMTItYmZjNDA3ZDJjZjk1In0.7JFZG5vxNoFhb0mO9i5CzuEV7GagPAqbsqR4fn1fqEtbjK2TJKPN_AmlLgzaDKnyZFnm37VfvGS31BQyU5vuDQ"

import requests

def turso_query(url, token, sql, args=None):
    resp = requests.post(
        f"{url.replace('libsql://', 'https://')}/v2/pipeline",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"requests": [
            {"type": "execute", "stmt": {"sql": sql, "args": args or []}},
            {"type": "close"}
        ]},
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()
    result = data["results"][0]
    if result.get("type") == "error":
        raise RuntimeError(f"SQL error: {result['error']['message']}")
    return result["response"]["result"]

print("=== Verifying Turso is live and writable ===\n")

# 1. Check current kpi_snapshots count
r = turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "SELECT COUNT(*) FROM kpi_snapshots")
count_before = int(r["rows"][0][0]["value"])
print(f"kpi_snapshots rows in Turso BEFORE test write: {count_before}")

# 2. Grab a real company_id to satisfy FK constraint
r = turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "SELECT id FROM companies ORDER BY id LIMIT 1")
test_company_id = int(r["rows"][0][0]["value"])
print(f"Using company_id={test_company_id} for test row")

# 3. Write a test row (float value must be a number, not a string)
turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "INSERT OR REPLACE INTO kpi_snapshots (company_id, period_end_date, reporting_currency, revenue_usd) VALUES (?, ?, ?, ?)",
    [{"type": "integer", "value": str(test_company_id)},
     {"type": "text",    "value": "2099-01-31"},
     {"type": "text",    "value": "USD"},
     {"type": "float",   "value": 1.0}])
print("Test row written to Turso.")

# 4. Read it back
r = turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "SELECT company_id, period_end_date, revenue_usd FROM kpi_snapshots WHERE period_end_date = '2099-01-31'")
rows = r["rows"]
print(f"Test row read back from Turso: {[[c['value'] for c in row] for row in rows]}")

# 5. Delete test row
turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "DELETE FROM kpi_snapshots WHERE period_end_date = '2099-01-31'")
print("Test row cleaned up.")

# 6. Confirm count unchanged
r = turso_query(TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN,
    "SELECT COUNT(*) FROM kpi_snapshots")
count_after = int(r["rows"][0][0]["value"])
print(f"kpi_snapshots rows in Turso AFTER cleanup: {count_after}")

print(f"\n{'='*50}")
if count_before == count_after and rows:
    print("TURSO CONNECTION VERIFIED — reads and writes are working correctly.")
    print("Data uploaded via the app will persist permanently.")
else:
    print("SOMETHING IS WRONG — investigate before re-uploading data.")
