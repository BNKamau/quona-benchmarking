# migrate_to_turso.py  —  migrates both local SQLite DBs to Turso via HTTP API
import sqlite3
import requests
import json
import os
import base64

# ── Config ────────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
LOCAL_BENCHMARKING = os.path.join(_HERE, "benchmarking.db")
LOCAL_COMPS        = os.path.join(_HERE, "data", "quona_exit_comps.db")

TURSO_BENCHMARKING_URL   = "https://quona-benchmarking-bnkamau.aws-eu-west-1.turso.io"
TURSO_BENCHMARKING_TOKEN = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODA1NzUzMDMsImlkIjoiMDE5ZTkyOGMtZWYwMS03YmY3LTlmMzAtNjI3YmM5NmVmZTZhIiwicmlkIjoiYTJmOGVlZTQtNGExYS00MjVjLWFjMTItYmZjNDA3ZDJjZjk1In0.7JFZG5vxNoFhb0mO9i5CzuEV7GagPAqbsqR4fn1fqEtbjK2TJKPN_AmlLgzaDKnyZFnm37VfvGS31BQyU5vuDQ"

TURSO_COMPS_URL   = "https://quona-exit-comps-bnkamau.aws-eu-west-1.turso.io"
TURSO_COMPS_TOKEN = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODA1NzU1OTEsImlkIjoiMDE5ZTkyOGQtZGEwMS03NTJlLWE4NjktZGNjNzNjNjBmYjVlIiwicmlkIjoiNDE5YjMzNTMtNmQ0Ny00ZTlkLTg3NzctODA3MzIxMDk2NTJiIn0.8pvd--6kEb1E_GocArR0VtUh3i6RLfqd4aE8i7R2wI_nZJmq6rg-yc3vxtolxl_Y2Di4oOhcoSUlzhFLa8LcAA"

BATCH_SIZE = 40  # statements per pipeline request (conservative)


# ── Turso HTTP helpers ────────────────────────────────────────────────────────

def _to_turso_arg(val):
    if val is None:
        return {"type": "null", "value": None}
    if isinstance(val, bool):
        return {"type": "integer", "value": "1" if val else "0"}
    if isinstance(val, int):
        return {"type": "integer", "value": str(val)}
    if isinstance(val, float):
        return {"type": "float", "value": val}
    if isinstance(val, bytes):
        return {"type": "blob", "base64": base64.b64encode(val).decode()}
    return {"type": "text", "value": str(val)}


def turso_pipeline(http_base_url, token, stmts):
    """POST a list of SQL statements as a pipeline. Returns results list."""
    requests_payload = [
        {"type": "execute", "stmt": s} for s in stmts
    ] + [{"type": "close"}]

    resp = requests.post(
        f"{http_base_url}/v2/pipeline",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"requests": requests_payload},
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    results = data.get("results", [])
    for r in results:
        if r.get("type") == "error":
            raise RuntimeError(f"Turso error: {r['error']['message']}")
    return results


def turso_exec(http_base_url, token, sql, args=None):
    stmt = {"sql": sql}
    if args:
        stmt["args"] = [_to_turso_arg(v) for v in args]
    return turso_pipeline(http_base_url, token, [stmt])


# ── Migration ─────────────────────────────────────────────────────────────────

def migrate_db(local_path, turso_url, turso_token, label):
    print(f"\n{'='*60}")
    print(f"Migrating: {label}")
    print(f"Local:     {local_path}")
    print(f"Turso:     {turso_url}")
    print(f"{'='*60}")

    local = sqlite3.connect(local_path)
    local.row_factory = sqlite3.Row

    tables = [
        r[0] for r in local.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()
    ]
    print(f"Tables found: {tables}")

    # Pass 1: create all tables first (respects FK declarations)
    # Sort so parent tables (companies) come before child tables
    TABLE_ORDER = ["schema_migrations", "companies", "portfolio_metadata",
                   "kpi_snapshots", "exit_comps", "funding_stage_snapshots",
                   "benchmarks", "exit_pathways", "buyer_tracking",
                   "quarterly_actions", "ipo_readiness"]
    ordered = [t for t in TABLE_ORDER if t in tables] + [t for t in tables if t not in TABLE_ORDER]

    for table in ordered:
        schema = local.execute(
            f"SELECT sql FROM sqlite_master WHERE type='table' AND name='{table}'"
        ).fetchone()[0]
        schema_safe = schema.replace("CREATE TABLE", "CREATE TABLE IF NOT EXISTS")
        turso_exec(turso_url, turso_token, schema_safe)

    # Pass 2: insert data table by table
    for table in ordered:
        rows = local.execute(f"SELECT * FROM {table}").fetchall()
        col_names = [d[0] for d in local.execute(f"SELECT * FROM {table} LIMIT 0").description]

        if not rows:
            print(f"  {table}: 0 rows — skipped")
            continue

        placeholders = ", ".join(["?" for _ in col_names])
        insert_sql   = f"INSERT OR REPLACE INTO {table} VALUES ({placeholders})"

        for i in range(0, len(rows), BATCH_SIZE):
            batch = [tuple(r) for r in rows[i:i + BATCH_SIZE]]
            stmts = [
                {"sql": insert_sql, "args": [_to_turso_arg(v) for v in row]}
                for row in batch
            ]
            turso_pipeline(turso_url, turso_token, stmts)

        # Verify
        count_result = turso_exec(turso_url, turso_token, f"SELECT COUNT(*) FROM {table}")
        turso_count = int(count_result[0]["response"]["result"]["rows"][0][0]["value"])
        local_count = len(rows)
        status = "OK" if turso_count >= local_count else "MISMATCH"
        print(f"  {table}: local={local_count} turso={turso_count} [{status}]")

    local.close()
    print(f"\n{label} migration complete.")


# ── Run both migrations ───────────────────────────────────────────────────────
migrate_db(LOCAL_BENCHMARKING, TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN, "benchmarking.db")
migrate_db(LOCAL_COMPS,        TURSO_COMPS_URL,        TURSO_COMPS_TOKEN,        "quona_exit_comps.db")

print("\n\nAll migrations complete. Verify row counts above before proceeding.")
