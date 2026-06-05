# migrate_to_supabase.py
# Migrates all data from Turso (benchmarking + comps) to Supabase PostgreSQL
import requests
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "psycopg2-binary", "-q"])
import psycopg2
import psycopg2.extras

SUPABASE_URL = "postgresql://postgres.ijnrpconmtfxavccdenx:8Z3eHE3nq3w7RbwR@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"

TURSO_BENCHMARKING_URL   = "libsql://quona-benchmarking-bnkamau.aws-eu-west-1.turso.io"
TURSO_BENCHMARKING_TOKEN = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODA1NzUzMDMsImlkIjoiMDE5ZTkyOGMtZWYwMS03YmY3LTlmMzAtNjI3YmM5NmVmZTZhIiwicmlkIjoiYTJmOGVlZTQtNGExYS00MjVjLWFjMTItYmZjNDA3ZDJjZjk1In0.7JFZG5vxNoFhb0mO9i5CzuEV7GagPAqbsqR4fn1fqEtbjK2TJKPN_AmlLgzaDKnyZFnm37VfvGS31BQyU5vuDQ"

TURSO_COMPS_URL   = "libsql://quona-exit-comps-bnkamau.aws-eu-west-1.turso.io"
TURSO_COMPS_TOKEN = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODA1NzU1OTEsImlkIjoiMDE5ZTkyOGQtZGEwMS03NTJlLWE4NjktZGNjNzNjNjBmYjVlIiwicmlkIjoiNDE5YjMzNTMtNmQ0Ny00ZTlkLTg3NzctODA3MzIxMDk2NTJiIn0.8pvd--6kEb1E_GocArR0VtUh3i6RLfqd4aE8i7R2wI_nZJmq6rg-yc3vxtolxl_Y2Di4oOhcoSUlzhFLa8LcAA"

# Tables in the benchmarking DB that clash with comps DB table names (and are empty):
# skip them during the benchmarking migration so the comps DB owns those names.
BENCH_SKIP_TABLES = {"exit_comps"}


def turso_query(url, token, sql, args=None):
    resp = requests.post(
        url.replace("libsql://", "https://") + "/v2/pipeline",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"requests": [
            {"type": "execute", "stmt": {"sql": sql, "args": args or []}},
            {"type": "close"}
        ]},
        timeout=30,
    )
    resp.raise_for_status()
    result = resp.json()["results"][0]
    if result.get("type") == "error":
        raise Exception(result["error"])
    return result["response"]["result"]


def get_turso_tables(url, token):
    result = turso_query(url, token, "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    return [r[0]["value"] for r in result["rows"]]


def get_turso_data(url, token, table):
    result = turso_query(url, token, f"SELECT * FROM {table}")
    cols = [c["name"] for c in result["cols"]]
    rows = []
    for row in result["rows"]:
        r = {}
        for col, val in zip(cols, row):
            r[col] = None if val.get("type") == "null" else val.get("value")
        rows.append(r)
    return cols, rows


def sqlite_type_to_pg(sqlite_type):
    t = (sqlite_type or "").upper()
    if "INT" in t:                                    return "BIGINT"
    if "REAL" in t or "FLOAT" in t or "DOUBLE" in t: return "DOUBLE PRECISION"
    if "TEXT" in t or "CHAR" in t or "CLOB" in t:    return "TEXT"
    if "BLOB" in t:                                   return "BYTEA"
    if "BOOL" in t:                                   return "BOOLEAN"
    return "TEXT"


def translate_default(dflt):
    """Translate SQLite-specific default expressions to PostgreSQL equivalents."""
    if dflt is None:
        return None
    if dflt.strip().lower() in ("datetime('now')", "current_timestamp"):
        return "NOW()"
    return dflt


def get_turso_schema(url, token, table):
    result = turso_query(url, token, f"PRAGMA table_info({table})")
    cols = [c["name"] for c in result["cols"]]
    rows = []
    for row in result["rows"]:
        r = {col: (val.get("value") if val.get("type") != "null" else None)
             for col, val in zip(cols, row)}
        rows.append(r)
    return rows


def migrate_db(turso_url, turso_token, pg_conn, label, skip_tables=None):
    print(f"\n{'='*60}")
    print(f"Migrating: {label}")
    print(f"{'='*60}")
    cur = pg_conn.cursor()
    skip_tables = skip_tables or set()

    tables = get_turso_tables(turso_url, turso_token)
    print(f"Tables: {tables}")

    order = ["companies", "kpi_snapshots", "exit_pathways", "buyer_tracking",
             "quarterly_actions", "ipo_readiness", "exit_comps",
             "portfolio_comp_mapping", "sub_sector_benchmarks", "taxonomy"]
    tables_ordered = [t for t in order if t in tables] + [t for t in tables if t not in order]

    for table in tables_ordered:
        if table in skip_tables:
            print(f"  {table}: skipped (reserved for comps DB)")
            continue

        schema_rows = get_turso_schema(turso_url, turso_token, table)
        col_defs = []
        for row in schema_rows:
            name    = row["name"]
            pg_type = sqlite_type_to_pg(row["type"])
            is_pk   = row["pk"] == 1
            # Don't propagate NOT NULL to non-PK columns — SQLite is lenient
            # and production data often has NULLs in nominally non-null columns.
            dflt    = translate_default(row["dflt_value"])
            default = f"DEFAULT {dflt}" if dflt is not None else ""
            if is_pk and pg_type == "BIGINT":
                col_defs.append(f'"{name}" BIGSERIAL PRIMARY KEY')
            elif is_pk:
                col_defs.append(f'"{name}" {pg_type} PRIMARY KEY')
            else:
                parts = [f'"{name}"', pg_type, default]
                col_defs.append(" ".join(p for p in parts if p))

        # Drop and recreate to guarantee schema matches source
        try:
            cur.execute(f'DROP TABLE IF EXISTS "{table}" CASCADE')
            create_sql = f'CREATE TABLE "{table}" ({", ".join(col_defs)})'
            cur.execute(create_sql)
            pg_conn.commit()
        except Exception as e:
            pg_conn.rollback()
            print(f"  {table}: CREATE failed — {e}")
            continue

        cols, rows = get_turso_data(turso_url, turso_token, table)
        if not rows:
            print(f"  {table}: 0 rows — skipped")
            continue

        col_list   = ", ".join(f'"{c}"' for c in cols)
        insert_sql = f'INSERT INTO "{table}" ({col_list}) VALUES %s ON CONFLICT DO NOTHING'
        try:
            values = [tuple(r[c] for c in cols) for r in rows]
            psycopg2.extras.execute_values(cur, insert_sql, values, page_size=100)
            pg_conn.commit()
            cur.execute(f'SELECT COUNT(*) FROM "{table}"')
            pg_count = cur.fetchone()[0]
            status = "OK" if pg_count >= len(rows) else "MISMATCH"
            print(f"  {table}: {len(rows)} rows -> Supabase: {pg_count} [{status}]")
        except Exception as e:
            pg_conn.rollback()
            print(f"  {table}: INSERT failed — {e}")

    cur.close()


# ── Run migration ─────────────────────────────────────────────────────────────
print("Connecting to Supabase...")
pg = psycopg2.connect(SUPABASE_URL, connect_timeout=15)
pg.autocommit = False
print("Connected.")

migrate_db(
    TURSO_BENCHMARKING_URL, TURSO_BENCHMARKING_TOKEN, pg,
    "benchmarking DB",
    skip_tables=BENCH_SKIP_TABLES,
)
migrate_db(
    TURSO_COMPS_URL, TURSO_COMPS_TOKEN, pg,
    "exit comps DB",
)

pg.close()
print("\n\nMigration complete. Verify row counts above before proceeding.")
