"""
Run once against Supabase to create the exit_documents table.
Usage: python scripts/migrate_exit_documents.py

Reads SUPABASE_DB_URL from .streamlit/secrets.toml and connects via psycopg2.
If SUPABASE_DB_URL is absent the app is on local SQLite and _init_db()
already manages the table — nothing to do.
"""
import os
import sys
import tomllib
import psycopg2

SECRETS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), ".streamlit", "secrets.toml"
)

with open(SECRETS_PATH, "rb") as f:
    secrets = tomllib.load(f)

url = secrets.get("SUPABASE_DB_URL", "")
if not url:
    print("No SUPABASE_DB_URL found in .streamlit/secrets.toml.")
    print("App is using local SQLite — _init_db() manages the table automatically.")
    print("Nothing to do.")
    sys.exit(0)

conn = psycopg2.connect(url)
cur  = conn.cursor()
cur.execute("""
    CREATE TABLE IF NOT EXISTS exit_documents (
        id          SERIAL PRIMARY KEY,
        company_id  INTEGER NOT NULL,
        doc_name    TEXT    NOT NULL,
        doc_type    TEXT    NOT NULL DEFAULT 'exit_planning',
        file_data   BYTEA   NOT NULL,
        file_size   INTEGER,
        uploaded_by TEXT    DEFAULT '',
        uploaded_at TEXT    NOT NULL DEFAULT (now()::text),
        notes       TEXT    DEFAULT ''
    )
""")
conn.commit()
cur.close()
conn.close()
print("exit_documents table created (or already existed) in Supabase.")
