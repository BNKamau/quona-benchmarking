"""
Applies pending SQL migrations to the database.

Usage:
    python scripts/migrate.py                  # uses default db path
    python scripts/migrate.py path/to/db.sqlite

Migration files live in migrations/ and are named NNN_description.sql.
Already-applied migrations are skipped. Never edit an applied migration —
add a new numbered file instead.
"""

import os
import re
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
MIGRATIONS_DIR = ROOT / "migrations"
DEFAULT_DB = ROOT / "benchmarking.db"


def get_db_path() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    return DEFAULT_DB


def ensure_migrations_table(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS schema_migrations (
            version     INTEGER PRIMARY KEY,
            description TEXT    NOT NULL,
            applied_at  TEXT    NOT NULL DEFAULT (datetime('now'))
        )
    """)
    conn.commit()


def applied_versions(conn: sqlite3.Connection) -> set[int]:
    rows = conn.execute("SELECT version FROM schema_migrations").fetchall()
    return {row[0] for row in rows}


def pending_migrations(applied: set[int]) -> list[tuple[int, Path]]:
    pattern = re.compile(r"^(\d+)_.+\.sql$")
    files = []
    for path in sorted(MIGRATIONS_DIR.glob("*.sql")):
        m = pattern.match(path.name)
        if m:
            version = int(m.group(1))
            if version not in applied:
                files.append((version, path))
    return files


def apply(conn: sqlite3.Connection, version: int, path: Path) -> None:
    sql = path.read_text(encoding="utf-8")
    print(f"  Applying {path.name} ...", end=" ")
    conn.executescript(sql)
    # executescript issues an implicit COMMIT, so INSERT runs in its own tx
    conn.execute(
        "INSERT OR IGNORE INTO schema_migrations (version, description) VALUES (?, ?)",
        (version, path.stem),
    )
    conn.commit()
    print("done")


def main() -> None:
    db_path = get_db_path()
    print(f"Database: {db_path}")

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")

    ensure_migrations_table(conn)
    applied = applied_versions(conn)
    pending = pending_migrations(applied)

    if not pending:
        print("No pending migrations.")
        return

    print(f"{len(pending)} migration(s) to apply:")
    for version, path in pending:
        apply(conn, version, path)

    conn.close()
    print("All migrations applied.")


if __name__ == "__main__":
    main()
