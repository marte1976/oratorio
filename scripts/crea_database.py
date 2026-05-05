from __future__ import annotations

import sqlite3
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "database" / "gestione_associazione.sqlite"
SCHEMA_PATH = BASE_DIR / "database" / "schema_associazione.sql"


def main() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    if DB_PATH.exists():
        DB_PATH.unlink()

    schema_sql = SCHEMA_PATH.read_text(encoding="utf-8")

    connection = sqlite3.connect(DB_PATH)
    try:
        connection.executescript(schema_sql)
        connection.commit()
    finally:
        connection.close()

    sys.path.insert(0, str(BASE_DIR))
    import app as gestionale_app

    gestionale_app.ensure_schema()

    print(f"Database creato: {DB_PATH}")


if __name__ == "__main__":
    main()
