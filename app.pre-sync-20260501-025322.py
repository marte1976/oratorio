from __future__ import annotations

import html
import hashlib
import hmac
import json
import mimetypes
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import uuid
import unicodedata
from calendar import monthrange
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime, timedelta
from functools import lru_cache
from http.cookies import SimpleCookie
from io import BytesIO
from pathlib import Path
from socketserver import ThreadingMixIn
from urllib.parse import parse_qs, quote, urlencode, urlsplit, urlunsplit
from wsgiref.simple_server import WSGIServer, make_server


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = BASE_DIR / "database" / "gestione_associazione.sqlite"
DB_PATH = Path(os.environ.get("ASSOCIAZIONE_DB_PATH", str(DEFAULT_DB_PATH)))
SCHEMA_PATH = BASE_DIR / "database" / "schema_associazione.sql"
STATIC_DIR = BASE_DIR / "static"
COMUNI_JSON_PATH = BASE_DIR / "data" / "comuni.json"
LOGO_URL = "/static/logo-ca.jpg"
APP_NAME = "Oratorio Carlo Acutis"
APP_SUBTITLE = ""
ESTATE_LABEL = "Campo estivo"
OUTPUT_DIR = BASE_DIR / "outputs"
ACTIVITY_LOG_XLS_PATH = OUTPUT_DIR / "registro_attivita.xls"
ACTIVITY_LOG_RETENTION_YEARS = 2
EXPORT_SCRIPT = BASE_DIR / "scripts" / "export_report.mjs"
LOCAL_RUNTIME_ROOT = BASE_DIR / "runtime"
LOCAL_NODE_BIN = LOCAL_RUNTIME_ROOT / "node" / "bin" / "node.exe"
LOCAL_ARTIFACT_TOOL_MODULE = (
    LOCAL_RUNTIME_ROOT
    / "node"
    / "node_modules"
    / "@oai"
    / "artifact-tool"
    / "dist"
    / "artifact_tool.mjs"
)
DEFAULT_NODE_BIN = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "node"
    / "bin"
    / "node.exe"
)
DEFAULT_ARTIFACT_TOOL_MODULE = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "node"
    / "node_modules"
    / "@oai"
    / "artifact-tool"
    / "dist"
    / "artifact_tool.mjs"
)
NODE_BIN = Path(
    os.environ.get(
        "ASSOCIAZIONE_NODE_BIN",
        str(LOCAL_NODE_BIN if LOCAL_NODE_BIN.exists() else DEFAULT_NODE_BIN),
    )
)
ARTIFACT_TOOL_MODULE = Path(
    os.environ.get(
        "ASSOCIAZIONE_ARTIFACT_TOOL_MODULE",
        str(LOCAL_ARTIFACT_TOOL_MODULE if LOCAL_ARTIFACT_TOOL_MODULE.exists() else DEFAULT_ARTIFACT_TOOL_MODULE),
    )
)

PROGRESSIVE_ENTITIES = {
    "associati": {
        "table": "associati",
        "column": "numero_progressivo",
        "code_column": "codice_associato",
        "prefix": "ASS",
        "label": "Associato",
    },
    "corsi": {
        "table": "corsi",
        "column": "numero_progressivo",
        "code_column": "codice_corso",
        "prefix": "COR",
        "label": "Corso",
    },
    "campi_estivi": {
        "table": "campi_estivi",
        "column": "numero_progressivo",
        "code_column": "codice_campo",
        "prefix": "CES",
        "label": ESTATE_LABEL,
    },
    "eventi": {
        "table": "eventi",
        "column": "numero_progressivo",
        "code_column": "codice_evento",
        "prefix": "EVT",
        "label": "Evento",
    },
}

LOCKED_CRUD_ENTITIES = {
    "tipologie_corsi": "Le tipologie corsi sono fisse e non possono essere modificate.",
    "campi_estivi": "Il Campo estivo annuale non puo essere modificato o eliminato da questa area.",
}

MONTH_NAMES = {
    1: "Gennaio",
    2: "Febbraio",
    3: "Marzo",
    4: "Aprile",
    5: "Maggio",
    6: "Giugno",
    7: "Luglio",
    8: "Agosto",
    9: "Settembre",
    10: "Ottobre",
    11: "Novembre",
    12: "Dicembre",
}

CARICA_VALUES = (
    "Associato",
    "Presidente",
    "Vice Presidente",
    "Segretario",
    "Tesoriere",
    "Consigliere",
    "Consigliere spirituale",
)
CONSIGLIO_DIRETTIVO_ORDER = (
    "Presidente",
    "Vice Presidente",
    "Tesoriere",
    "Segretario",
    "Consigliere",
    "Consigliere spirituale",
)
CONSIGLIO_DIRETTIVO_LABELS = {
    "Presidente": "Presidente",
    "Vice Presidente": "Vice Presidente",
    "Tesoriere": "Tesoriere",
    "Segretario": "Segretario",
    "Consigliere": "Consigliere",
    "Consigliere spirituale": "Consigliere Spirituale",
}

REPORT_RECIPIENT_CARICHE = CARICA_VALUES[1:]
SESSION_COOKIE_NAME = "oratorio_session"
SESSION_DURATION_DAYS = 30
PASSWORD_HASH_ITERATIONS = 210000
ADMIN_ONLY_REPORT_KEYS = {"registro-attivita"}

CODICE_FISCALE_MONTHS = {
    "A": 1,
    "B": 2,
    "C": 3,
    "D": 4,
    "E": 5,
    "H": 6,
    "L": 7,
    "M": 8,
    "P": 9,
    "R": 10,
    "S": 11,
    "T": 12,
}

CODICE_FISCALE_MONTH_CODES = {value: key for key, value in CODICE_FISCALE_MONTHS.items()}

CODICE_FISCALE_ODD_MAP = {
    "0": 1,
    "1": 0,
    "2": 5,
    "3": 7,
    "4": 9,
    "5": 13,
    "6": 15,
    "7": 17,
    "8": 19,
    "9": 21,
    "A": 1,
    "B": 0,
    "C": 5,
    "D": 7,
    "E": 9,
    "F": 13,
    "G": 15,
    "H": 17,
    "I": 19,
    "J": 21,
    "K": 2,
    "L": 4,
    "M": 18,
    "N": 20,
    "O": 11,
    "P": 3,
    "Q": 6,
    "R": 8,
    "S": 12,
    "T": 14,
    "U": 16,
    "V": 10,
    "W": 22,
    "X": 25,
    "Y": 24,
    "Z": 23,
}

CODICE_FISCALE_EVEN_MAP = {
    **{str(number): number for number in range(10)},
    **{chr(ord("A") + index): index for index in range(26)},
}

CODICE_FISCALE_CONTROL_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def associato_base_name_sql(alias: str = "a") -> str:
    prefix = f"{alias}." if alias else ""
    return f"trim(COALESCE({prefix}cognome, '') || ' ' || COALESCE({prefix}nome, ''))"


def associato_display_sql(alias: str = "a") -> str:
    prefix = f"{alias}." if alias else ""
    base_name = associato_base_name_sql(alias)
    return (
        f"{base_name} || CASE "
        f"WHEN COALESCE({prefix}data_nascita, '') <> '' AND julianday({prefix}data_nascita) IS NOT NULL "
        f"THEN ' (' || CAST((julianday('now') - julianday({prefix}data_nascita)) / 365.2425 AS INTEGER) || ' anni)' "
        f"ELSE '' END"
    )


REPORTING_VIEWS_SQL = f"""
DROP VIEW IF EXISTS v_incassi_totali;
DROP VIEW IF EXISTS v_riepilogo_associati;
DROP VIEW IF EXISTS v_scadenze_da_incassare;
DROP VIEW IF EXISTS v_eventi_saldo;
DROP VIEW IF EXISTS v_campi_estivi_saldo;
DROP VIEW IF EXISTS v_rate_corsi_saldo;
DROP VIEW IF EXISTS v_iscrizioni_corsi_saldo;
DROP VIEW IF EXISTS v_tesseramenti_saldo;

CREATE VIEW v_tesseramenti_saldo AS
SELECT
    t.id,
    t.anno_sociale,
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    t.data_tesseramento,
    t.data_scadenza,
    t.importo_dovuto,
    COALESCE(SUM(pt.importo), 0) AS importo_pagato,
    t.importo_dovuto - COALESCE(SUM(pt.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pt.importo), 0) >= t.importo_dovuto THEN 'Pagato'
        WHEN COALESCE(SUM(pt.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM tesseramenti_annuali t
JOIN associati a ON a.id = t.associato_id
LEFT JOIN pagamenti_tesseramenti pt ON pt.tesseramento_id = t.id
GROUP BY t.id;

CREATE VIEW v_iscrizioni_corsi_saldo AS
SELECT
    ic.id,
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    tc.codice_tipologia,
    COALESCE(tc.nome, 'Senza tipologia') AS tipologia_corso,
    c.codice_corso,
    c.nome AS corso,
    ic.data_iscrizione,
    ic.data_inizio,
    ic.data_fine,
    ic.stato_iscrizione,
    ic.quota_iscrizione,
    COALESCE(SUM(pic.importo), 0) AS importo_pagato_iscrizione,
    ic.quota_iscrizione - COALESCE(SUM(pic.importo), 0) AS saldo_iscrizione,
    CASE
        WHEN COALESCE(SUM(pic.importo), 0) >= ic.quota_iscrizione THEN 'Pagato'
        WHEN COALESCE(SUM(pic.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento_iscrizione,
    ic.quota_mensile
FROM iscrizioni_corsi ic
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN tipologie_corsi tc ON tc.id = c.tipologia_corso_id
LEFT JOIN pagamenti_iscrizioni_corsi pic ON pic.iscrizione_corso_id = ic.id
GROUP BY ic.id;

CREATE VIEW v_rate_corsi_saldo AS
SELECT
    r.id,
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    tc.codice_tipologia,
    COALESCE(tc.nome, 'Senza tipologia') AS tipologia_corso,
    c.codice_corso,
    c.nome AS corso,
    r.anno,
    r.mese,
    printf('%04d-%02d', r.anno, r.mese) AS competenza,
    r.data_scadenza,
    r.importo_dovuto,
    COALESCE(SUM(prc.importo), 0) AS importo_pagato,
    r.importo_dovuto - COALESCE(SUM(prc.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(prc.importo), 0) >= r.importo_dovuto THEN 'Pagato'
        WHEN COALESCE(SUM(prc.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM rate_corsi_mensili r
JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN tipologie_corsi tc ON tc.id = c.tipologia_corso_id
LEFT JOIN pagamenti_rate_corsi prc ON prc.rata_corso_id = r.id
GROUP BY r.id;

CREATE VIEW v_campi_estivi_saldo AS
SELECT
    ice.id,
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    ce.codice_campo,
    ce.nome AS campo_estivo,
    ce.anno,
    ce.data_inizio,
    ce.data_fine,
    ice.data_iscrizione,
    ice.stato_iscrizione,
    ice.quota_partecipazione AS importo_dovuto,
    COALESCE(SUM(pce.importo), 0) AS importo_pagato,
    ice.quota_partecipazione - COALESCE(SUM(pce.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pce.importo), 0) >= ice.quota_partecipazione THEN 'Pagato'
        WHEN COALESCE(SUM(pce.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM iscrizioni_campi_estivi ice
JOIN associati a ON a.id = ice.associato_id
JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
LEFT JOIN pagamenti_campi_estivi pce ON pce.iscrizione_campo_id = ice.id
GROUP BY ice.id;

CREATE VIEW v_eventi_saldo AS
SELECT
    ie.id,
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    e.codice_evento,
    e.nome AS evento,
    e.tipologia,
    e.data_evento,
    ie.data_iscrizione,
    ie.stato_iscrizione,
    ie.quota_partecipazione AS importo_dovuto,
    COALESCE(SUM(pe.importo), 0) AS importo_pagato,
    ie.quota_partecipazione - COALESCE(SUM(pe.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pe.importo), 0) >= ie.quota_partecipazione THEN 'Pagato'
        WHEN COALESCE(SUM(pe.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM iscrizioni_eventi ie
JOIN associati a ON a.id = ie.associato_id
JOIN eventi e ON e.id = ie.evento_id
LEFT JOIN pagamenti_eventi pe ON pe.iscrizione_evento_id = ie.id
GROUP BY ie.id;

CREATE VIEW v_scadenze_da_incassare AS
SELECT
    associato_id,
    codice_associato,
    associato,
    data_nascita,
    'Tesseramento annuale' AS area,
    'Anno ' || anno_sociale AS riferimento,
    data_scadenza AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_tesseramenti_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    data_nascita,
    'Corso - iscrizione' AS area,
    corso AS riferimento,
    data_iscrizione AS scadenza,
    quota_iscrizione AS importo_dovuto,
    importo_pagato_iscrizione AS importo_pagato,
    saldo_iscrizione AS saldo_residuo,
    stato_pagamento_iscrizione AS stato_pagamento
FROM v_iscrizioni_corsi_saldo
WHERE saldo_iscrizione > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    data_nascita,
    'Corso - quota mensile' AS area,
    corso || ' ' || competenza AS riferimento,
    data_scadenza AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_rate_corsi_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    data_nascita,
    'Campo estivo' AS area,
    campo_estivo AS riferimento,
    data_inizio AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_campi_estivi_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    data_nascita,
    'Evento' AS area,
    evento AS riferimento,
    data_evento AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_eventi_saldo
WHERE saldo_residuo > 0;

CREATE VIEW v_riepilogo_associati AS
WITH movimenti AS (
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_tesseramenti_saldo
    UNION ALL
    SELECT associato_id, quota_iscrizione AS importo_dovuto, importo_pagato_iscrizione AS importo_pagato
    FROM v_iscrizioni_corsi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_rate_corsi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_campi_estivi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_eventi_saldo
)
SELECT
    a.id AS associato_id,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    a.stato_associato,
    COALESCE(SUM(m.importo_dovuto), 0) AS totale_dovuto,
    COALESCE(SUM(m.importo_pagato), 0) AS totale_pagato,
    COALESCE(SUM(m.importo_dovuto), 0) - COALESCE(SUM(m.importo_pagato), 0) AS saldo_residuo
FROM associati a
LEFT JOIN movimenti m ON m.associato_id = a.id
GROUP BY a.id;

CREATE VIEW v_incassi_totali AS
SELECT
    'Tesseramento annuale' AS area,
    pt.data_pagamento,
    pt.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    'Anno ' || t.anno_sociale AS riferimento
FROM pagamenti_tesseramenti pt
JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
JOIN associati a ON a.id = t.associato_id
LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id

UNION ALL

SELECT
    'Corso - iscrizione' AS area,
    pic.data_pagamento,
    pic.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    c.nome AS riferimento
FROM pagamenti_iscrizioni_corsi pic
JOIN iscrizioni_corsi ic ON ic.id = pic.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN metodi_pagamento mp ON mp.id = pic.metodo_pagamento_id

UNION ALL

SELECT
    'Corso - quota mensile' AS area,
    prc.data_pagamento,
    prc.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS riferimento
FROM pagamenti_rate_corsi prc
JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id

UNION ALL

SELECT
    'Campo estivo' AS area,
    pce.data_pagamento,
    pce.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    ce.nome AS riferimento
FROM pagamenti_campi_estivi pce
JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
JOIN associati a ON a.id = ice.associato_id
JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id

UNION ALL

SELECT
    'Evento' AS area,
    pe.data_pagamento,
    pe.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    {associato_display_sql('a')} AS associato,
    a.data_nascita,
    e.nome AS riferimento
FROM pagamenti_eventi pe
JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
JOIN associati a ON a.id = ie.associato_id
JOIN eventi e ON e.id = ie.evento_id
LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id;
"""


def rebuild_reporting_views(connection: sqlite3.Connection) -> None:
    connection.executescript(REPORTING_VIEWS_SQL)


def initialize_database_if_missing() -> None:
    if DB_PATH.exists() or not SCHEMA_PATH.exists():
        return

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    try:
        connection.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
        connection.commit()
    finally:
        connection.close()


def table_columns(connection: sqlite3.Connection, table_name: str) -> set[str]:
    rows = connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    return {row[1] for row in rows}


def ensure_column(connection: sqlite3.Connection, table_name: str, column_sql: str, column_name: str) -> None:
    if column_name in table_columns(connection, table_name):
        return
    connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_sql}")


def cleanup_expired_sessions(connection: sqlite3.Connection) -> None:
    connection.execute("DELETE FROM sessioni_accesso WHERE scade_il <= ?", (current_timestamp(),))


def access_users_count() -> int:
    return int(fetch_scalar("SELECT COUNT(*) FROM utenti_accesso") or 0)


def bootstrap_admin_required() -> bool:
    return access_users_count() == 0


def create_access_user(
    connection: sqlite3.Connection,
    *,
    username: str,
    password: str,
    is_admin: bool,
    email_recupero: str | None = None,
) -> int:
    password_salt, password_hash, iterations = hash_password(password)
    cursor = connection.execute(
        """
        INSERT INTO utenti_accesso (
            username, password_hash, password_salt, password_iterations, is_admin, email_recupero
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            validate_username(username),
            password_hash,
            password_salt,
            iterations,
            1 if is_admin else 0,
            (email_recupero or "").strip() or None,
        ),
    )
    return int(cursor.lastrowid)


def set_access_user_password(connection: sqlite3.Connection, user_id: int, password: str) -> None:
    password_salt, password_hash, iterations = hash_password(password)
    connection.execute(
        """
        UPDATE utenti_accesso
        SET password_hash = ?, password_salt = ?, password_iterations = ?, aggiornato_il = ?
        WHERE id = ?
        """,
        (password_hash, password_salt, iterations, current_timestamp(), user_id),
    )


def access_user_row(user_id: int) -> sqlite3.Row | None:
    return fetch_one(
        """
        SELECT
            id,
            username,
            email_recupero,
            is_admin,
            attivo,
            creato_il,
            aggiornato_il,
            COALESCE(ultimo_accesso, '') AS ultimo_accesso
        FROM utenti_accesso
        WHERE id = ?
        """,
        (user_id,),
    )


def access_user_by_username(username: str) -> sqlite3.Row | None:
    return fetch_one(
        """
        SELECT *
        FROM utenti_accesso
        WHERE username = ?
        """,
        (validate_username(username),),
    )


def validate_recovery_email(email_value: str) -> str:
    email = (email_value or "").strip()
    if not email or "@" not in email:
        raise ValueError("Indica un'email di recupero valida.")
    return email


def create_user_session(connection: sqlite3.Connection, user_id: int) -> str:
    cleanup_expired_sessions(connection)
    token = secrets.token_urlsafe(32)
    expires_at = (datetime.now() + timedelta(days=SESSION_DURATION_DAYS)).replace(microsecond=0).isoformat(sep=" ")
    now_value = current_timestamp()
    connection.execute(
        """
        INSERT INTO sessioni_accesso (session_token, utente_id, scade_il)
        VALUES (?, ?, ?)
        """,
        (token, user_id, expires_at),
    )
    connection.execute(
        """
        UPDATE utenti_accesso
        SET ultimo_accesso = ?, aggiornato_il = ?
        WHERE id = ?
        """,
        (now_value, now_value, user_id),
    )
    return token


def current_user_from_cookies(cookies: dict[str, str]) -> dict[str, object] | None:
    token = cookies.get(SESSION_COOKIE_NAME, "")
    if not token:
        return None

    with get_connection() as connection:
        cleanup_expired_sessions(connection)
        row = connection.execute(
            """
            SELECT
                u.id,
                u.username,
                u.is_admin,
                u.attivo
            FROM sessioni_accesso s
            JOIN utenti_accesso u ON u.id = s.utente_id
            WHERE s.session_token = ?
              AND s.scade_il > ?
            """,
            (token, current_timestamp()),
        ).fetchone()
        if row is None:
            connection.commit()
            return None
        if not row["attivo"]:
            connection.execute("DELETE FROM sessioni_accesso WHERE session_token = ?", (token,))
            connection.commit()
            return None
        connection.commit()
        return {
            "id": int(row["id"]),
            "username": row["username"],
            "is_admin": bool(row["is_admin"]),
        }


def login_destination(target: str | None) -> str:
    candidate = (target or "").strip()
    if not candidate.startswith("/") or candidate.startswith("//"):
        return "/"
    return candidate


def format_progressive_code(entity_key: str, number: int) -> str:
    prefix = PROGRESSIVE_ENTITIES[entity_key]["prefix"]
    return f"{prefix}-{number:04d}"


def backfill_progressive_numbers(connection: sqlite3.Connection, entity_key: str) -> None:
    meta = PROGRESSIVE_ENTITIES[entity_key]
    rows = connection.execute(
        f"""
        SELECT id, {meta["column"]}, {meta["code_column"]}
        FROM {meta["table"]}
        ORDER BY id
        """
    ).fetchall()

    used_numbers: set[int] = set()
    next_number = 1
    updates: list[tuple[int, str, int]] = []

    for row in rows:
        current_number = row[1]
        current_code = row[2]
        if isinstance(current_number, int) and current_number > 0 and current_number not in used_numbers:
            number = current_number
        else:
            while next_number in used_numbers:
                next_number += 1
            number = next_number
            next_number += 1

        used_numbers.add(number)
        if current_number != number or not current_code:
            code = current_code or format_progressive_code(entity_key, number)
            updates.append((number, code, row[0]))

    if updates:
        connection.executemany(
            f"""
            UPDATE {meta["table"]}
            SET {meta["column"]} = ?, {meta["code_column"]} = COALESCE(NULLIF({meta["code_column"]}, ''), ?)
            WHERE id = ?
            """,
            updates,
        )

    ultimo_valore = max(used_numbers, default=0)
    connection.execute(
        """
        INSERT INTO sequenze_progressive (chiave, ultimo_valore)
        VALUES (?, ?)
        ON CONFLICT(chiave) DO UPDATE SET ultimo_valore = excluded.ultimo_valore
        """,
        (entity_key, ultimo_valore),
    )


def ensure_schema() -> None:
    initialize_database_if_missing()
    with get_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS sequenze_progressive (
                chiave TEXT PRIMARY KEY,
                ultimo_valore INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS utenti_accesso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                password_iterations INTEGER NOT NULL DEFAULT 210000,
                email_recupero TEXT,
                is_admin INTEGER NOT NULL DEFAULT 0 CHECK (is_admin IN (0, 1)),
                attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
                creato_il TEXT NOT NULL DEFAULT (datetime('now')),
                aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
                ultimo_accesso TEXT
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS sessioni_accesso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_token TEXT NOT NULL UNIQUE,
                utente_id INTEGER NOT NULL,
                creata_il TEXT NOT NULL DEFAULT (datetime('now')),
                scade_il TEXT NOT NULL,
                FOREIGN KEY (utente_id) REFERENCES utenti_accesso (id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS impostazioni_app (
                chiave TEXT PRIMARY KEY,
                valore TEXT,
                aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS registro_attivita (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_ora TEXT NOT NULL DEFAULT (datetime('now')),
                username TEXT,
                associato_id INTEGER,
                associato_codice TEXT,
                associato_nominativo TEXT,
                nome_pc TEXT,
                ip_client TEXT,
                metodo_http TEXT,
                percorso TEXT,
                categoria TEXT,
                descrizione_attivita TEXT NOT NULL,
                dettaglio TEXT,
                esito TEXT,
                anno_lavoro INTEGER,
                user_agent TEXT
            )
            """
        )
        for meta in PROGRESSIVE_ENTITIES.values():
            ensure_column(
                connection,
                meta["table"],
                f'{meta["column"]} INTEGER',
                meta["column"],
            )

        ensure_column(
            connection,
            "pagamenti_rate_corsi",
            "gruppo_ricevuta TEXT",
            "gruppo_ricevuta",
        )
        ensure_column(
            connection,
            "pagamenti_tesseramenti",
            "gruppo_ricevuta TEXT",
            "gruppo_ricevuta",
        )
        ensure_column(
            connection,
            "pagamenti_campi_estivi",
            "gruppo_ricevuta TEXT",
            "gruppo_ricevuta",
        )
        ensure_column(
            connection,
            "pagamenti_eventi",
            "gruppo_ricevuta TEXT",
            "gruppo_ricevuta",
        )
        ensure_column(
            connection,
            "associati",
            "comune_nascita TEXT",
            "comune_nascita",
        )
        ensure_column(
            connection,
            "associati",
            "provincia_nascita TEXT",
            "provincia_nascita",
        )
        ensure_column(
            connection,
            "associati",
            "sesso TEXT NOT NULL DEFAULT 'M'",
            "sesso",
        )
        ensure_column(
            connection,
            "associati",
            "carica TEXT NOT NULL DEFAULT 'Associato'",
            "carica",
        )
        ensure_column(
            connection,
            "utenti_accesso",
            "email_recupero TEXT",
            "email_recupero",
        )
        ensure_column(
            connection,
            "registro_attivita",
            "associato_id INTEGER",
            "associato_id",
        )
        ensure_column(
            connection,
            "registro_attivita",
            "associato_codice TEXT",
            "associato_codice",
        )
        ensure_column(
            connection,
            "registro_attivita",
            "associato_nominativo TEXT",
            "associato_nominativo",
        )

        for entity_key, meta in PROGRESSIVE_ENTITIES.items():
            backfill_progressive_numbers(connection, entity_key)
            connection.execute(
                f"""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_{meta["table"]}_{meta["column"]}
                ON {meta["table"]} ({meta["column"]})
                """
            )

        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_pagamenti_rate_corsi_gruppo
            ON pagamenti_rate_corsi (gruppo_ricevuta)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_pagamenti_tesseramenti_gruppo
            ON pagamenti_tesseramenti (gruppo_ricevuta)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_pagamenti_campi_estivi_gruppo
            ON pagamenti_campi_estivi (gruppo_ricevuta)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_pagamenti_eventi_gruppo
            ON pagamenti_eventi (gruppo_ricevuta)
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS quote_predefinite (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                area TEXT NOT NULL CHECK (area IN ('tesseramenti', 'campi-estivi')),
                descrizione TEXT NOT NULL,
                importo NUMERIC NOT NULL CHECK (importo >= 0),
                attiva INTEGER NOT NULL DEFAULT 1 CHECK (attiva IN (0, 1)),
                note TEXT,
                creato_il TEXT NOT NULL DEFAULT (datetime('now')),
                aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_quote_predefinite_area
            ON quote_predefinite (area, attiva, descrizione)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_utenti_accesso_admin_attivo
            ON utenti_accesso (is_admin, attivo, username)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sessioni_accesso_utente
            ON sessioni_accesso (utente_id, scade_il)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_registro_attivita_data
            ON registro_attivita (data_ora DESC, id DESC)
            """
        )
        prune_old_activity_log_rows(connection)
        cleanup_expired_sessions(connection)
        rebuild_reporting_views(connection)
        connection.commit()
    export_activity_log_xls()


def reserve_progressive_number(connection: sqlite3.Connection, entity_key: str) -> int:
    row = connection.execute(
        "SELECT ultimo_valore FROM sequenze_progressive WHERE chiave = ?",
        (entity_key,),
    ).fetchone()
    current_value = int(row[0]) if row else 0
    next_value = current_value + 1
    connection.execute(
        """
        INSERT INTO sequenze_progressive (chiave, ultimo_valore)
        VALUES (?, ?)
        ON CONFLICT(chiave) DO UPDATE SET ultimo_valore = excluded.ultimo_valore
        """,
        (entity_key, next_value),
    )
    return next_value


def peek_next_progressive_code(entity_key: str) -> str:
    row = fetch_one(
        "SELECT ultimo_valore FROM sequenze_progressive WHERE chiave = ?",
        (entity_key,),
    )
    next_value = (int(row[0]) if row else 0) + 1
    return format_progressive_code(entity_key, next_value)


def generate_receipt_group_code() -> str:
    return f"GRP-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6].upper()}"


def generate_multi_area_group_code() -> str:
    return f"MGR-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6].upper()}"


def current_work_year(query_params: dict[str, str]) -> int:
    raw_value = (query_params.get("anno_lavoro") or "").strip()
    if raw_value.isdigit():
        return int(raw_value)
    return date.today().year


def work_year_query(query_params: dict[str, str]) -> dict[str, str]:
    return {"anno_lavoro": str(current_work_year(query_params))}


def work_year_query_from_form(form_data: dict[str, str]) -> dict[str, str]:
    params: dict[str, str] = {}
    raw_value = normalized(form_data, "anno_lavoro", "")
    if raw_value.isdigit():
        params["anno_lavoro"] = raw_value
    if normalized(form_data, "vista", "") == "dati":
        params["vista"] = "dati"
    return params


def year_start_end(year: int) -> tuple[str, str]:
    return f"{year}-01-01", f"{year}-12-31"


def ensure_output_dir() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
    row = connection.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type = 'table' AND name = ?
        """,
        (table_name,),
    ).fetchone()
    return row is not None


def report_requires_admin(report_key: str) -> bool:
    return report_key in ADMIN_ONLY_REPORT_KEYS


def latest_generated_course_month(work_year: int) -> tuple[int, int] | None:
    row = fetch_one(
        """
        SELECT anno, mese
        FROM rate_corsi_mensili
        WHERE anno = ?
        ORDER BY anno DESC, mese DESC
        LIMIT 1
        """,
        (work_year,),
    )
    if row is None:
        return None
    return int(row["anno"]), int(row["mese"])


def increment_year_month(year: int, month: int) -> tuple[int, int]:
    if month >= 12:
        return year + 1, 1
    return year, month + 1


def parse_year_month_value(value: str, field_label: str = "Competenza") -> tuple[int, int]:
    raw_value = (value or "").strip()
    match = re.fullmatch(r"(\d{4})-(\d{2})", raw_value)
    if not match:
        raise ValueError(f"{field_label} non valida.")
    year = int(match.group(1))
    month = int(match.group(2))
    if month < 1 or month > 12:
        raise ValueError(f"{field_label} non valida.")
    return year, month


def default_mass_rate_due_date(year: int, month: int) -> str:
    next_year, next_month = increment_year_month(year, month)
    return f"{next_year:04d}-{next_month:02d}-14"


def pending_mass_generation_cutoff(today_value: date | None = None) -> tuple[int, int] | None:
    today_value = today_value or date.today()
    if today_value.day > 14:
        return today_value.year, today_value.month
    if today_value.month == 1:
        return today_value.year - 1, 12
    return today_value.year, today_value.month - 1


def pending_course_monthly_generations(
    work_year: int,
    today_value: date | None = None,
) -> dict[str, object] | None:
    today_value = today_value or date.today()
    if work_year != today_value.year:
        return None

    cutoff = pending_mass_generation_cutoff(today_value)
    if cutoff is None:
        return None
    cutoff_year, cutoff_month = cutoff
    if cutoff_year != work_year:
        return None

    latest = latest_generated_course_month(work_year)
    start_year, start_month = (work_year, 1) if latest is None else increment_year_month(*latest)
    if (start_year, start_month) > (cutoff_year, cutoff_month):
        return None

    missing_months: list[tuple[int, int]] = []
    current_year, current_month = start_year, start_month
    while (current_year, current_month) <= (cutoff_year, cutoff_month):
        missing_months.append((current_year, current_month))
        current_year, current_month = increment_year_month(current_year, current_month)

    if not missing_months:
        return None

    first_missing_year, first_missing_month = missing_months[0]
    cutoff_last_day = f"{cutoff_year:04d}-{cutoff_month:02d}-{monthrange(cutoff_year, cutoff_month)[1]:02d}"
    first_missing_day = f"{first_missing_year:04d}-{first_missing_month:02d}-01"
    active_count = int(
        fetch_scalar(
            """
            SELECT COUNT(*)
            FROM iscrizioni_corsi
            WHERE stato_iscrizione = 'Attiva'
              AND COALESCE(NULLIF(data_inizio, ''), data_iscrizione, ?) <= ?
              AND (data_fine IS NULL OR data_fine = '' OR data_fine >= ?)
            """,
            (cutoff_last_day, cutoff_last_day, first_missing_day),
        )
        or 0
    )
    if active_count <= 0:
        return None

    missing_labels = [f"{month_label(month)} {year}" for year, month in missing_months]
    latest_label = (
        f"{month_label(latest[1])} {latest[0]}"
        if latest is not None
        else "nessuna generazione ancora presente"
    )
    return {
        "work_year": work_year,
        "latest_label": latest_label,
        "missing_months": missing_months,
        "missing_labels": missing_labels,
        "missing_labels_text": ", ".join(missing_labels),
        "message": (
            f"L'ultima generazione massiva registrata e {latest_label}. "
            f"Risultano ancora da generare: {', '.join(missing_labels)}. "
            "Vuoi eseguire ora la generazione automatica dei mesi mancanti?"
        ),
    }


def generate_course_rates_for_month(
    connection: sqlite3.Connection,
    anno: int,
    mese: int,
    *,
    data_scadenza: str = "",
    note: str = "",
) -> tuple[int, int]:
    if mese < 1 or mese > 12:
        raise ValueError("Il mese deve essere compreso tra 1 e 12.")

    first_day = f"{anno}-{mese:02d}-01"
    last_day = f"{anno}-{mese:02d}-{monthrange(anno, mese)[1]:02d}"
    inserted = 0
    skipped = 0
    iscrizioni_attive = connection.execute(
        """
        SELECT id, quota_mensile
        FROM iscrizioni_corsi
        WHERE stato_iscrizione = 'Attiva'
          AND COALESCE(NULLIF(data_inizio, ''), data_iscrizione, ?) <= ?
          AND (data_fine IS NULL OR data_fine = '' OR data_fine >= ?)
        ORDER BY id
        """,
        (last_day, last_day, first_day),
    ).fetchall()

    for iscrizione in iscrizioni_attive:
        cursor = connection.execute(
            """
            INSERT OR IGNORE INTO rate_corsi_mensili (
                iscrizione_corso_id, anno, mese, importo_dovuto, data_scadenza, note
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                iscrizione["id"],
                anno,
                mese,
                iscrizione["quota_mensile"],
                data_scadenza or None,
                note or None,
            ),
        )
        if cursor.rowcount == 1:
            inserted += 1
        else:
            skipped += 1
    return inserted, skipped


def ensure_course_rate_for_enrollment(
    connection: sqlite3.Connection,
    iscrizione_corso_id: int,
    anno: int,
    mese: int,
    *,
    note: str = "",
) -> int:
    if mese < 1 or mese > 12:
        raise ValueError("Il mese deve essere compreso tra 1 e 12.")

    first_day = f"{anno}-{mese:02d}-01"
    last_day = f"{anno}-{mese:02d}-{monthrange(anno, mese)[1]:02d}"
    iscrizione = connection.execute(
        """
        SELECT id, quota_mensile, COALESCE(NULLIF(data_inizio, ''), data_iscrizione, ?) AS data_attivazione, data_fine
        FROM iscrizioni_corsi
        WHERE id = ?
        """,
        (last_day, iscrizione_corso_id),
    ).fetchone()
    if iscrizione is None:
        raise ValueError("Iscrizione corso non trovata.")

    if iscrizione["data_attivazione"] and str(iscrizione["data_attivazione"]) > last_day:
        raise ValueError("La quota del mese di iscrizione non puo essere generata prima della data di attivazione.")
    if iscrizione["data_fine"] and str(iscrizione["data_fine"]) < first_day:
        raise ValueError("L'iscrizione corso risulta gia conclusa per il mese selezionato.")

    data_scadenza = last_day
    connection.execute(
        """
        INSERT OR IGNORE INTO rate_corsi_mensili (
            iscrizione_corso_id, anno, mese, importo_dovuto, data_scadenza, note
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            iscrizione_corso_id,
            anno,
            mese,
            iscrizione["quota_mensile"],
            data_scadenza,
            note or None,
        ),
    )
    row = connection.execute(
        """
        SELECT id
        FROM rate_corsi_mensili
        WHERE iscrizione_corso_id = ? AND anno = ? AND mese = ?
        """,
        (iscrizione_corso_id, anno, mese),
    ).fetchone()
    if row is None:
        raise ValueError("Impossibile generare la quota mensile del mese di iscrizione.")
    return int(row["id"])


def ensure_course_rates_for_enrollment_range(
    connection: sqlite3.Connection,
    iscrizione_corso_id: int,
    start_year: int,
    start_month: int,
    end_year: int,
    end_month: int,
    *,
    note: str = "",
) -> list[int]:
    if start_month < 1 or start_month > 12 or end_month < 1 or end_month > 12:
        raise ValueError("Il mese deve essere compreso tra 1 e 12.")
    if (end_year, end_month) < (start_year, start_month):
        raise ValueError("Il mese finale deve essere uguale o successivo al mese iniziale.")

    rate_ids: list[int] = []
    current_year = start_year
    current_month = start_month
    while (current_year, current_month) <= (end_year, end_month):
        rate_ids.append(
            ensure_course_rate_for_enrollment(
                connection,
                iscrizione_corso_id,
                current_year,
                current_month,
                note=note,
            )
        )
        if current_month == 12:
            current_year += 1
            current_month = 1
        else:
            current_month += 1
    return rate_ids


def auto_generate_course_rates_on_startup(today_value: date | None = None) -> dict[str, object]:
    today_value = today_value or date.today()
    work_year = int(today_value.year)
    generated_months: list[str] = []
    inserted_total = 0
    skipped_total = 0
    with get_connection() as connection:
        for mese in range(1, today_value.month + 1):
            inserted, skipped = generate_course_rates_for_month(
                connection,
                work_year,
                mese,
                data_scadenza=f"{work_year:04d}-{mese:02d}-{monthrange(work_year, mese)[1]:02d}",
                note="Generazione automatica all'avvio del software",
            )
            inserted_total += inserted
            skipped_total += skipped
            if inserted:
                generated_months.append(f"{month_label(mese)} {work_year}")
        connection.commit()
    return {
        "work_year": work_year,
        "generated_months": generated_months,
        "inserted_total": inserted_total,
        "skipped_total": skipped_total,
    }


def export_activity_log_xls() -> None:
    ensure_output_dir()
    with get_connection() as connection:
        if not table_exists(connection, "registro_attivita"):
            return
        prune_old_activity_log_rows(connection)
        connection.commit()
        rows = connection.execute(
            """
            SELECT
                id,
                data_ora,
                substr(data_ora, 1, 10) AS data,
                substr(data_ora, 12, 8) AS ora,
                COALESCE(username, '') AS username,
                COALESCE(associato_codice, '') AS associato_codice,
                COALESCE(associato_nominativo, '') AS associato_nominativo,
                COALESCE(nome_pc, '') AS nome_pc,
                COALESCE(ip_client, '') AS ip_client,
                COALESCE(metodo_http, '') AS metodo_http,
                COALESCE(percorso, '') AS percorso,
                COALESCE(categoria, '') AS categoria,
                COALESCE(descrizione_attivita, '') AS descrizione_attivita,
                COALESCE(dettaglio, '') AS dettaglio,
                COALESCE(esito, '') AS esito,
                COALESCE(anno_lavoro, '') AS anno_lavoro,
                COALESCE(user_agent, '') AS user_agent
            FROM registro_attivita
            ORDER BY data_ora DESC, id DESC
            """
        ).fetchall()

    headers = [
        "ID",
        "Data",
        "Ora",
        "DataOra",
        "Utente",
        "Codice associato",
        "Associato",
        "Nome PC",
        "IP client",
        "Metodo",
        "Percorso",
        "Categoria",
        "Attivita",
        "Dettaglio",
        "Esito",
        "Anno lavoro",
        "User agent",
    ]
    workbook_rows = [headers]
    for row in rows:
        workbook_rows.append(
            [
                str(row["id"]),
                row["data"],
                row["ora"],
                row["data_ora"],
                row["username"],
                row["associato_codice"],
                row["associato_nominativo"],
                row["nome_pc"],
                row["ip_client"],
                row["metodo_http"],
                row["percorso"],
                row["categoria"],
                row["descrizione_attivita"],
                row["dettaglio"],
                row["esito"],
                str(row["anno_lavoro"]),
                row["user_agent"],
            ]
        )

    def xml_cell(value: object) -> str:
        safe_value = html.escape("" if value is None else str(value))
        return f'<Cell><Data ss:Type="String">{safe_value}</Data></Cell>'

    xml_rows = [
        "<Row>" + "".join(xml_cell(cell) for cell in row_values) + "</Row>"
        for row_values in workbook_rows
    ]
    workbook = (
        '<?xml version="1.0"?>\n'
        '<?mso-application progid="Excel.Sheet"?>\n'
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n'
        ' xmlns:o="urn:schemas-microsoft-com:office:office"\n'
        ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n'
        ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n'
        ' <Worksheet ss:Name="Registro attivita">\n'
        "  <Table>\n"
        f"   {''.join(xml_rows)}\n"
        "  </Table>\n"
        " </Worksheet>\n"
        "</Workbook>\n"
    )
    ACTIVITY_LOG_XLS_PATH.write_text(workbook, encoding="utf-8")


def record_activity(
    *,
    username: str,
    path: str,
    method: str,
    description: str,
    category: str,
    outcome: str,
    work_year: int | None = None,
    detail: str = "",
    associato_id: int | None = None,
    associato_codice: str = "",
    associato_nominativo: str = "",
    ip_client: str = "",
    user_agent: str = "",
) -> None:
    try:
        with get_connection() as connection:
            if not table_exists(connection, "registro_attivita"):
                return
            prune_old_activity_log_rows(connection)
            connection.execute(
                """
                INSERT INTO registro_attivita (
                    data_ora,
                    username,
                    associato_id,
                    associato_codice,
                    associato_nominativo,
                    nome_pc,
                    ip_client,
                    metodo_http,
                    percorso,
                    categoria,
                    descrizione_attivita,
                    dettaglio,
                    esito,
                    anno_lavoro,
                    user_agent
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    current_timestamp(),
                    username,
                    associato_id,
                    associato_codice,
                    associato_nominativo,
                    os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or "",
                    ip_client,
                    method,
                    path,
                    category,
                    description,
                    detail,
                    outcome,
                    work_year,
                    user_agent,
                ),
            )
            connection.commit()
        export_activity_log_xls()
    except Exception:
        pass


def iscrizione_corso_year_relevance_sql(alias: str = "ic") -> str:
    return f"""
        (
            substr(COALESCE(NULLIF({alias}.data_inizio, ''), {alias}.data_iscrizione, ''), 1, 4) = ?
            OR EXISTS (
                SELECT 1
                FROM rate_corsi_mensili r
                WHERE r.iscrizione_corso_id = {alias}.id
                  AND r.anno = ?
            )
        )
    """


def iscrizione_corso_year_relevance_params(work_year: int) -> tuple[object, ...]:
    return (str(work_year), work_year)


def associato_year_relevance_sql(alias: str = "a") -> str:
    return f"""
        (
            substr(COALESCE({alias}.data_prima_iscrizione, ''), 1, 4) = ?
            OR EXISTS (
                SELECT 1
                FROM tesseramenti_annuali t
                WHERE t.associato_id = {alias}.id
                  AND t.anno_sociale = ?
            )
            OR EXISTS (
                SELECT 1
                FROM iscrizioni_corsi ic
                WHERE ic.associato_id = {alias}.id
                  AND {iscrizione_corso_year_relevance_sql('ic')}
            )
            OR EXISTS (
                SELECT 1
                FROM iscrizioni_campi_estivi ice
                JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
                WHERE ice.associato_id = {alias}.id
                  AND ce.anno = ?
            )
            OR EXISTS (
                SELECT 1
                FROM iscrizioni_eventi ie
                JOIN eventi e ON e.id = ie.evento_id
                WHERE ie.associato_id = {alias}.id
                  AND substr(COALESCE(e.data_evento, ''), 1, 4) = ?
            )
        )
    """


def associato_year_relevance_params(work_year: int) -> tuple[object, ...]:
    return (
        str(work_year),
        work_year,
        *iscrizione_corso_year_relevance_params(work_year),
        work_year,
        str(work_year),
    )


def corso_year_relevance_sql(alias: str = "c") -> str:
    return f"""
        (
            substr(COALESCE({alias}.creato_il, ''), 1, 4) = ?
            OR EXISTS (
                SELECT 1
                FROM iscrizioni_corsi ic
                WHERE ic.corso_id = {alias}.id
                  AND {iscrizione_corso_year_relevance_sql('ic')}
            )
        )
    """


def corso_year_relevance_params(work_year: int) -> tuple[object, ...]:
    return (str(work_year), *iscrizione_corso_year_relevance_params(work_year))


def available_work_years(selected_year: int) -> list[int]:
    rows = fetch_all(
        """
        WITH anni AS (
            SELECT anno_sociale AS anno FROM tesseramenti_annuali
            UNION
            SELECT anno FROM rate_corsi_mensili
            UNION
            SELECT anno FROM campi_estivi
            UNION
            SELECT CAST(strftime('%Y', data_evento) AS INTEGER) FROM eventi
            UNION
            SELECT CAST(strftime('%Y', data_prima_iscrizione) AS INTEGER) FROM associati
        )
        SELECT DISTINCT anno
        FROM anni
        WHERE anno IS NOT NULL
        ORDER BY anno DESC
        """
    )
    values = {selected_year, date.today().year}
    values.update(int(row["anno"]) for row in rows if row["anno"])
    return sorted(values, reverse=True)


def month_label(month_number: int) -> str:
    return MONTH_NAMES.get(month_number, str(month_number))


def clean_phone_number(value: str | None) -> str:
    digits = re.sub(r"\D", "", value or "")
    if digits.startswith("00"):
        digits = digits[2:]
    if not digits:
        return ""
    if digits.startswith("39"):
        return digits
    if digits.startswith("0"):
        return f"39{digits.lstrip('0')}"
    if digits.startswith("3"):
        return f"39{digits}"
    return digits


def plain_text(value: object) -> str:
    return re.sub(r"<[^>]+>", "", "" if value is None else str(value)).strip()


def normalize_codice_fiscale(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (value or "").upper())[:16]


def normalize_lookup_key(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", (value or "").strip())
    ascii_only = "".join(character for character in normalized if not unicodedata.combining(character))
    return re.sub(r"[^A-Z0-9]", "", ascii_only.upper())


def calculate_age(birth_date_value: str | None, *, today: date | None = None) -> int | None:
    if not birth_date_value:
        return None
    try:
        born = date.fromisoformat(birth_date_value)
    except ValueError:
        return None
    reference = today or date.today()
    years = reference.year - born.year
    if (reference.month, reference.day) < (born.month, born.day):
        years -= 1
    return max(years, 0)


def label_with_age(label: str, birth_date_value: str | None) -> str:
    age = calculate_age(birth_date_value)
    if age is None:
        return label
    return f"{label} ({age} anni)"


@lru_cache(maxsize=1)
def comuni_records() -> tuple[dict[str, object], ...]:
    if not COMUNI_JSON_PATH.exists():
        return ()
    try:
        data = json.loads(COMUNI_JSON_PATH.read_text(encoding="utf-8"))
    except Exception:
        return ()

    records: list[dict[str, object]] = []
    for item in data:
        nome = str(item.get("nome") or "").strip()
        sigla = str(item.get("sigla") or "").strip()
        codice_catastale = str(item.get("codiceCatastale") or "").upper().strip()
        caps = [str(cap).strip() for cap in (item.get("cap") or []) if str(cap).strip()]
        if not nome:
            continue
        records.append(
            {
                "nome": nome,
                "sigla": sigla,
                "codice_catastale": codice_catastale,
                "caps": caps,
                "key": normalize_lookup_key(nome),
            }
        )
    return tuple(records)


@lru_cache(maxsize=1)
def comuni_lookup_by_catastale() -> dict[str, tuple[str, str]]:
    lookup: dict[str, tuple[str, str]] = {}
    for item in comuni_records():
        code = str(item["codice_catastale"])
        if not code:
            continue
        comune = str(item["nome"])
        provincia = str(item["sigla"])
        lookup[code] = (comune, provincia)
    return lookup


@lru_cache(maxsize=1)
def comuni_lookup_by_name() -> dict[str, dict[str, object]]:
    lookup: dict[str, dict[str, object]] = {}
    for item in comuni_records():
        key = str(item["key"])
        if key and key not in lookup:
            lookup[key] = dict(item)
    return lookup


@lru_cache(maxsize=1)
def comuni_lookup_by_cap() -> dict[str, list[dict[str, object]]]:
    lookup: dict[str, list[dict[str, object]]] = {}
    for item in comuni_records():
        for cap in item["caps"]:
            lookup.setdefault(cap, []).append(dict(item))
    for cap, items in lookup.items():
        items.sort(key=lambda row: (str(row["nome"]), str(row["sigla"])))
    return lookup


def lookup_comune_details(value: str) -> dict[str, str] | None:
    key = normalize_lookup_key(value)
    if not key:
        return None
    item = comuni_lookup_by_name().get(key)
    if item is None:
        partial_matches = [record for record in comuni_records() if str(record["key"]).startswith(key)]
        if len(partial_matches) == 1:
            item = dict(partial_matches[0])
    if item is None:
        return None
    caps = [str(cap).strip() for cap in (item.get("caps") or []) if str(cap).strip()]
    return {
        "comune": str(item["nome"]),
        "provincia": str(item["sigla"]),
        "codice_catastale": str(item["codice_catastale"]),
        "cap": caps[0] if caps else "",
    }


def lookup_cap_details(value: str) -> dict[str, object] | None:
    cap = re.sub(r"\D", "", value or "")[:5]
    if len(cap) != 5:
        return None
    matches = comuni_lookup_by_cap().get(cap, [])
    if not matches:
        return None
    first = matches[0]
    province_values = sorted({str(item["sigla"]) for item in matches if str(item["sigla"])})
    city_values = sorted({str(item["nome"]) for item in matches if str(item["nome"])})
    return {
        "cap": cap,
        "citta": str(first["nome"]),
        "provincia": str(first["sigla"]),
        "ambiguous": len(city_values) > 1,
        "candidates": city_values[:10],
        "province_candidates": province_values[:10],
    }


def resolve_birth_year(two_digits: int, month: int, day: int) -> int | None:
    reference = date.today()
    candidates = [1900 + two_digits, 2000 + two_digits]
    valid_candidates: list[tuple[int, int]] = []
    for year_value in candidates:
        try:
            candidate_date = date(year_value, month, day)
        except ValueError:
            continue
        if candidate_date > reference:
            continue
        age = calculate_age(candidate_date.isoformat())
        if age is None:
            continue
        if 0 <= age <= 115:
            valid_candidates.append((age, year_value))

    if valid_candidates:
        valid_candidates.sort(key=lambda item: (item[0], item[1]))
        return valid_candidates[0][1]

    for year_value in candidates:
        try:
            date(year_value, month, day)
            return year_value
        except ValueError:
            continue
    return None


def decode_codice_fiscale_birth_data(value: str) -> dict[str, str] | None:
    codice_fiscale = normalize_codice_fiscale(value)
    if len(codice_fiscale) != 16 or not re.fullmatch(r"[A-Z0-9]{16}", codice_fiscale):
        return None

    month = CODICE_FISCALE_MONTHS.get(codice_fiscale[8])
    if month is None:
        return None

    try:
        two_digits_year = int(codice_fiscale[6:8])
        encoded_day = int(codice_fiscale[9:11])
    except ValueError:
        return None

    is_female = encoded_day > 40
    day = encoded_day - 40 if is_female else encoded_day
    year_value = resolve_birth_year(two_digits_year, month, day)
    if year_value is None:
        return None

    try:
        birth_date_value = date(year_value, month, day).isoformat()
    except ValueError:
        return None

    comune_nascita = ""
    provincia_nascita = ""
    birth_place = comuni_lookup_by_catastale().get(codice_fiscale[11:15])
    if birth_place:
        comune_nascita, provincia_nascita = birth_place

    return {
        "codice_fiscale": codice_fiscale,
        "data_nascita": birth_date_value,
        "comune_nascita": comune_nascita,
        "provincia_nascita": provincia_nascita,
        "sesso": "F" if is_female else "M",
    }


def codice_fiscale_token_from_text(value: str, *, is_name: bool = False) -> str:
    cleaned = re.sub(r"[^A-Z]", "", normalize_lookup_key(value))
    consonants = [character for character in cleaned if character not in "AEIOU"]
    vowels = [character for character in cleaned if character in "AEIOU"]
    if is_name and len(consonants) >= 4:
        code = consonants[0] + consonants[2] + consonants[3]
    else:
        code = "".join(consonants[:3])
    if len(code) < 3:
        code += "".join(vowels[: 3 - len(code)])
    return (code + "XXX")[:3]


def codice_fiscale_control_char(partial_code: str) -> str:
    total = 0
    for index, character in enumerate(partial_code.upper(), start=1):
        if index % 2 == 1:
            total += CODICE_FISCALE_ODD_MAP[character]
        else:
            total += CODICE_FISCALE_EVEN_MAP[character]
    return CODICE_FISCALE_CONTROL_CHARS[total % 26]


def calculate_codice_fiscale(
    *,
    nome: str,
    cognome: str,
    data_nascita: str,
    sesso: str,
    comune_nascita: str,
) -> dict[str, str] | None:
    if not all([nome.strip(), cognome.strip(), data_nascita.strip(), sesso.strip(), comune_nascita.strip()]):
        return None
    try:
        born = date.fromisoformat(data_nascita)
    except ValueError:
        return None

    comune = lookup_comune_details(comune_nascita)
    if comune is None or not comune["codice_catastale"]:
        return None

    gender = (sesso or "").strip().upper()
    if gender not in {"M", "F"}:
        return None

    surname_code = codice_fiscale_token_from_text(cognome)
    name_code = codice_fiscale_token_from_text(nome, is_name=True)
    year_code = f"{born.year % 100:02d}"
    month_code = CODICE_FISCALE_MONTH_CODES[born.month]
    day_number = born.day + (40 if gender == "F" else 0)
    day_code = f"{day_number:02d}"
    partial_code = f"{surname_code}{name_code}{year_code}{month_code}{day_code}{comune['codice_catastale']}"
    control = codice_fiscale_control_char(partial_code)
    return {
        "codice_fiscale": partial_code + control,
        "comune_nascita": comune["comune"],
        "provincia_nascita": comune["provincia"],
    }


def json_response(start_response, payload: dict, *, status: str = "200 OK") -> list[bytes]:
    content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    start_response(
        status,
        [
            ("Content-Type", "application/json; charset=utf-8"),
            ("Content-Length", str(len(content))),
        ],
    )
    return [content]


def normalize_username(value: str) -> str:
    return re.sub(r"\s+", "", (value or "").strip()).lower()


def validate_username(value: str) -> str:
    username = normalize_username(value)
    if not re.fullmatch(r"[a-z0-9._-]{3,32}", username):
        raise ValueError("Lo username deve contenere da 3 a 32 caratteri: lettere, numeri, punto, trattino o underscore.")
    return username


def validate_password(password: str, confirmation: str) -> str:
    if len(password) < 8:
        raise ValueError("La password deve contenere almeno 8 caratteri.")
    if password != confirmation:
        raise ValueError("La conferma password non corrisponde.")
    return password


def hash_password(password: str, *, salt: bytes | None = None, iterations: int = PASSWORD_HASH_ITERATIONS) -> tuple[str, str, int]:
    current_salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), current_salt, iterations)
    return current_salt.hex(), digest.hex(), iterations


def verify_password(password: str, row: sqlite3.Row) -> bool:
    salt = bytes.fromhex(row["password_salt"])
    iterations = int(row["password_iterations"] or PASSWORD_HASH_ITERATIONS)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations).hex()
    return hmac.compare_digest(digest, row["password_hash"])


def current_timestamp() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def years_ago_timestamp(years: int) -> str:
    reference = datetime.now().replace(microsecond=0)
    try:
        target = reference.replace(year=reference.year - years)
    except ValueError:
        # Gestisce il 29 febbraio negli anni non bisestili.
        target = reference.replace(year=reference.year - years, day=28)
    return target.isoformat(sep=" ")


def prune_old_activity_log_rows(connection: sqlite3.Connection) -> int:
    if not table_exists(connection, "registro_attivita"):
        return 0
    cursor = connection.execute(
        """
        DELETE FROM registro_attivita
        WHERE data_ora < ?
        """,
        (years_ago_timestamp(ACTIVITY_LOG_RETENTION_YEARS),),
    )
    return max(int(cursor.rowcount or 0), 0)


def build_cookie_header(name: str, value: str, *, max_age: int | None = None) -> tuple[str, str]:
    cookie = SimpleCookie()
    cookie[name] = value
    morsel = cookie[name]
    morsel["path"] = "/"
    morsel["httponly"] = True
    morsel["samesite"] = "Lax"
    if max_age is not None:
        morsel["max-age"] = str(max_age)
    return "Set-Cookie", morsel.OutputString()


def session_cookie_header(token: str) -> tuple[str, str]:
    return build_cookie_header(
        SESSION_COOKIE_NAME,
        token,
        max_age=SESSION_DURATION_DAYS * 24 * 60 * 60,
    )


def clear_session_cookie_header() -> tuple[str, str]:
    return build_cookie_header(SESSION_COOKIE_NAME, "", max_age=0)


def get_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def fetch_all(query: str, params: tuple | list = ()) -> list[sqlite3.Row]:
    with get_connection() as connection:
        return connection.execute(query, params).fetchall()


def fetch_one(query: str, params: tuple | list = ()) -> sqlite3.Row | None:
    with get_connection() as connection:
        return connection.execute(query, params).fetchone()


def fetch_scalar(query: str, params: tuple | list = ()) -> object:
    row = fetch_one(query, params)
    if row is None:
        return None
    return row[0]


def execute(query: str, params: tuple | list = ()) -> None:
    with get_connection() as connection:
        connection.execute(query, params)
        connection.commit()


def parse_request(environ: dict) -> tuple[str, str, dict[str, str], dict[str, str], dict[str, str]]:
    path = environ.get("PATH_INFO") or "/"
    method = (environ.get("REQUEST_METHOD") or "GET").upper()
    query_params = {
        key: ",".join(values)
        for key, values in parse_qs(
            environ.get("QUERY_STRING", ""),
            keep_blank_values=True,
        ).items()
    }

    form_data: dict[str, str] = {}
    if method == "POST":
        size = int(environ.get("CONTENT_LENGTH") or 0)
        payload = environ["wsgi.input"].read(size).decode("utf-8")
        form_data = {
            key: ",".join(values)
            for key, values in parse_qs(payload, keep_blank_values=True).items()
        }

    cookie_jar = SimpleCookie()
    cookie_jar.load(environ.get("HTTP_COOKIE", ""))
    cookies = {key: morsel.value for key, morsel in cookie_jar.items()}

    return path, method, query_params, form_data, cookies


def esc(value: object) -> str:
    return html.escape("" if value is None else str(value))


def normalized(form_data: dict[str, str], key: str, default: str = "") -> str:
    return (form_data.get(key, default) or "").strip()


def multi_values(form_data: dict[str, str], key: str) -> list[str]:
    raw_value = form_data.get(key, "") or ""
    if not raw_value:
        return []
    return [item.strip() for item in raw_value.split(",") if item.strip()]


def required(form_data: dict[str, str], key: str, label: str) -> str:
    value = normalized(form_data, key)
    if not value:
        raise ValueError(f"Il campo '{label}' e obbligatorio.")
    return value


def optional(form_data: dict[str, str], key: str) -> str | None:
    value = normalized(form_data, key)
    return value or None


def money(value: object) -> str:
    if value in (None, ""):
        return ""
    amount = float(value)
    rendered = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{rendered} EUR"


def decimal_amount(value: object, *, minimum: str = "0.00") -> Decimal:
    amount = Decimal(str(value or "0")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    floor = Decimal(minimum).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if amount < floor:
        raise ValueError("L'importo indicato non e valido.")
    return amount


def decimal_input(value: object) -> str:
    return format(decimal_amount(value), ".2f")


def decimal_or_none(value: object) -> Decimal | None:
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value).strip()).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None


def summary_label_index(columns: list[tuple]) -> int:
    for index, column in enumerate(columns):
        label = (column[1] or "").lower()
        if label not in {"ricevuta", "azioni"}:
            return index
    return 0


def summary_total_indexes(columns: list[tuple]) -> list[int]:
    total_hints = ("importo", "dovuto", "pagato", "residuo", "quota", "saldo")
    indexes: list[int] = []
    for index, column in enumerate(columns):
        label = str(column[1] or "").lower()
        if any(hint in label for hint in total_hints):
            indexes.append(index)
    return indexes


def summary_rows_for_table(rows: list[sqlite3.Row], columns: list[tuple]) -> list[list[str]]:
    if not rows:
        return []

    total_indexes = summary_total_indexes(columns)
    totals: dict[int, str] = {}
    for index in total_indexes:
        column = columns[index]
        total = Decimal("0.00")
        found = False
        for row in rows:
            amount = decimal_or_none(row[column[0]])
            if amount is None:
                continue
            total += amount
            found = True
        if found:
            totals[index] = money(total)

    if not totals:
        return []

    summary_row = [""] * len(columns)
    summary_row[summary_label_index(columns)] = "Totali"
    for index, value in totals.items():
        summary_row[index] = value
    return [summary_row]


def option_data_attrs(row: sqlite3.Row, data_keys: list[str] | None = None) -> str:
    if not data_keys:
        return ""
    row_keys = set(row.keys())
    attrs: list[str] = []
    for key in data_keys:
        if key not in row_keys:
            continue
        value = row[key]
        if value in (None, ""):
            continue
        html_key = key.replace("_", "-")
        attrs.append(f' data-{esc(html_key)}="{esc(value)}"')
    return "".join(attrs)


def render_associato_options(rows: list[sqlite3.Row], selected: str | None = None) -> str:
    return render_select_options(rows, selected, data_keys=["search_text", "autocomplete_label"])


def render_select_options(
    rows: list[sqlite3.Row],
    selected: str | None = None,
    blank_label: str = "Seleziona...",
    data_keys: list[str] | None = None,
) -> str:
    items = [f'<option value="">{esc(blank_label)}</option>']
    selected_value = "" if selected is None else str(selected)
    for row in rows:
        value = str(row["id"])
        label = row["label"]
        selected_attr = ' selected="selected"' if value == selected_value else ""
        items.append(
            f'<option value="{esc(value)}"{selected_attr}{option_data_attrs(row, data_keys)}>{esc(label)}</option>'
        )
    return "".join(items)


def render_select_options_multi(
    rows: list[sqlite3.Row],
    selected_values: list[str] | None = None,
    data_keys: list[str] | None = None,
) -> str:
    selected = set(selected_values or [])
    items = []
    for row in rows:
        value = str(row["id"])
        label = row["label"]
        selected_attr = ' selected="selected"' if value in selected else ""
        items.append(
            f'<option value="{esc(value)}"{selected_attr}{option_data_attrs(row, data_keys)}>{esc(label)}</option>'
        )
    return "".join(items)


def render_static_options(
    options: list[tuple[str, str]],
    selected: str | None = None,
    *,
    blank_label: str | None = None,
) -> str:
    items: list[str] = []
    if blank_label is not None:
        items.append(f'<option value="">{esc(blank_label)}</option>')

    selected_value = "" if selected is None else str(selected)
    for value, label in options:
        selected_attr = ' selected="selected"' if str(value) == selected_value else ""
        items.append(f'<option value="{esc(value)}"{selected_attr}>{esc(label)}</option>')
    return "".join(items)


def hidden_fields_html(hidden_fields: dict[str, str] | None = None) -> str:
    if not hidden_fields:
        return ""
    return "".join(
        f'<input type="hidden" name="{esc(key)}" value="{esc(value)}">'
        for key, value in hidden_fields.items()
    )


def input_field(
    label: str,
    name: str,
    *,
    input_type: str = "text",
    value: str = "",
    required_field: bool = False,
    step: str | None = None,
    minimum: str | None = None,
    placeholder: str = "",
    wide: bool = False,
    element_id: str | None = None,
    attrs: dict[str, str] | None = None,
    revealable: bool = False,
) -> str:
    attr_parts = [
        f'type="{esc(input_type)}"',
        f'name="{esc(name)}"',
        f'value="{esc(value)}"',
        'class="control"',
    ]
    if element_id:
        attr_parts.append(f'id="{esc(element_id)}"')
    if required_field:
        attr_parts.append("required")
    if step is not None:
        attr_parts.append(f'step="{esc(step)}"')
    if minimum is not None:
        attr_parts.append(f'min="{esc(minimum)}"')
    if placeholder:
        attr_parts.append(f'placeholder="{esc(placeholder)}"')
    if attrs is not None:
        for key, attr_value in attrs.items():
            attr_parts.append(f'{esc(key)}="{esc(attr_value)}"')
    class_name = "field wide" if wide else "field"
    input_html = f"<input {' '.join(attr_parts)}>"
    if revealable and input_type == "password":
        input_html = (
            '<div class="password-field-wrap">'
            f"{input_html}"
            '<button type="button" class="password-toggle" onclick="togglePasswordVisibility(this)">Mostra</button>'
            "</div>"
        )
    return (
        f'<label class="{class_name}"><span>{esc(label)}</span>'
        f"{input_html}</label>"
    )


def readonly_field(label: str, value: str, *, wide: bool = False) -> str:
    class_name = "field wide" if wide else "field"
    return (
        f'<label class="{class_name}"><span>{esc(label)}</span>'
        f'<div class="control readonly-control">{esc(value)}</div></label>'
    )


def select_field(
    label: str,
    name: str,
    options_html: str,
    *,
    required_field: bool = False,
    wide: bool = False,
    element_id: str | None = None,
    attrs: dict[str, str] | None = None,
    control_class: str = "control",
    searchable: bool = False,
    search_placeholder: str = "Cerca per codice, nome, telefono o email...",
) -> str:
    class_name = "field wide" if wide else "field"
    required_attr = " required" if required_field and not searchable else ""
    search_required_attr = " required" if required_field and searchable else ""
    select_id = element_id or (f"{name}-{uuid.uuid4().hex[:8]}" if searchable else None)
    panel_id = f"{select_id}-autocomplete" if searchable and select_id else None
    extra_attrs = ""
    if select_id:
        extra_attrs += f' id="{esc(select_id)}"'
    if attrs:
        extra_attrs += "".join(f' {esc(key)}="{esc(value)}"' for key, value in attrs.items())
    select_class = control_class if not searchable else f"{control_class} searchable-select-source".strip()
    search_html = ""
    if searchable and select_id:
        search_html = (
            f'<div class="select-autocomplete">'
            f'<input type="search" class="control select-search-control" '
            f'placeholder="{esc(search_placeholder)}" data-select-search-target="{esc(select_id)}" '
            f'data-select-search-panel="{esc(panel_id or "")}" '
            f'oninput="handleSelectSearch(this)" onfocus="openSelectAutocomplete(this)" '
            f'onkeydown="handleSelectSearchKeydown(event, this)" onblur="closeSelectAutocompleteLater(this)" '
            f'autocomplete="off"{search_required_attr}>'
            f'<div id="{esc(panel_id or "")}" class="select-autocomplete-panel" hidden></div>'
            "</div>"
        )
    return (
        f'<label class="{class_name}"><span>{esc(label)}</span>'
        f"{search_html}"
        f'<select name="{esc(name)}" class="{esc(select_class)}"{required_attr}{extra_attrs}>{options_html}</select>'
        "</label>"
    )


def multi_select_field(
    label: str,
    name: str,
    options_html: str,
    *,
    required_field: bool = False,
    wide: bool = True,
    size: int = 8,
    element_id: str | None = None,
    attrs: dict[str, str] | None = None,
) -> str:
    class_name = "field wide" if wide else "field"
    required_attr = " required" if required_field else ""
    extra_attrs = ""
    if element_id:
        extra_attrs += f' id="{esc(element_id)}"'
    if attrs:
        extra_attrs += "".join(f' {esc(key)}="{esc(value)}"' for key, value in attrs.items())
    return (
        f'<label class="{class_name}"><span>{esc(label)}</span>'
        f'<select name="{esc(name)}" class="control multi-control" multiple size="{size}"{required_attr}{extra_attrs}>{options_html}</select>'
        "</label>"
    )


def textarea_field(
    label: str,
    name: str,
    *,
    value: str = "",
    rows: int = 4,
    wide: bool = True,
    ) -> str:
    class_name = "field wide" if wide else "field"
    return (
        f'<label class="{class_name}"><span>{esc(label)}</span>'
        f'<textarea name="{esc(name)}" rows="{rows}" class="control">{esc(value)}</textarea>'
        "</label>"
    )


def form_card(
    title: str,
    subtitle: str,
    action: str,
    fields_html: str,
    button_label: str,
    *,
    hidden_fields: dict[str, str] | None = None,
    card_class: str = "",
    form_attrs: dict[str, str] | None = None,
) -> str:
    section_class = f"card {card_class}".strip()
    form_extra_attrs = ""
    if form_attrs:
        form_extra_attrs = "".join(f' {esc(key)}="{esc(value)}"' for key, value in form_attrs.items())
    return f"""
    <section class="{section_class}">
      <div class="card-head">
        <h2>{esc(title)}</h2>
        <p>{esc(subtitle)}</p>
      </div>
      <form method="post" action="{esc(action)}" class="form-grid"{form_extra_attrs}>
        {hidden_fields_html(hidden_fields)}
        {fields_html}
        <div class="form-actions">
          <button type="submit" class="button">{esc(button_label)}</button>
        </div>
      </form>
    </section>
    """


def table_card(
    title: str,
    subtitle: str,
    rows: list[sqlite3.Row],
    columns: list[tuple],
    *,
    empty_message: str = "Nessun dato disponibile.",
    table_class: str = "",
    summary_rows: list[list[str]] | None = None,
    head_actions_html: str = "",
) -> str:
    subtitle_html = f"<p>{esc(subtitle)}</p>" if subtitle else ""
    actions_html = f'<div class="card-head-actions screen-only">{head_actions_html}</div>' if head_actions_html else ""
    return f"""
    <section class="card">
      <div class="card-head">
        <h2>{esc(title)}</h2>
        {subtitle_html}
        {actions_html}
      </div>
      {render_table(rows, columns, empty_message=empty_message, table_class=table_class, summary_rows=summary_rows)}
    </section>
    """


def render_table(
    rows: list[sqlite3.Row],
    columns: list[tuple],
    *,
    empty_message: str = "Nessun dato disponibile.",
    table_class: str = "",
    summary_rows: list[list[str]] | None = None,
) -> str:
    if not rows:
        return f'<div class="empty-state">{esc(empty_message)}</div>'

    total_indexes = summary_total_indexes(columns)
    total_index_set = set(total_indexes)
    head = []
    body = []
    for _, label, *_ in columns:
        head.append(f"<th>{esc(label)}</th>")

    for row in rows:
        cells = []
        for index, column in enumerate(columns):
            key = column[0]
            formatter = column[2] if len(column) > 2 else None
            value = row[key]
            rendered = formatter(value, row) if callable(formatter) else esc(value)
            cell_attrs = ""
            if index in total_index_set:
                amount = decimal_or_none(value)
                if amount is not None:
                    cell_attrs = f' data-sum-value="{format(amount, ".2f")}"'
            cells.append(f"<td{cell_attrs}>{rendered}</td>")
        body.append(f"<tr>{''.join(cells)}</tr>")

    footer_html = ""
    if summary_rows:
        footer_rows = []
        for summary_row in summary_rows:
            footer_cells = [f"<td>{esc(cell)}</td>" for cell in summary_row]
            footer_rows.append(f"<tr>{''.join(footer_cells)}</tr>")
        footer_html = f"<tfoot>{''.join(footer_rows)}</tfoot>"

    table_attrs = ""
    if summary_rows and total_indexes:
        table_attrs = (
            f' data-summary-columns="{esc(",".join(str(index) for index in total_indexes))}"'
            f' data-summary-label-index="{summary_label_index(columns)}"'
        )

    return (
        f'<div class="table-wrap"><table class="data-table {esc(table_class)}"{table_attrs}><thead><tr>'
        + "".join(head)
        + "</tr></thead><tbody>"
        + "".join(body)
        + f"</tbody>{footer_html}</table></div>"
    )


def stat_card(label: str, value: object, link: str) -> str:
    return f"""
    <a class="stat-card" href="{esc(link)}">
      <span class="stat-label">{esc(label)}</span>
      <strong class="stat-value">{esc(value)}</strong>
    </a>
    """


def message_banner(query_params: dict[str, str]) -> str:
    ok_message = query_params.get("ok")
    error_message = query_params.get("err")
    items = []
    if ok_message:
        items.append(f'<div class="notice success">{esc(ok_message)}</div>')
    if error_message:
        items.append(f'<div class="notice error">{esc(error_message)}</div>')
    return "".join(items)


def work_year_selector(query_params: dict[str, str], *, action: str = "/") -> str:
    selected_year = current_work_year(query_params)
    options = render_static_options(
        [(str(year), str(year)) for year in available_work_years(selected_year)],
        str(selected_year),
    )
    return f"""
    <section class="card compact screen-only work-year-card">
      <div class="card-head">
        <h2>Anno di lavoro</h2>
        <p>Seleziona l'anno operativo da usare in dashboard, report e filtri predefiniti.</p>
      </div>
      <form method="get" action="{esc(action)}" class="form-grid">
        {select_field("Anno di lavoro", "anno_lavoro", options, required_field=True)}
        <div class="form-actions">
          <button type="submit" class="button">Applica anno</button>
        </div>
      </form>
    </section>
    """


def header_work_year_selector(query_params: dict[str, str], *, action: str = "/") -> str:
    selected_year = current_work_year(query_params)
    options = render_static_options(
        [(str(year), str(year)) for year in available_work_years(selected_year)],
        str(selected_year),
        blank_label=None,
    )
    return f"""
    <form method="get" action="{esc(action)}" class="header-year-form screen-only">
      <label class="header-year-field">
        <span>Anno di lavoro</span>
        <select name="anno_lavoro" class="control">{options}</select>
      </label>
      <button type="submit" class="button">Applica</button>
    </form>
    """


def data_view_query(query_params: dict[str, str]) -> dict[str, str]:
    params = work_year_query(query_params)
    params["vista"] = "dati"
    return params


def insert_view_query(query_params: dict[str, str]) -> dict[str, str]:
    return work_year_query(query_params)


def current_page_query(query_params: dict[str, str]) -> dict[str, str]:
    if normalized(query_params, "vista", "") == "dati":
        return data_view_query(query_params)
    return work_year_query(query_params)


def view_mode_switch(current_path: str, query_params: dict[str, str], data_label: str) -> str:
    data_only = normalized(query_params, "vista", "") == "dati"
    primary_href = with_query(current_path, insert_view_query(query_params) if data_only else data_view_query(query_params))
    primary_label = "Torna a inserimento" if data_only else data_label
    secondary_href = with_query(current_path, data_view_query(query_params) if data_only else insert_view_query(query_params))
    secondary_label = data_label if data_only else "Vista inserimento"
    return f"""
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Vista area</span>
        <p>Passa dalla maschera di inserimento alla schermata dei dati gia registrati.</p>
      </div>
      <div class="report-toolbar-actions">
        <a class="button" href="{esc(primary_href)}">{esc(primary_label)}</a>
        <a class="button secondary" href="{esc(secondary_href)}">{esc(secondary_label)}</a>
      </div>
    </section>
    """


def data_view_search_toolbar() -> str:
    return """
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Ricerca dati</span>
        <p>Filtra rapidamente tutte le colonne presenti nelle tabelle dei dati gia registrati.</p>
      </div>
      <div class="report-toolbar-actions">
        <label class="report-search">
          <span>Cerca</span>
          <input
            type="search"
            class="control"
            placeholder="Filtra tutte le colonne..."
            oninput="handleDataSearch(this)"
          >
        </label>
      </div>
    </section>
    """


def dashboard_associati_search_toolbar() -> str:
    return """
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Posizione associati</span>
        <p>Cerca rapidamente un associato nella tabella della dashboard.</p>
      </div>
      <div class="report-toolbar-actions">
        <label class="report-search">
          <span>Cerca</span>
          <input
            type="search"
            class="control"
            placeholder="Filtra per nome, codice o stato..."
            oninput="handleDataSearch(this)"
          >
        </label>
      </div>
    </section>
    """


def posizione_associati_params(work_year: int) -> tuple[object, ...]:
    return (
        work_year,
        work_year,
        work_year,
        str(work_year),
        *associato_year_relevance_params(work_year),
    )


def posizione_associati_query(*, limit: int | None = None) -> str:
    limit_sql = f"\n        LIMIT {int(limit)}" if limit else ""
    return f"""
        WITH movimenti AS (
            SELECT
                associato_id,
                importo_dovuto,
                importo_pagato,
                saldo_residuo
            FROM v_tesseramenti_saldo
            WHERE anno_sociale = ?

            UNION ALL

            SELECT
                associato_id,
                importo_dovuto,
                importo_pagato,
                saldo_residuo
            FROM v_rate_corsi_saldo
            WHERE anno = ?

            UNION ALL

            SELECT
                associato_id,
                importo_dovuto,
                importo_pagato,
                saldo_residuo
            FROM v_campi_estivi_saldo
            WHERE anno = ?

            UNION ALL

            SELECT
                associato_id,
                importo_dovuto,
                importo_pagato,
                saldo_residuo
            FROM v_eventi_saldo
            WHERE substr(COALESCE(data_evento, ''), 1, 4) = ?
        ),
        posizioni AS (
            SELECT
                associato_id,
                COALESCE(SUM(importo_dovuto), 0) AS totale_dovuto,
                COALESCE(SUM(importo_pagato), 0) AS totale_pagato,
                COALESCE(SUM(saldo_residuo), 0) AS saldo_residuo
            FROM movimenti
            GROUP BY associato_id
        )
        SELECT
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            a.stato_associato,
            COALESCE(posizioni.totale_dovuto, 0) AS totale_dovuto,
            COALESCE(posizioni.totale_pagato, 0) AS totale_pagato,
            COALESCE(posizioni.saldo_residuo, 0) AS saldo_residuo
        FROM associati a
        LEFT JOIN posizioni ON posizioni.associato_id = a.id
        WHERE {associato_year_relevance_sql('a')}
          AND COALESCE(posizioni.totale_dovuto, 0) > 0
        ORDER BY associato{limit_sql}
    """


def delete_action_form(action: str, prompt: str, *, extra_fields: dict[str, str] | None = None) -> str:
    fields = dict(extra_fields or {})
    fields.setdefault("vista", "")
    return f"""
    <form
      method="post"
      action="{esc(action)}"
      class="inline-form"
      data-confirm-dialog="true"
      data-confirm-title="Conferma eliminazione"
      data-confirm-message="{esc(prompt)}"
      data-confirm-button="Elimina">
      {hidden_fields_html(fields)}
      <button type="submit" class="table-action danger">Elimina</button>
    </form>
    """


def action_links_html(
    *,
    edit_href: str | None = None,
    delete_action: str | None = None,
    delete_prompt: str | None = None,
    extra_links: list[tuple[str, str]] | None = None,
    extra_fields: dict[str, str] | None = None,
) -> str:
    items: list[str] = []
    if edit_href:
        items.append(f'<a class="table-action" href="{esc(edit_href)}">Modifica</a>')
    if extra_links:
        for href, label in extra_links:
            items.append(f'<a class="table-action" href="{esc(href)}">{esc(label)}</a>')
    if delete_action and delete_prompt:
        items.append(delete_action_form(delete_action, delete_prompt, extra_fields=extra_fields))
    return f'<div class="table-actions">{ "".join(items) }</div>'


def nav_link(current_path: str, href: str, label: str) -> str:
    active = " active" if current_path == href.split("?", 1)[0] else ""
    return f'<a class="nav-link{active}" href="{esc(href)}">{esc(label)}</a>'


def access_role_label(current_user: dict[str, object] | None) -> str:
    if not current_user:
        return ""
    return "Amministratore" if current_user.get("is_admin") else "Utente"


def render_navigation(current_path: str, query_params: dict[str, str], current_user: dict[str, object] | None = None) -> str:
    kicker_html = f'<span class="brand-kicker">{esc(APP_SUBTITLE)}</span>' if APP_SUBTITLE else ""
    copy_class = "brand-copy with-kicker" if APP_SUBTITLE else "brand-copy"
    admin_group = ""
    if current_user and current_user.get("is_admin"):
        admin_group = f"""
      <div class="nav-group">
        <span class="nav-group-title">Amministrazione</span>
        {nav_link(current_path, with_query("/maschere/utenti", work_year_query(query_params)), "Utenti")}
        {nav_link(current_path, with_query("/report/registro-attivita", work_year_query(query_params)), "Registro attivita")}
      </div>
      """
    user_block = ""
    if current_user:
        user_block = f"""
      <div class="sidebar-user-card">
        <div class="sidebar-user-inline">
          <strong class="sidebar-user-name">{esc(current_user.get("username", ""))}</strong>
        </div>
        <form method="post" action="/azioni/accesso/logout" class="sidebar-user-form">
          {hidden_fields_html(work_year_query(query_params))}
          <button type="submit" class="sidebar-user-button">Logout</button>
        </form>
      </div>
      """
    return f"""
    <nav class="sidebar">
      <div class="brand">
        <div class="brand-mark">
          <div class="brand-logo-wrap">
            <img class="brand-logo" src="{LOGO_URL}" alt="Logo {APP_NAME}">
          </div>
          <div class="{copy_class}">
            {kicker_html}
            <strong class="brand-title">{APP_NAME}</strong>
          </div>
        </div>
      </div>
      {user_block}
      <div class="sidebar-nav-scroll">
      <div class="nav-group">
        <span class="nav-group-title">Panoramica</span>
        {nav_link(current_path, with_query("/", work_year_query(query_params)), "Dashboard")}
      </div>
      <div class="nav-group">
        <span class="nav-group-title">STRUTTURA SOCIALE</span>
        {nav_link(current_path, with_query("/maschere/consiglio-direttivo", work_year_query(query_params)), "Consiglio Direttivo")}
        {nav_link(current_path, with_query("/maschere/associati", work_year_query(query_params)), "Associati")}
      </div>
      <div class="nav-group">
        <span class="nav-group-title">ISCRIZIONI</span>
        {nav_link(current_path, with_query("/maschere/tesseramenti", work_year_query(query_params)), "Tesseramenti")}
        {nav_link(current_path, with_query("/maschere/corsi", work_year_query(query_params)), "Corsi")}
        {nav_link(current_path, with_query("/maschere/campi-estivi", work_year_query(query_params)), ESTATE_LABEL)}
        {nav_link(current_path, with_query("/maschere/eventi", work_year_query(query_params)), "Eventi")}
      </div>
      <div class="nav-group">
        <span class="nav-group-title">GESTIONE ECONOMICA</span>
        {nav_link(current_path, with_query("/maschere/pagamenti-multi-area", work_year_query(query_params)), "Pagamenti")}
        {nav_link(current_path, with_query("/report/scadenze", work_year_query(query_params)), "Scadenze da incassare")}
        {nav_link(current_path, with_query("/report/incassi", work_year_query(query_params)), "Incassi totali")}
      </div>
      <div class="nav-group">
        <span class="nav-group-title">Report</span>
        {nav_link(current_path, with_query("/report/associati", work_year_query(query_params)), "Posizione associati")}
        {nav_link(current_path, with_query("/report/partecipanti", work_year_query(query_params)), "Partecipanti attivita")}
        {nav_link(current_path, with_query("/report/tesseramenti", work_year_query(query_params)), "Situazione tesseramenti")}
        {nav_link(current_path, with_query("/report/corsi", work_year_query(query_params)), "Situazione corsi")}
        {nav_link(current_path, with_query("/report/campi-estivi", work_year_query(query_params)), "Situazione campo estivo")}
        {nav_link(current_path, with_query("/report/eventi", work_year_query(query_params)), "Situazione eventi")}
      </div>
      <div class="nav-group">
        <span class="nav-group-title">Accesso</span>
        {nav_link(current_path, with_query("/maschere/accesso", work_year_query(query_params)), "Profilo accesso")}
      </div>
      {admin_group}
      </div>
    </nav>
    """


def app_dialog_markup() -> str:
    return """
    <section id="app-dialog-overlay" class="app-dialog-overlay" hidden>
      <div class="app-dialog-backdrop"></div>
      <div class="app-dialog-card" data-variant="info" role="alertdialog" aria-modal="true" aria-labelledby="app-dialog-title" aria-describedby="app-dialog-message">
        <div class="app-dialog-accent">
          <span id="app-dialog-badge" class="app-dialog-badge">Notifica</span>
          <h2 id="app-dialog-title">Notifica</h2>
          <p id="app-dialog-message" class="app-dialog-message"></p>
          <div id="app-dialog-extra" class="app-dialog-extra" hidden></div>
        </div>
        <div class="app-dialog-actions">
          <button type="button" id="app-dialog-cancel" class="button ghost">Annulla</button>
          <button type="button" id="app-dialog-confirm" class="button">Conferma</button>
        </div>
      </div>
    </section>
    """


def shared_dialog_script() -> str:
    return """
      let activeAppDialog = null;

      function togglePasswordVisibility(button) {
        const wrapper = button.closest('.password-field-wrap');
        if (!wrapper) {
          return;
        }
        const input = wrapper.querySelector('input');
        if (!input) {
          return;
        }
        const isHidden = input.type === 'password';
        input.type = isHidden ? 'text' : 'password';
        button.textContent = isHidden ? 'Nascondi' : 'Mostra';
      }

      function syncModalOpenState() {
        const reminderVisible = !!document.querySelector('[data-course-generation-reminder="true"]:not([hidden])');
        const dialogOverlay = document.getElementById('app-dialog-overlay');
        const dialogVisible = !!dialogOverlay && !dialogOverlay.hidden;
        document.body.classList.toggle('modal-open', reminderVisible || dialogVisible);
      }

      function appDialogElements() {
        const overlay = document.getElementById('app-dialog-overlay');
        if (!overlay) {
          return null;
        }
        return {
          overlay,
          backdrop: overlay.querySelector('.app-dialog-backdrop'),
          card: overlay.querySelector('.app-dialog-card'),
          badge: document.getElementById('app-dialog-badge'),
          title: document.getElementById('app-dialog-title'),
          message: document.getElementById('app-dialog-message'),
          extra: document.getElementById('app-dialog-extra'),
          cancelButton: document.getElementById('app-dialog-cancel'),
          confirmButton: document.getElementById('app-dialog-confirm'),
        };
      }

      function closeAppDialog(result = false) {
        const elements = appDialogElements();
        if (!elements || elements.overlay.hidden) {
          activeAppDialog = null;
          return;
        }
        const resolver = activeAppDialog && typeof activeAppDialog.resolve === 'function'
          ? activeAppDialog.resolve
          : null;
        elements.overlay.hidden = true;
        if (elements.card) {
          elements.card.dataset.variant = 'info';
        }
        if (elements.extra) {
          elements.extra.hidden = true;
          elements.extra.innerHTML = '';
        }
        if (elements.confirmButton) {
          elements.confirmButton.classList.remove('danger');
        }
        activeAppDialog = null;
        syncModalOpenState();
        if (resolver) {
          resolver(result);
        }
      }

      function showAppDialog(options = {}) {
        const elements = appDialogElements();
        if (!elements) {
          if (options.cancelLabel) {
            return Promise.resolve(window.confirm(options.message || 'Confermare l\\'operazione?'));
          }
          window.alert(options.message || 'Operazione non disponibile.');
          return Promise.resolve(true);
        }

        if (activeAppDialog && typeof activeAppDialog.resolve === 'function') {
          activeAppDialog.resolve(false);
        }

        const variant = options.variant || 'info';
        const dismissible = options.dismissible !== false;
        const confirmLabel = options.confirmLabel || 'Chiudi';
        const cancelLabel = options.cancelLabel || '';
        const badgeText = options.badge || (
          variant === 'danger' ? 'Attenzione'
            : variant === 'warning' ? 'Conferma'
            : 'Notifica'
        );

        elements.overlay.hidden = false;
        if (elements.card) {
          elements.card.dataset.variant = variant;
        }
        if (elements.badge) {
          elements.badge.textContent = badgeText;
        }
        if (elements.title) {
          elements.title.textContent = options.title || 'Notifica';
        }
        if (elements.message) {
          elements.message.textContent = options.message || '';
        }
        if (elements.extra) {
          if (options.html) {
            elements.extra.hidden = false;
            elements.extra.innerHTML = options.html;
          } else {
            elements.extra.hidden = true;
            elements.extra.innerHTML = '';
          }
        }
        if (elements.cancelButton) {
          if (cancelLabel) {
            elements.cancelButton.hidden = false;
            elements.cancelButton.textContent = cancelLabel;
          } else {
            elements.cancelButton.hidden = true;
          }
        }
        if (elements.confirmButton) {
          elements.confirmButton.textContent = confirmLabel;
          elements.confirmButton.classList.toggle('danger', variant === 'danger');
        }

        syncModalOpenState();

        return new Promise((resolve) => {
          activeAppDialog = {
            resolve,
            dismissible,
            resolveValue: typeof options.resolveValue === 'function' ? options.resolveValue : null,
          };

          if (elements.confirmButton) {
            elements.confirmButton.onclick = () => {
              if (activeAppDialog && typeof activeAppDialog.resolveValue === 'function') {
                const resolved = activeAppDialog.resolveValue(elements);
                if (resolved === false || typeof resolved === 'undefined') {
                  return;
                }
                closeAppDialog(resolved);
                return;
              }
              closeAppDialog(true);
            };
          }
          if (elements.cancelButton) {
            elements.cancelButton.onclick = () => closeAppDialog(false);
          }
          if (elements.backdrop) {
            elements.backdrop.onclick = () => {
              if (activeAppDialog && activeAppDialog.dismissible) {
                closeAppDialog(false);
              }
            };
          }

          window.setTimeout(() => {
            const target = (!cancelLabel || variant === 'danger') && elements.confirmButton
              ? elements.confirmButton
              : elements.cancelButton || elements.confirmButton;
            if (target) {
              target.focus();
            }
          }, 30);
        });
      }

      async function appAlert(message, options = {}) {
        return showAppDialog({
          title: options.title || 'Notifica',
          message,
          badge: options.badge || 'Notifica',
          variant: options.variant || 'info',
          confirmLabel: options.confirmLabel || 'Chiudi',
          dismissible: options.dismissible !== false,
        });
      }

      async function appConfirm(message, options = {}) {
        return showAppDialog({
          title: options.title || 'Conferma operazione',
          message,
          badge: options.badge || 'Conferma',
          variant: options.variant || 'warning',
          confirmLabel: options.confirmLabel || 'Conferma',
          cancelLabel: options.cancelLabel || 'Annulla',
          dismissible: options.dismissible !== false,
        });
      }
    """


def public_page(title: str, content: str, query_params: dict[str, str]) -> bytes:
    document = f"""<!doctype html>
<html lang="it">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{esc(title)} - {esc(APP_NAME)}</title>
    <link rel="icon" type="image/x-icon" href="/static/logo-ca.ico">
    <link rel="shortcut icon" href="/static/logo-ca.ico">
    <link rel="stylesheet" href="/static/style.css">
  </head>
  <body class="auth-body">
    <div class="auth-shell">
      <section class="auth-brand-panel">
        <div class="auth-brand-card">
          <div class="auth-brand-mark">
            <div class="auth-brand-logo-wrap">
              <img class="auth-brand-logo" src="{LOGO_URL}" alt="Logo {APP_NAME}">
            </div>
            <div>
              <span class="eyebrow">Accesso riservato</span>
              <h1>{esc(APP_NAME)}</h1>
            </div>
          </div>
          <p class="auth-brand-text">Un unico spazio per anagrafiche, tesseramenti, corsi, campo estivo, eventi, incassi e ricevute.</p>
          <div class="auth-feature-list">
            <div class="auth-feature-item">Login protetto con username e password</div>
            <div class="auth-feature-item">Cruscotto unico per quote, scadenze, incassi e ricevute</div>
            <div class="auth-feature-item">Interfaccia chiara, pronta per il lavoro quotidiano</div>
          </div>
        </div>
      </section>
      <main class="auth-content">
        {message_banner(query_params)}
        {content}
      </main>
    </div>
    {app_dialog_markup()}
    <script>
      {shared_dialog_script()}
    </script>
  </body>
</html>"""
    return document.encode("utf-8")


def page(
    title: str,
    current_path: str,
    content: str,
    query_params: dict[str, str],
    current_user: dict[str, object] | None = None,
) -> bytes:
    header_class = "page-header home-header" if current_path == "/" else "page-header"
    header_tools = header_work_year_selector(query_params, action="/") if current_path == "/" else ""
    payment_methods_json = json.dumps(
        [{"id": str(row["id"]), "label": plain_text(row["label"])} for row in metodi_options()],
        ensure_ascii=False,
    )
    document = f"""<!doctype html>
<html lang="it">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{esc(title)} - {esc(APP_NAME)}</title>
    <link rel="icon" type="image/x-icon" href="/static/logo-ca.ico">
    <link rel="shortcut icon" href="/static/logo-ca.ico">
    <link rel="stylesheet" href="/static/style.css">
  </head>
  <body>
    <div class="app-shell">
      {render_navigation(current_path, query_params, current_user)}
      <main class="content">
        <header class="{header_class}">
          <div class="page-title-block">
            <span class="eyebrow">{esc(APP_NAME)}</span>
            <h1>{esc(title)}</h1>
          </div>
          {header_tools}
        </header>
        {message_banner(query_params)}
        {content}
      </main>
    </div>
    {app_dialog_markup()}
    <script>
      {shared_dialog_script()}
      const paymentMethodChoices = {payment_methods_json};

      function setFormHiddenValue(form, fieldName, value) {{
        if (!form || !fieldName) {{
          return;
        }}
        let hidden = form.querySelector(`input[name="${{fieldName}}"]`);
        if (!hidden) {{
          hidden = document.createElement('input');
          hidden.type = 'hidden';
          hidden.name = fieldName;
          form.appendChild(hidden);
        }}
        hidden.value = value;
      }}

      function paymentMethodOptionsHtml(selectedValue) {{
        const selected = selectedValue || '';
        return paymentMethodChoices.map((choice) => {{
          const isSelected = String(choice.id) === String(selected) ? ' selected' : '';
          return `<option value="${{choice.id}}"${{isSelected}}>${{choice.label}}</option>`;
        }}).join('');
      }}

      function parseCourseMonthValue(value) {{
        const raw = String(value || '').trim();
        const match = raw.match(/^(\\d{{4}})-(\\d{{2}})(?:-\\d{{2}})?$/);
        if (!match) {{
          return null;
        }}
        const year = Number.parseInt(match[1], 10);
        const month = Number.parseInt(match[2], 10);
        if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {{
          return null;
        }}
        return {{ year, month }};
      }}

      function formatCourseMonthLabel(year, month) {{
        const safeYear = Number.parseInt(year, 10);
        const safeMonth = Number.parseInt(month, 10);
        if (!Number.isInteger(safeYear) || !Number.isInteger(safeMonth) || safeMonth < 1 || safeMonth > 12) {{
          return '';
        }}
        try {{
          const formatter = new Intl.DateTimeFormat('it-IT', {{ month: 'long', year: 'numeric' }});
          const base = formatter.format(new Date(safeYear, safeMonth - 1, 1));
          return base.charAt(0).toUpperCase() + base.slice(1);
        }} catch (_) {{
          return `${{String(safeMonth).padStart(2, '0')}}/${{safeYear}}`;
        }}
      }}

      function courseMonthSpan(startYear, startMonth, endYear, endMonth) {{
        return ((endYear - startYear) * 12) + (endMonth - startMonth) + 1;
      }}

      async function appPromptEnrollmentPayment(options = {{}}) {{
        const amountValue = options.defaultAmount ? String(options.defaultAmount) : '';
        const defaultMethodId = options.defaultMethodId || '';
        return showAppDialog({{
          title: options.title || 'Registrazione pagamento',
          message: options.message || 'Conferma i dati del pagamento da registrare.',
          badge: 'Pagamento',
          variant: 'warning',
          confirmLabel: options.confirmLabel || 'Registra pagamento',
          cancelLabel: 'Annulla',
          dismissible: false,
          html: `
            <div class="payment-flow-grid">
              <label class="field wide">
                <span>Metodo</span>
                <select class="control" name="dialog_metodo_pagamento_id">
                  ${{paymentMethodOptionsHtml(defaultMethodId)}}
                </select>
              </label>
              <label class="field wide">
                <span>Importo pagato</span>
                <input class="control" type="number" name="dialog_importo_pagato" step="0.01" min="0.01" value="${{amountValue}}">
              </label>
              <p class="payment-flow-error" hidden></p>
            </div>
          `,
          resolveValue: (elements) => {{
            const extra = elements.extra;
            if (!extra) {{
              return false;
            }}
            const methodField = extra.querySelector('[name="dialog_metodo_pagamento_id"]');
            const amountField = extra.querySelector('[name="dialog_importo_pagato"]');
            const errorField = extra.querySelector('.payment-flow-error');
            const methodValue = methodField ? String(methodField.value || '').trim() : '';
            const amountValueRaw = amountField ? String(amountField.value || '').trim() : '';
            const amountNumber = Number.parseFloat(amountValueRaw.replace(',', '.'));
            let errorMessage = '';
            if (!methodValue) {{
              errorMessage = 'Seleziona un metodo di pagamento.';
            }} else if (!Number.isFinite(amountNumber) || amountNumber <= 0) {{
              errorMessage = 'Indica un importo pagato valido.';
            }}
            if (errorField) {{
              if (errorMessage) {{
                errorField.hidden = false;
                errorField.textContent = errorMessage;
              }} else {{
                errorField.hidden = true;
                errorField.textContent = '';
              }}
            }}
            if (errorMessage) {{
              return false;
            }}
            return {{
              confirmed: true,
              methodId: methodValue,
              importo: amountNumber.toFixed(2),
            }};
          }},
        }});
      }}

      function syncCourseEnrollmentPaymentDialog(extra, options = {{}}) {{
        if (!extra) {{
          return;
        }}
        const scopeField = extra.querySelector('[name="dialog_payment_scope"]');
        const untilWrap = extra.querySelector('.payment-flow-until');
        const untilField = extra.querySelector('[name="dialog_fino_competenza"]');
        const amountField = extra.querySelector('[name="dialog_importo_pagato"]');
        const summaryField = extra.querySelector('[data-payment-summary]');
        const errorField = extra.querySelector('.payment-flow-error');
        if (!scopeField || !untilWrap || !untilField || !amountField || !summaryField) {{
          return;
        }}
        const monthlyAmount = Number.parseFloat(String(options.monthlyAmount || '0').replace(',', '.'));
        const startYear = Number.parseInt(options.startYear, 10);
        const startMonth = Number.parseInt(options.startMonth, 10);
        const startLabel = formatCourseMonthLabel(startYear, startMonth);
        const scopeValue = scopeField.value === 'mensilita-future' ? 'mensilita-future' : 'mese-iscrizione';
        untilWrap.hidden = scopeValue !== 'mensilita-future';
        const endPeriod = parseCourseMonthValue(untilField.value) || {{ year: startYear, month: startMonth }};
        const span = courseMonthSpan(startYear, startMonth, endPeriod.year, endPeriod.month);
        const safeSpan = span > 0 ? span : 1;
        const suggestedAmount = Number.isFinite(monthlyAmount)
          ? (monthlyAmount * safeSpan).toFixed(2)
          : '';
        if (amountField.dataset.userEdited !== '1' || !String(amountField.value || '').trim()) {{
          amountField.value = suggestedAmount;
        }}
        amountField.dataset.suggestedValue = suggestedAmount;
        if (errorField) {{
          errorField.hidden = true;
          errorField.textContent = '';
        }}
        if (scopeValue === 'mensilita-future') {{
          if (span < 1) {{
            summaryField.innerHTML = `Seleziona un mese finale uguale o successivo a <strong>${{startLabel}}</strong>.`;
          }} else {{
            const endLabel = formatCourseMonthLabel(endPeriod.year, endPeriod.month);
            summaryField.innerHTML = `Verranno generate automaticamente <strong>${{safeSpan}}</strong> mensilita da <strong>${{startLabel}}</strong> a <strong>${{endLabel}}</strong>.`;
          }}
        }} else {{
          summaryField.innerHTML = `Verrà generata e proposta la quota del mese di iscrizione: <strong>${{startLabel}}</strong>.`;
        }}
      }}

      async function appPromptCourseEnrollmentPayment(form, options = {{}}) {{
        const enrollmentField = form ? form.querySelector('[name="data_iscrizione"]') : null;
        const enrollmentPeriod = parseCourseMonthValue(enrollmentField ? enrollmentField.value : '') || parseCourseMonthValue(new Date().toISOString().slice(0, 10));
        const startYear = enrollmentPeriod ? enrollmentPeriod.year : new Date().getFullYear();
        const startMonth = enrollmentPeriod ? enrollmentPeriod.month : (new Date().getMonth() + 1);
        const startMonthValue = `${{startYear}}-${{String(startMonth).padStart(2, '0')}}`;
        const monthlyAmount = Number.parseFloat(String(options.defaultAmount || '0').replace(',', '.'));
        const dialogPromise = showAppDialog({{
          title: options.title || 'Pagamento quote corso',
          message: options.message || 'Scegli se saldare il solo mese di iscrizione oppure anche mensilita future.',
          badge: 'Corso',
          variant: 'warning',
          confirmLabel: options.confirmLabel || 'Registra pagamento',
          cancelLabel: 'Annulla',
          dismissible: false,
          html: `
            <div class="payment-flow-grid">
              <label class="field wide">
                <span>Mensilita da pagare</span>
                <select class="control" name="dialog_payment_scope">
                  <option value="mese-iscrizione">Solo mese di iscrizione</option>
                  <option value="mensilita-future">Anche mensilita future</option>
                </select>
              </label>
              <label class="field wide payment-flow-until" hidden>
                <span>Fino al mese di competenza</span>
                <input class="control" type="month" name="dialog_fino_competenza" value="${{startMonthValue}}" min="${{startMonthValue}}">
              </label>
              <div class="payment-flow-note" data-payment-summary></div>
              <label class="field wide">
                <span>Metodo</span>
                <select class="control" name="dialog_metodo_pagamento_id">
                  ${{paymentMethodOptionsHtml(options.defaultMethodId || '')}}
                </select>
              </label>
              <label class="field wide">
                <span>Importo pagato</span>
                <input class="control" type="number" name="dialog_importo_pagato" step="0.01" min="0.01" value="${{Number.isFinite(monthlyAmount) ? monthlyAmount.toFixed(2) : ''}}">
              </label>
              <p class="payment-flow-error" hidden></p>
            </div>
          `,
          resolveValue: (elements) => {{
            const extra = elements.extra;
            if (!extra) {{
              return false;
            }}
            const scopeField = extra.querySelector('[name="dialog_payment_scope"]');
            const untilField = extra.querySelector('[name="dialog_fino_competenza"]');
            const methodField = extra.querySelector('[name="dialog_metodo_pagamento_id"]');
            const amountField = extra.querySelector('[name="dialog_importo_pagato"]');
            const errorField = extra.querySelector('.payment-flow-error');
            const scopeValue = scopeField && scopeField.value === 'mensilita-future' ? 'mensilita-future' : 'mese-iscrizione';
            const methodValue = methodField ? String(methodField.value || '').trim() : '';
            const amountValueRaw = amountField ? String(amountField.value || '').trim() : '';
            const amountNumber = Number.parseFloat(amountValueRaw.replace(',', '.'));
            const endPeriod = scopeValue === 'mensilita-future'
              ? parseCourseMonthValue(untilField ? untilField.value : '')
              : {{ year: startYear, month: startMonth }};
            let errorMessage = '';
            if (!methodValue) {{
              errorMessage = 'Seleziona un metodo di pagamento.';
            }} else if (!Number.isFinite(amountNumber) || amountNumber <= 0) {{
              errorMessage = 'Indica un importo pagato valido.';
            }} else if (!endPeriod) {{
              errorMessage = 'Seleziona il mese finale delle quote da generare.';
            }} else if (courseMonthSpan(startYear, startMonth, endPeriod.year, endPeriod.month) < 1) {{
              errorMessage = 'Il mese finale deve essere uguale o successivo al mese di iscrizione.';
            }}
            if (errorField) {{
              if (errorMessage) {{
                errorField.hidden = false;
                errorField.textContent = errorMessage;
              }} else {{
                errorField.hidden = true;
                errorField.textContent = '';
              }}
            }}
            if (errorMessage) {{
              return false;
            }}
            const untilCompetenza = `${{endPeriod.year}}-${{String(endPeriod.month).padStart(2, '0')}}`;
            return {{
              confirmed: true,
              methodId: methodValue,
              importo: amountNumber.toFixed(2),
              scope: scopeValue,
              untilCompetenza,
            }};
          }},
        }});

        window.setTimeout(() => {{
          const elements = appDialogElements();
          const extra = elements.extra;
          if (!extra) {{
            return;
          }}
          const scopeField = extra.querySelector('[name="dialog_payment_scope"]');
          const untilField = extra.querySelector('[name="dialog_fino_competenza"]');
          const amountField = extra.querySelector('[name="dialog_importo_pagato"]');
          if (!scopeField || !untilField || !amountField) {{
            return;
          }}
          const syncState = () => syncCourseEnrollmentPaymentDialog(extra, {{
            startYear,
            startMonth,
            monthlyAmount,
          }});
          amountField.addEventListener('input', () => {{
            amountField.dataset.userEdited = '1';
          }});
          scopeField.addEventListener('change', () => {{
            amountField.dataset.userEdited = '0';
            syncState();
          }});
          untilField.addEventListener('change', () => {{
            amountField.dataset.userEdited = '0';
            syncState();
          }});
          syncState();
        }}, 0);

        return dialogPromise;
      }}

      async function handleEnrollmentPaymentFlow(event) {{
        const form = event.currentTarget;
        if (!form || form.dataset.paymentFlowSubmitting === '1') {{
          return;
        }}
        event.preventDefault();
        const amountFieldName = form.dataset.paymentAmountField || '';
        const amountField = amountFieldName ? form.querySelector(`[name="${{amountFieldName}}"]`) : null;
        const defaultAmount = amountField ? amountField.value : '';
        const wantsPayment = await appConfirm(
          form.dataset.paymentPromptMessage || 'Vuoi procedere anche al pagamento?',
          {{
            title: form.dataset.paymentPromptTitle || 'Conferma iscrizione',
            badge: 'Iscrizione',
            confirmLabel: form.dataset.paymentPromptYes || 'Si, procedi',
            cancelLabel: form.dataset.paymentPromptNo || 'No, solo iscrizione',
          }}
        );
        if (!wantsPayment) {{
          setFormHiddenValue(form, 'procedi_pagamento', '0');
          setFormHiddenValue(form, 'pagamento_metodo_id', '');
          setFormHiddenValue(form, 'pagamento_importo', '');
          setFormHiddenValue(form, 'pagamento_scope', '');
          setFormHiddenValue(form, 'pagamento_competenza_fine', '');
          form.dataset.paymentFlowSubmitting = '1';
          form.submit();
          return;
        }}

        const paymentFlowKind = String(form.dataset.paymentFlow || '').trim();
        const paymentResult = paymentFlowKind === 'corso'
          ? await appPromptCourseEnrollmentPayment(form, {{
              title: form.dataset.paymentDialogTitle || 'Pagamento quote corso',
              message: form.dataset.paymentDialogMessage || 'Scegli se pagare solo il mese di iscrizione oppure anche mensilita future.',
              confirmLabel: form.dataset.paymentDialogConfirm || 'Registra pagamento',
              defaultAmount,
              defaultMethodId: form.dataset.paymentMethodDefault || '',
            }})
          : await appPromptEnrollmentPayment({{
              title: form.dataset.paymentDialogTitle || 'Pagamento',
              message: form.dataset.paymentDialogMessage || 'Conferma il pagamento da registrare.',
              confirmLabel: form.dataset.paymentDialogConfirm || 'Registra pagamento',
              defaultAmount,
              defaultMethodId: form.dataset.paymentMethodDefault || '',
            }});
        if (!paymentResult || paymentResult.confirmed !== true) {{
          return;
        }}

        setFormHiddenValue(form, 'procedi_pagamento', '1');
        setFormHiddenValue(form, 'pagamento_metodo_id', paymentResult.methodId);
        setFormHiddenValue(form, 'pagamento_importo', paymentResult.importo);
        setFormHiddenValue(form, 'pagamento_scope', paymentResult.scope || '');
        setFormHiddenValue(form, 'pagamento_competenza_fine', paymentResult.untilCompetenza || '');
        form.dataset.paymentFlowSubmitting = '1';
        form.submit();
      }}

      function bindEnrollmentPaymentFlows() {{
        document.querySelectorAll('form[data-payment-flow]').forEach((form) => {{
          if (form.dataset.paymentFlowBound === '1') {{
            return;
          }}
          form.dataset.paymentFlowBound = '1';
          form.addEventListener('submit', handleEnrollmentPaymentFlow);
        }});
      }}

      function syncDeleteFormView(form) {{
        if (!form) {{
          return;
        }}
        const vistaField = form.querySelector('input[name="vista"]');
        if (!vistaField) {{
          return;
        }}
        const params = new URLSearchParams(window.location.search);
        vistaField.value = params.get('vista') || '';
      }}

      function filterRowsBySelector(input, selector) {{
        const needle = (input.value || '').toLowerCase().trim();
        document.querySelectorAll(selector + ' tbody tr').forEach((row) => {{
          const haystack = (row.textContent || '').toLowerCase();
          row.style.display = !needle || haystack.includes(needle) ? '' : 'none';
        }});
        document.querySelectorAll(selector).forEach((table) => updateVisibleTableTotals(table));
      }}

      function formatSummaryMoney(value) {{
        const amount = Number.isFinite(value) ? value : 0;
        return amount.toLocaleString('it-IT', {{
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }}) + ' EUR';
      }}

      function updateVisibleTableTotals(table) {{
        if (!table) {{
          return;
        }}
        const summaryColumns = (table.dataset.summaryColumns || '')
          .split(',')
          .map((value) => Number.parseInt(value, 10))
          .filter((value) => Number.isInteger(value));
        if (!summaryColumns.length) {{
          return;
        }}
        const footerRow = table.querySelector('tfoot tr');
        if (!footerRow) {{
          return;
        }}
        const labelIndex = Number.parseInt(table.dataset.summaryLabelIndex || '0', 10);
        const footerCells = Array.from(footerRow.children);
        footerCells.forEach((cell, index) => {{
          if (index === labelIndex) {{
            cell.textContent = 'Totali';
            return;
          }}
          if (!summaryColumns.includes(index)) {{
            cell.textContent = '';
          }}
        }});
        summaryColumns.forEach((columnIndex) => {{
          let total = 0;
          let found = false;
          table.querySelectorAll('tbody tr').forEach((row) => {{
            if (row.style.display === 'none') {{
              return;
            }}
            const cell = row.children[columnIndex];
            if (!cell || !cell.dataset.sumValue) {{
              return;
            }}
            const amount = Number.parseFloat(cell.dataset.sumValue);
            if (!Number.isFinite(amount)) {{
              return;
            }}
            total += amount;
            found = true;
          }});
          if (footerCells[columnIndex]) {{
            footerCells[columnIndex].textContent = found ? formatSummaryMoney(total) : '';
          }}
        }});
      }}

      function syncToolbarSearchLinks(input) {{
        const toolbar = input.closest('.report-toolbar');
        if (!toolbar) {{
          return;
        }}
        toolbar.querySelectorAll('[data-search-link]').forEach((link) => {{
          const url = new URL(link.href, window.location.origin);
          if (input.value) {{
            url.searchParams.set('search', input.value);
          }} else {{
            url.searchParams.delete('search');
          }}
          link.href = url.pathname + url.search;
        }});
      }}

      function handleReportSearch(input) {{
        filterRowsBySelector(input, '.report-table');
        syncToolbarSearchLinks(input);
      }}

      function handleDataSearch(input) {{
        filterRowsBySelector(input, '.search-table');
      }}

      function normalizeSearchValue(value) {{
        return (value || '').toLowerCase().trim();
      }}

      function getAutocompleteElements(input) {{
        const targetId = input.dataset.selectSearchTarget || '';
        const panelId = input.dataset.selectSearchPanel || '';
        return {{
          select: targetId ? document.getElementById(targetId) : null,
          panel: panelId ? document.getElementById(panelId) : null,
        }};
      }}

      function autocompleteOptions(select) {{
        return Array.from(select.options).filter((option) => !option.disabled && option.value);
      }}

      function autocompleteDisplayLabel(option) {{
        return option.dataset.autocompleteLabel || option.textContent || '';
      }}

      function autocompleteSearchText(option) {{
        return normalizeSearchValue(option.dataset.searchText || option.textContent || '');
      }}

      function autocompleteMatches(option, query) {{
        const visibleText = normalizeSearchValue(option.textContent || '');
        const searchText = autocompleteSearchText(option);
        const displayText = normalizeSearchValue(autocompleteDisplayLabel(option));
        return (
          visibleText.startsWith(query)
          || searchText.startsWith(query)
          || displayText.startsWith(query)
          || visibleText.includes(query)
          || searchText.includes(query)
          || displayText.includes(query)
        );
      }}

      function findAutocompleteOption(select, value) {{
        return autocompleteOptions(select).find((option) => option.value === value) || null;
      }}

      function autocompleteItems(panel) {{
        return Array.from(panel.querySelectorAll('.select-autocomplete-item'));
      }}

      function setActiveAutocompleteItem(panel, nextIndex) {{
        const items = autocompleteItems(panel);
        if (!items.length) {{
          panel.dataset.activeIndex = '-1';
          return;
        }}
        const boundedIndex = Math.max(0, Math.min(nextIndex, items.length - 1));
        items.forEach((item, index) => {{
          item.classList.toggle('active', index === boundedIndex);
        }});
        panel.dataset.activeIndex = String(boundedIndex);
      }}

      function closeSelectAutocomplete(input) {{
        const {{ panel }} = getAutocompleteElements(input);
        if (!panel) {{
          return;
        }}
        panel.hidden = true;
        panel.replaceChildren();
        panel.dataset.activeIndex = '-1';
      }}

      function selectAutocompleteValue(input, optionValue) {{
        const {{ select }} = getAutocompleteElements(input);
        if (!select) {{
          return false;
        }}
        const option = findAutocompleteOption(select, optionValue);
        if (!option) {{
          return false;
        }}
        select.value = option.value;
        input.value = autocompleteDisplayLabel(option);
        input.setCustomValidity('');
        select.dispatchEvent(new Event('change', {{ bubbles: true }}));
        closeSelectAutocomplete(input);
        return true;
      }}

      function clearAutocompleteValue(input) {{
        const {{ select }} = getAutocompleteElements(input);
        if (!select) {{
          return;
        }}
        select.value = '';
      }}

      function renderSelectAutocomplete(input, showAllOnEmpty = false) {{
        const {{ select, panel }} = getAutocompleteElements(input);
        if (!select || !panel) {{
          return [];
        }}

        const query = normalizeSearchValue(input.value);
        if (!query && !showAllOnEmpty) {{
          closeSelectAutocomplete(input);
          return [];
        }}

        const matches = autocompleteOptions(select)
          .filter((option) => !query || autocompleteMatches(option, query))
          .slice(0, 8);

        panel.replaceChildren();
        if (!matches.length) {{
          const emptyState = document.createElement('div');
          emptyState.className = 'select-autocomplete-empty';
          emptyState.textContent = 'Nessun associato trovato';
          panel.appendChild(emptyState);
          panel.hidden = false;
          panel.dataset.activeIndex = '-1';
          return [];
        }}

        matches.forEach((option, index) => {{
          const item = document.createElement('button');
          item.type = 'button';
          item.className = 'select-autocomplete-item';
          item.dataset.optionValue = option.value;
          item.textContent = autocompleteDisplayLabel(option);
          item.addEventListener('mousedown', (event) => {{
            event.preventDefault();
          }});
          item.addEventListener('click', () => {{
            selectAutocompleteValue(input, option.value);
          }});
          panel.appendChild(item);
          if (index === 0) {{
            item.classList.add('active');
          }}
        }});

        panel.hidden = false;
        panel.dataset.activeIndex = matches.length ? '0' : '-1';
        return matches;
      }}

      function syncAutocompleteInput(input) {{
        const {{ select }} = getAutocompleteElements(input);
        if (!select) {{
          return;
        }}
        const option = select.options[select.selectedIndex];
        input.value = option && option.value ? autocompleteDisplayLabel(option) : '';
      }}

      function commitSelectAutocomplete(input, strictRequired = false) {{
        const {{ select }} = getAutocompleteElements(input);
        if (!select) {{
          return true;
        }}

        const query = normalizeSearchValue(input.value);
        input.setCustomValidity('');
        if (!query) {{
          clearAutocompleteValue(input);
          closeSelectAutocomplete(input);
          if (strictRequired && input.required) {{
            input.setCustomValidity('Seleziona un associato valido.');
            return false;
          }}
          return true;
        }}

        const current = select.options[select.selectedIndex];
        if (current && current.value && autocompleteMatches(current, query)) {{
          input.value = autocompleteDisplayLabel(current);
          closeSelectAutocomplete(input);
          return true;
        }}

        const matches = renderSelectAutocomplete(input);
        if (matches.length) {{
          return selectAutocompleteValue(input, matches[0].value);
        }}

        clearAutocompleteValue(input);
        closeSelectAutocomplete(input);
        if (strictRequired && input.required) {{
          input.setCustomValidity('Seleziona un associato valido.');
          return false;
        }}
        return true;
      }}

      function openSelectAutocomplete(input) {{
        renderSelectAutocomplete(input, true);
      }}

      function handleSelectSearch(input) {{
        input.setCustomValidity('');
        if (!normalizeSearchValue(input.value)) {{
          clearAutocompleteValue(input);
          closeSelectAutocomplete(input);
          return;
        }}
        renderSelectAutocomplete(input);
      }}

      function handleSelectSearchKeydown(event, input) {{
        const {{ panel }} = getAutocompleteElements(input);
        if (!panel) {{
          return;
        }}

        if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {{
          event.preventDefault();
          if (panel.hidden) {{
            renderSelectAutocomplete(input, true);
          }}
          const items = autocompleteItems(panel);
          if (!items.length) {{
            return;
          }}
          const currentIndex = parseInt(panel.dataset.activeIndex || '0', 10);
          const delta = event.key === 'ArrowDown' ? 1 : -1;
          setActiveAutocompleteItem(panel, currentIndex + delta);
          return;
        }}

        if (event.key === 'Enter') {{
          const items = autocompleteItems(panel);
          if (!panel.hidden && items.length) {{
            event.preventDefault();
            const activeIndex = parseInt(panel.dataset.activeIndex || '0', 10);
            const activeItem = items[Math.max(0, Math.min(activeIndex, items.length - 1))];
            if (activeItem) {{
              selectAutocompleteValue(input, activeItem.dataset.optionValue || '');
            }}
          }}
          return;
        }}

        if (event.key === 'Escape') {{
          closeSelectAutocomplete(input);
        }}
      }}

      function closeSelectAutocompleteLater(input) {{
        window.setTimeout(() => {{
          commitSelectAutocomplete(input);
        }}, 120);
      }}

      function setInputValue(inputId, value) {{
        const target = document.getElementById(inputId);
        if (!target) {{
          return;
        }}
        target.value = value || '';
      }}

      function syncResidualInput(select) {{
        const targetId = select.dataset.residualTarget || '';
        if (!targetId) {{
          return;
        }}
        const option = select.options[select.selectedIndex];
        const residuo = option && option.dataset ? option.dataset.residuo || '' : '';
        setInputValue(targetId, residuo);
      }}

      function syncStandardQuota(select) {{
        const targetId = select.dataset.standardTarget || '';
        if (!targetId) {{
          return;
        }}
        const option = select.options[select.selectedIndex];
        const quota = option && option.dataset ? option.dataset.quotaStandard || '' : '';
        setInputValue(targetId, quota);
      }}

      function syncAmountProposal(select) {{
        const targetId = select.dataset.amountTarget || '';
        if (!targetId) {{
          return;
        }}
        const option = select.options[select.selectedIndex];
        const amount = option && option.dataset ? option.dataset.importo || '' : '';
        setInputValue(targetId, amount);
      }}

      function syncYearEndDate(input) {{
        const targetId = input.dataset.yearEndTarget || '';
        if (!targetId) {{
          return;
        }}
        const normalized = (input.value || '').replace(/\\D/g, '').slice(0, 4);
        input.value = normalized;
        if (normalized.length !== 4) {{
          return;
        }}
        setInputValue(targetId, `${{normalized}}-12-31`);
      }}

      function syncMonthEndDate(control) {{
        const targetId = control.dataset.monthEndTarget || '';
        const yearSourceId = control.dataset.monthEndYearSource || '';
        const monthSourceId = control.dataset.monthEndMonthSource || '';
        if (!targetId || !yearSourceId || !monthSourceId) {{
          return;
        }}
        const yearSource = document.getElementById(yearSourceId);
        const monthSource = document.getElementById(monthSourceId);
        if (!yearSource || !monthSource) {{
          return;
        }}
        const yearValue = (yearSource.value || '').replace(/\D/g, '').slice(0, 4);
        yearSource.value = yearValue;
        const monthValue = Number.parseInt(monthSource.value || '', 10);
        if (yearValue.length !== 4 || !Number.isInteger(monthValue) || monthValue < 1 || monthValue > 12) {{
          return;
        }}
        const lastDay = new Date(Number.parseInt(yearValue, 10), monthValue, 0).getDate();
        setInputValue(targetId, `${{yearValue}}-${{String(monthValue).padStart(2, '0')}}-${{String(lastDay).padStart(2, '0')}}`);
      }}

      function findTypeaheadMatch(options, buffer) {{
        const normalized = (buffer || '').toLowerCase().trim();
        if (!normalized) {{
          return null;
        }}
        return (
          options.find((option) => {{
            const visibleText = (option.textContent || '').toLowerCase();
            const searchText = (option.dataset.searchText || '').toLowerCase();
            return visibleText.startsWith(normalized) || searchText.startsWith(normalized);
          }})
          || options.find((option) => {{
            const visibleText = (option.textContent || '').toLowerCase();
            const searchText = (option.dataset.searchText || '').toLowerCase();
            return visibleText.includes(normalized) || searchText.includes(normalized);
          }})
          || null
        );
      }}

      function installSelectTypeahead(select) {{
        let buffer = '';
        let resetTimer = null;
        select.addEventListener('keydown', (event) => {{
          if (event.key.length !== 1 || event.altKey || event.ctrlKey || event.metaKey) {{
            return;
          }}
          buffer += event.key.toLowerCase();
          window.clearTimeout(resetTimer);
          resetTimer = window.setTimeout(() => {{
            buffer = '';
          }}, 700);

          const options = Array.from(select.options).filter((option) => !option.disabled && !option.hidden && option.value);
          const match = findTypeaheadMatch(options, buffer)
            || findTypeaheadMatch(options, event.key.toLowerCase());
          if (!match) {{
            return;
          }}

          event.preventDefault();
          if (select.multiple) {{
            Array.from(select.options).forEach((option) => {{
              if (!option.hidden && !option.disabled) {{
                option.selected = false;
              }}
            }});
            match.selected = true;
          }} else {{
            select.value = match.value;
          }}
          match.scrollIntoView({{ block: 'nearest' }});
          select.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }});
        select.addEventListener('blur', () => {{
          buffer = '';
          window.clearTimeout(resetTimer);
        }});
      }}

      function updateMonthlyRates() {{
        const associatoSelect = document.getElementById('monthly-payment-associato');
        const rateSelect = document.getElementById('monthly-rate-select');
        const amountInput = document.getElementById('monthly-payment-amount');
        if (!associatoSelect || !rateSelect || !amountInput) {{
          return;
        }}

        const associatoId = associatoSelect.value;
        let total = 0;
        let visibleCount = 0;
        Array.from(rateSelect.options).forEach((option) => {{
          const visible = !!associatoId && option.dataset.associatoId === associatoId;
          option.hidden = !visible;
          option.disabled = !visible;
          option.selected = visible;
          if (visible) {{
            total += parseFloat(option.dataset.residuo || '0');
            visibleCount += 1;
          }}
        }});
        amountInput.value = visibleCount ? total.toFixed(2) : '';
      }}

      function syncMonthlySelectedAmount() {{
        const rateSelect = document.getElementById('monthly-rate-select');
        const amountInput = document.getElementById('monthly-payment-amount');
        if (!rateSelect || !amountInput) {{
          return;
        }}
        const total = Array.from(rateSelect.selectedOptions)
          .filter((option) => !option.hidden)
          .reduce((sum, option) => sum + parseFloat(option.dataset.residuo || '0'), 0);
        amountInput.value = total ? total.toFixed(2) : '';
      }}

      function updateMultiAreaScadenze(selectedValues = null) {{
        const associatoSelect = document.getElementById('multi-area-associato');
        const scadenzeSelect = document.getElementById('multi-area-scadenze');
        const amountInput = document.getElementById('multi-area-amount');
        if (!associatoSelect || !scadenzeSelect || !amountInput) {{
          return;
        }}

        const associatoId = associatoSelect.value;
        const selectedSet = new Set(Array.isArray(selectedValues) ? selectedValues : []);
        const hasSelectionOverride = Array.isArray(selectedValues);
        let total = 0;
        Array.from(scadenzeSelect.options).forEach((option) => {{
          const visible = !!associatoId && option.dataset.associatoId === associatoId;
          option.hidden = !visible;
          option.disabled = !visible;
          if (!visible) {{
            option.selected = false;
            return;
          }}
          if (hasSelectionOverride) {{
            option.selected = selectedSet.has(option.value);
          }}
          if (option.selected) {{
            total += parseFloat(option.dataset.residuo || '0');
          }}
        }});
        amountInput.value = total ? total.toFixed(2) : '';
      }}

      function syncMultiAreaSelectedAmount() {{
        const scadenzeSelect = document.getElementById('multi-area-scadenze');
        const amountInput = document.getElementById('multi-area-amount');
        if (!scadenzeSelect || !amountInput) {{
          return;
        }}
        const total = Array.from(scadenzeSelect.selectedOptions)
          .filter((option) => !option.hidden)
          .reduce((sum, option) => sum + parseFloat(option.dataset.residuo || '0'), 0);
        amountInput.value = total ? total.toFixed(2) : '';
      }}

      function multiAreaSelectedScadenze() {{
        const scadenzeSelect = document.getElementById('multi-area-scadenze');
        if (!scadenzeSelect) {{
          return [];
        }}
        return Array.from(scadenzeSelect.selectedOptions)
          .filter((option) => !option.hidden)
          .map((option) => option.value);
      }}

      function replaceMultiAreaScadenzeOptions(rows, selectedValues = []) {{
        const scadenzeSelect = document.getElementById('multi-area-scadenze');
        if (!scadenzeSelect) {{
          return;
        }}
        scadenzeSelect.innerHTML = '';
        (rows || []).forEach((row) => {{
          const option = document.createElement('option');
          option.value = String(row.id || '');
          option.textContent = String(row.label || '');
          option.dataset.associatoId = String(row.associato_id || '');
          option.dataset.residuo = String(row.residuo || '0');
          scadenzeSelect.appendChild(option);
        }});
        updateMultiAreaScadenze(selectedValues);
      }}

      function syncMultiAreaFutureCourseControls() {{
        const associatoSelect = document.getElementById('multi-area-associato');
        const corsoSelect = document.getElementById('multi-area-course-enrollment');
        const untilInput = document.getElementById('multi-area-course-until');
        const actionButton = document.getElementById('multi-area-course-generate');
        const feedback = document.getElementById('multi-area-course-feedback');
        if (!associatoSelect || !corsoSelect || !untilInput || !actionButton) {{
          return;
        }}

        const associatoId = associatoSelect.value;
        let firstVisibleValue = '';
        Array.from(corsoSelect.options).forEach((option) => {{
          if (!option.value) {{
            option.hidden = false;
            option.disabled = false;
            return;
          }}
          const visible = !!associatoId && option.dataset.associatoId === associatoId;
          option.hidden = !visible;
          option.disabled = !visible;
          if (visible && !firstVisibleValue) {{
            firstVisibleValue = option.value;
          }}
        }});

        const selectedOption = corsoSelect.selectedOptions.length ? corsoSelect.selectedOptions[0] : null;
        if (!selectedOption || selectedOption.hidden || selectedOption.disabled) {{
          corsoSelect.value = firstVisibleValue;
        }}

        const activeOption = corsoSelect.selectedOptions.length ? corsoSelect.selectedOptions[0] : null;
        const startCompetenza = activeOption ? String(activeOption.dataset.startCompetenza || '') : '';
        const workYear = String(corsoSelect.dataset.workYear || '');
        const workYearStart = workYear ? `${{workYear}}-01` : '';
        const maxCompetenza = workYear ? `${{workYear}}-12` : '';
        untilInput.disabled = !activeOption || !activeOption.value;
        actionButton.disabled = !associatoId || !activeOption || !activeOption.value || !untilInput.value;
        const effectiveMin = startCompetenza && workYearStart
          ? (startCompetenza > workYearStart ? startCompetenza : workYearStart)
          : (startCompetenza || workYearStart);
        if (effectiveMin) {{
          untilInput.min = effectiveMin;
          if (!untilInput.value || untilInput.value < effectiveMin) {{
            untilInput.value = effectiveMin;
          }}
        }} else {{
          untilInput.min = '';
        }}
        if (maxCompetenza) {{
          untilInput.max = maxCompetenza;
          if (untilInput.value && untilInput.value > maxCompetenza) {{
            untilInput.value = maxCompetenza;
          }}
        }} else {{
          untilInput.max = '';
        }}
        actionButton.disabled = !associatoId || !activeOption || !activeOption.value || !untilInput.value;
        if (feedback && !feedback.dataset.preserveMessage) {{
          feedback.textContent = activeOption && activeOption.value
            ? 'Le mensilita del corso fino al mese indicato verranno selezionate automaticamente in basso.'
            : 'Seleziona prima associato e corso per aggiungere quote future.';
        }}
      }}

      async function generateFutureCourseMultiAreaScadenze() {{
        const associatoSelect = document.getElementById('multi-area-associato');
        const corsoSelect = document.getElementById('multi-area-course-enrollment');
        const untilInput = document.getElementById('multi-area-course-until');
        const actionButton = document.getElementById('multi-area-course-generate');
        const feedback = document.getElementById('multi-area-course-feedback');
        if (!associatoSelect || !corsoSelect || !untilInput || !actionButton) {{
          return false;
        }}

        const associatoId = String(associatoSelect.value || '').trim();
        const iscrizioneCorsoId = String(corsoSelect.value || '').trim();
        const fineCompetenza = String(untilInput.value || '').trim();
        if (!associatoId) {{
          await appAlert('Seleziona prima un associato.', {{ title: 'Quote future corso', badge: 'Pagamenti' }});
          return false;
        }}
        if (!iscrizioneCorsoId) {{
          await appAlert('Seleziona il corso per cui vuoi generare le mensilita future.', {{ title: 'Quote future corso', badge: 'Pagamenti' }});
          return false;
        }}
        if (!fineCompetenza) {{
          await appAlert('Indica fino a quale mensilita vuoi arrivare.', {{ title: 'Quote future corso', badge: 'Pagamenti' }});
          return false;
        }}

        const previousSelected = multiAreaSelectedScadenze();
        const params = new URLSearchParams({{
          anno_lavoro: String(corsoSelect.dataset.workYear || ''),
          associato_id: associatoId,
          iscrizione_corso_id: iscrizioneCorsoId,
          fine_competenza: fineCompetenza,
        }});

        actionButton.disabled = true;
        if (feedback) {{
          feedback.dataset.preserveMessage = '1';
          feedback.textContent = 'Generazione quote in corso...';
        }}
        try {{
          const payload = await fetchJson(`/api/pagamenti-multi-area/quote-corso-future?${{params.toString()}}`);
          if (!payload || payload.ok !== true) {{
            throw new Error(payload && payload.error ? payload.error : 'Impossibile generare le quote future del corso.');
          }}
          const selectedValues = Array.from(new Set([...(previousSelected || []), ...((payload.selected_scadenze || []).map(String))]));
          replaceMultiAreaScadenzeOptions(payload.options || [], selectedValues);
          if (feedback) {{
            feedback.textContent = payload.message || 'Quote corso aggiornate e selezionate.';
          }}
          return false;
        }} catch (error) {{
          if (feedback) {{
            feedback.textContent = error && error.message ? error.message : 'Impossibile generare le quote future del corso.';
          }}
          await appAlert(
            feedback ? feedback.textContent : 'Impossibile generare le quote future del corso.',
            {{ title: 'Quote future corso', badge: 'Pagamenti', variant: 'danger' }}
          );
          return false;
        }} finally {{
          if (feedback) {{
            feedback.dataset.preserveMessage = '';
          }}
          actionButton.disabled = false;
          syncMultiAreaFutureCourseControls();
        }}
      }}

      const codiceFiscaleLookupTimers = new WeakMap();
      const comuneLookupTimers = new WeakMap();
      const capLookupTimers = new WeakMap();
      const cittaLookupTimers = new WeakMap();
      const codiceFiscaleGenerationTimers = new WeakMap();

      function assignFieldValue(form, fieldName, value) {{
        const field = form.querySelector(`[name="${{fieldName}}"]`);
        if (!field) {{
          return;
        }}
        field.value = value || '';
      }}

      function applyCodiceFiscalePayload(input, payload) {{
        const form = input.closest('form');
        if (!form || !payload) {{
          return;
        }}
        assignFieldValue(form, 'data_nascita', payload.data_nascita || '');
        assignFieldValue(form, 'sesso', payload.sesso || '');
        assignFieldValue(form, 'comune_nascita', payload.comune_nascita || '');
        assignFieldValue(form, 'provincia_nascita', payload.provincia_nascita || '');
      }}

      async function fetchJson(url, options = {{}}) {{
        const response = await fetch(url, {{
          ...options,
          headers: {{ 'Accept': 'application/json', ...(options.headers || {{}}) }},
          cache: options.cache || 'no-store',
        }});
        if (!response.ok) {{
          return null;
        }}
        return response.json();
      }}

      function associatoFormField(form, fieldName) {{
        return form ? form.querySelector(`[name="${{fieldName}}"]`) : null;
      }}

      async function lookupCodiceFiscale(input) {{
        const normalized = (input.value || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 16);
        input.value = normalized;
        if (normalized.length !== 16) {{
          return;
        }}
        try {{
          const payload = await fetchJson(`/api/codice-fiscale?value=${{encodeURIComponent(normalized)}}`);
          if (!payload || !payload.found) {{
            return;
          }}
          applyCodiceFiscalePayload(input, payload);
        }} catch (error) {{
          console.warn('Lookup codice fiscale non disponibile', error);
        }}
      }}

      function scheduleCodiceFiscaleLookup(input) {{
        const normalized = (input.value || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 16);
        input.value = normalized;
        const previousTimer = codiceFiscaleLookupTimers.get(input);
        if (previousTimer) {{
          window.clearTimeout(previousTimer);
        }}
        if (normalized.length !== 16) {{
          return;
        }}
        const timer = window.setTimeout(() => {{
          lookupCodiceFiscale(input);
        }}, 180);
        codiceFiscaleLookupTimers.set(input, timer);
      }}

      async function lookupComuneNascita(input) {{
        const form = input.closest('form');
        const comuneField = associatoFormField(form, 'comune_nascita');
        if (!form || !comuneField) {{
          return null;
        }}
        const comune = (comuneField.value || '').trim();
        if (!comune) {{
          return null;
        }}
        try {{
          const payload = await fetchJson(`/api/comuni?nome=${{encodeURIComponent(comune)}}`);
          if (!payload || !payload.found) {{
            return null;
          }}
          assignFieldValue(form, 'comune_nascita', payload.comune || comune);
          assignFieldValue(form, 'provincia_nascita', payload.provincia || '');
          return payload;
        }} catch (error) {{
          console.warn('Lookup comune di nascita non disponibile', error);
          return null;
        }}
      }}

      function scheduleComuneNascitaLookup(input) {{
        const previousTimer = comuneLookupTimers.get(input);
        if (previousTimer) {{
          window.clearTimeout(previousTimer);
        }}
        const timer = window.setTimeout(async () => {{
          await lookupComuneNascita(input);
          await generateCodiceFiscaleFromFields(input.closest('form'));
        }}, 180);
        comuneLookupTimers.set(input, timer);
      }}

      async function lookupCittaDetails(input) {{
        const form = input.closest('form');
        const cittaField = associatoFormField(form, 'citta');
        if (!form || !cittaField) {{
          return null;
        }}
        const citta = (cittaField.value || '').trim();
        if (!citta) {{
          return null;
        }}
        try {{
          const payload = await fetchJson(`/api/comuni?nome=${{encodeURIComponent(citta)}}`);
          if (!payload || !payload.found) {{
            return null;
          }}
          assignFieldValue(form, 'citta', payload.comune || citta);
          assignFieldValue(form, 'provincia', payload.provincia || '');
          if (payload.cap) {{
            assignFieldValue(form, 'cap', payload.cap);
          }}
          return payload;
        }} catch (error) {{
          console.warn('Lookup citta non disponibile', error);
          return null;
        }}
      }}

      function scheduleCittaLookup(input) {{
        const previousTimer = cittaLookupTimers.get(input);
        if (previousTimer) {{
          window.clearTimeout(previousTimer);
        }}
        const timer = window.setTimeout(() => {{
          lookupCittaDetails(input);
        }}, 180);
        cittaLookupTimers.set(input, timer);
      }}

      async function lookupCapDetails(input) {{
        const form = input.closest('form');
        const capField = associatoFormField(form, 'cap');
        if (!form || !capField) {{
          return null;
        }}
        const cap = (capField.value || '').replace(/\\D/g, '').slice(0, 5);
        capField.value = cap;
        if (cap.length !== 5) {{
          return null;
        }}
        try {{
          const payload = await fetchJson(`/api/cap?value=${{encodeURIComponent(cap)}}`);
          if (!payload || !payload.found) {{
            return null;
          }}
          assignFieldValue(form, 'citta', payload.citta || '');
          assignFieldValue(form, 'provincia', payload.provincia || '');
          return payload;
        }} catch (error) {{
          console.warn('Lookup CAP non disponibile', error);
          return null;
        }}
      }}

      function scheduleCapLookup(input) {{
        const previousTimer = capLookupTimers.get(input);
        if (previousTimer) {{
          window.clearTimeout(previousTimer);
        }}
        const timer = window.setTimeout(() => {{
          lookupCapDetails(input);
        }}, 180);
        capLookupTimers.set(input, timer);
      }}

      async function generateCodiceFiscaleFromFields(form) {{
        if (!form) {{
          return null;
        }}
        const nome = associatoFormField(form, 'nome');
        const cognome = associatoFormField(form, 'cognome');
        const dataNascita = associatoFormField(form, 'data_nascita');
        const sesso = associatoFormField(form, 'sesso');
        const comuneNascita = associatoFormField(form, 'comune_nascita');
        const codiceFiscale = associatoFormField(form, 'codice_fiscale');
        if (!nome || !cognome || !dataNascita || !sesso || !comuneNascita || !codiceFiscale) {{
          return null;
        }}

        if (!(nome.value || '').trim() || !(cognome.value || '').trim() || !(dataNascita.value || '').trim() || !(sesso.value || '').trim() || !(comuneNascita.value || '').trim()) {{
          return null;
        }}

        try {{
          const url = new URL('/api/codice-fiscale/calcola', window.location.origin);
          url.searchParams.set('nome', nome.value || '');
          url.searchParams.set('cognome', cognome.value || '');
          url.searchParams.set('data_nascita', dataNascita.value || '');
          url.searchParams.set('sesso', sesso.value || '');
          url.searchParams.set('comune_nascita', comuneNascita.value || '');
          const payload = await fetchJson(url.pathname + url.search);
          if (!payload || !payload.found) {{
            return null;
          }}
          assignFieldValue(form, 'codice_fiscale', payload.codice_fiscale || '');
          assignFieldValue(form, 'comune_nascita', payload.comune_nascita || comuneNascita.value || '');
          assignFieldValue(form, 'provincia_nascita', payload.provincia_nascita || '');
          return payload;
        }} catch (error) {{
          console.warn('Calcolo codice fiscale non disponibile', error);
          return null;
        }}
      }}

      function scheduleCodiceFiscaleGeneration(trigger) {{
        const form = trigger ? trigger.closest('form') : null;
        if (!form) {{
          return;
        }}
        const previousTimer = codiceFiscaleGenerationTimers.get(form);
        if (previousTimer) {{
          window.clearTimeout(previousTimer);
        }}
        const timer = window.setTimeout(() => {{
          generateCodiceFiscaleFromFields(form);
        }}, 180);
        codiceFiscaleGenerationTimers.set(form, timer);
      }}

      let whatsappShareWindow = null;
      let activeCourseReminder = null;

      function closeWhatsappShareWindow() {{
        if (!whatsappShareWindow) {{
          return;
        }}
        try {{
          if (!whatsappShareWindow.closed) {{
            whatsappShareWindow.close();
          }}
        }} catch (error) {{
          console.warn('Chiusura finestra WhatsApp non disponibile', error);
        }}
        whatsappShareWindow = null;
      }}

      function openCourseGenerationReminder() {{
        const reminder = document.querySelector('[data-course-generation-reminder="true"]');
        if (!reminder) {{
          return;
        }}
        const reminderKey = reminder.dataset.reminderKey || '';
        if (!reminderKey) {{
          return;
        }}
        const storageKey = `course-generation-reminder:${{reminderKey}}`;
        if (window.sessionStorage && window.sessionStorage.getItem(storageKey) === 'done') {{
          return;
        }}
        activeCourseReminder = {{ element: reminder, storageKey }};
        window.setTimeout(() => {{
          reminder.hidden = false;
          syncModalOpenState();
          const submitButton = reminder.querySelector('button[type=\"submit\"]');
          if (submitButton) {{
            submitButton.focus();
          }}
        }}, 180);
      }}

      function closeCourseGenerationReminder(markDone = true) {{
        const reminderElement = activeCourseReminder && activeCourseReminder.element
          ? activeCourseReminder.element
          : document.querySelector('[data-course-generation-reminder="true"]');
        if (!reminderElement) {{
          return;
        }}
        const reminderKey = activeCourseReminder && activeCourseReminder.storageKey
          ? activeCourseReminder.storageKey
          : (reminderElement.dataset.reminderKey ? `course-generation-reminder:${{reminderElement.dataset.reminderKey}}` : '');
        reminderElement.hidden = true;
        syncModalOpenState();
        if (markDone && window.sessionStorage && reminderKey) {{
          window.sessionStorage.setItem(reminderKey, 'done');
        }}
        activeCourseReminder = null;
      }}

      document.addEventListener('DOMContentLoaded', () => {{
        document.querySelectorAll('table[data-summary-columns]').forEach((table) => updateVisibleTableTotals(table));

        document.querySelectorAll('.select-search-control').forEach((input) => {{
          const {{ select }} = getAutocompleteElements(input);
          if (!select) {{
            return;
          }}
          syncAutocompleteInput(input);
          select.addEventListener('change', () => syncAutocompleteInput(input));
        }});

        document.querySelectorAll('form').forEach((form) => {{
          form.addEventListener('submit', (event) => {{
            for (const input of form.querySelectorAll('.select-search-control')) {{
              if (!commitSelectAutocomplete(input, input.required)) {{
                event.preventDefault();
                input.reportValidity();
                input.focus();
                return;
              }}
            }}
          }});
        }});

        document.querySelectorAll('form[data-confirm-dialog=\"true\"]').forEach((form) => {{
          form.addEventListener('submit', async (event) => {{
            event.preventDefault();
            syncDeleteFormView(form);
            const confirmed = await appConfirm(form.dataset.confirmMessage || 'Confermare l\\'eliminazione del record selezionato?', {{
              title: form.dataset.confirmTitle || 'Conferma eliminazione',
              confirmLabel: form.dataset.confirmButton || 'Elimina',
              cancelLabel: 'Annulla',
              variant: 'danger',
              badge: 'Eliminazione',
            }});
            if (confirmed) {{
              form.submit();
            }}
          }});
        }});

        document.querySelectorAll('select.control:not(.searchable-select-source)').forEach((select) => {{
          installSelectTypeahead(select);
        }});

        document.querySelectorAll('select[data-residual-target]').forEach((select) => {{
          select.addEventListener('change', () => syncResidualInput(select));
          syncResidualInput(select);
        }});

        document.querySelectorAll('select[data-standard-target]').forEach((select) => {{
          select.addEventListener('change', () => syncStandardQuota(select));
          syncStandardQuota(select);
        }});

        document.querySelectorAll('select[data-amount-target]').forEach((select) => {{
          select.addEventListener('change', () => syncAmountProposal(select));
          syncAmountProposal(select);
        }});

        document.querySelectorAll('input[data-year-end-target]').forEach((input) => {{
          input.addEventListener('input', () => syncYearEndDate(input));
          input.addEventListener('blur', () => syncYearEndDate(input));
          syncYearEndDate(input);
        }});

        document.querySelectorAll('[data-month-end-target]').forEach((control) => {{
          const eventName = control.tagName === 'SELECT' ? 'change' : 'input';
          control.addEventListener(eventName, () => syncMonthEndDate(control));
          if (eventName !== 'change') {{
            control.addEventListener('blur', () => syncMonthEndDate(control));
          }}
          syncMonthEndDate(control);
        }});

        document.querySelectorAll('input[data-codice-fiscale]').forEach((input) => {{
          input.addEventListener('input', () => scheduleCodiceFiscaleLookup(input));
          input.addEventListener('blur', () => lookupCodiceFiscale(input));
          if ((input.value || '').trim().length === 16) {{
            lookupCodiceFiscale(input);
          }}
        }});

        document.querySelectorAll('form').forEach((form) => {{
          ['nome', 'cognome', 'data_nascita'].forEach((fieldName) => {{
            const field = associatoFormField(form, fieldName);
            if (!field) {{
              return;
            }}
            field.addEventListener('input', () => scheduleCodiceFiscaleGeneration(field));
            field.addEventListener('blur', () => generateCodiceFiscaleFromFields(form));
          }});

          const sessoField = associatoFormField(form, 'sesso');
          if (sessoField) {{
            sessoField.addEventListener('change', () => generateCodiceFiscaleFromFields(form));
          }}

          const comuneField = associatoFormField(form, 'comune_nascita');
          if (comuneField) {{
            comuneField.addEventListener('input', () => scheduleComuneNascitaLookup(comuneField));
            comuneField.addEventListener('blur', async () => {{
              await lookupComuneNascita(comuneField);
              await generateCodiceFiscaleFromFields(form);
            }});
          }}

          const capField = associatoFormField(form, 'cap');
          if (capField) {{
            capField.addEventListener('input', () => scheduleCapLookup(capField));
            capField.addEventListener('blur', () => lookupCapDetails(capField));
          }}

          const cittaField = associatoFormField(form, 'citta');
          if (cittaField) {{
            cittaField.addEventListener('input', () => scheduleCittaLookup(cittaField));
            cittaField.addEventListener('blur', () => lookupCittaDetails(cittaField));
          }}
        }});

        const associatoSelect = document.getElementById('monthly-payment-associato');
        const rateSelect = document.getElementById('monthly-rate-select');
        if (associatoSelect && rateSelect) {{
          associatoSelect.addEventListener('change', updateMonthlyRates);
          rateSelect.addEventListener('change', syncMonthlySelectedAmount);
          updateMonthlyRates();
        }}

        const multiAssociatoSelect = document.getElementById('multi-area-associato');
        const multiScadenzeSelect = document.getElementById('multi-area-scadenze');
        if (multiAssociatoSelect && multiScadenzeSelect) {{
          multiAssociatoSelect.addEventListener('change', () => {{
            syncMultiAreaFutureCourseControls();
            updateMultiAreaScadenze([]);
          }});
          multiScadenzeSelect.addEventListener('change', syncMultiAreaSelectedAmount);
          updateMultiAreaScadenze([]);
        }}

        const multiCourseSelect = document.getElementById('multi-area-course-enrollment');
        const multiCourseUntil = document.getElementById('multi-area-course-until');
        const multiCourseButton = document.getElementById('multi-area-course-generate');
        if (multiCourseSelect && multiCourseUntil) {{
          multiCourseSelect.addEventListener('change', syncMultiAreaFutureCourseControls);
          multiCourseUntil.addEventListener('change', syncMultiAreaFutureCourseControls);
          syncMultiAreaFutureCourseControls();
        }}
        if (multiCourseButton) {{
          multiCourseButton.addEventListener('click', generateFutureCourseMultiAreaScadenze);
        }}

        bindEnrollmentPaymentFlows();

        window.addEventListener('focus', () => {{
          closeWhatsappShareWindow();
        }});

        document.querySelectorAll('[data-reminder-dismiss=\"true\"]').forEach((element) => {{
          element.addEventListener('click', () => closeCourseGenerationReminder(true));
        }});

        document.addEventListener('keydown', (event) => {{
          if (event.key === 'Escape' && activeAppDialog && activeAppDialog.dismissible) {{
            closeAppDialog(false);
            return;
          }}
          if (event.key === 'Escape' && activeCourseReminder) {{
            closeCourseGenerationReminder(true);
          }}
        }});

        const reminderForm = document.getElementById('course-generation-reminder-form');
        if (reminderForm) {{
          reminderForm.addEventListener('submit', () => {{
            if (activeCourseReminder && window.sessionStorage && activeCourseReminder.storageKey) {{
              window.sessionStorage.setItem(activeCourseReminder.storageKey, 'done');
            }}
          }});
        }}

        openCourseGenerationReminder();
      }});

      function shareReceipt(button) {{
        const channel = button.dataset.channel || '';
        if (channel === 'email' && button.dataset.mailto) {{
          window.location.href = button.dataset.mailto;
          return false;
        }}
        if (channel === 'whatsapp' && button.dataset.whatsapp) {{
          closeWhatsappShareWindow();
          whatsappShareWindow = window.open(button.dataset.whatsapp, 'oratorioCarloAcutisWhatsappShare');
          if (whatsappShareWindow) {{
            try {{
              whatsappShareWindow.focus();
            }} catch (error) {{
              console.warn('Focus finestra WhatsApp non disponibile', error);
            }}
          }}
          return false;
        }}
        return false;
      }}

      async function shareReport(button) {{
        const channel = button.dataset.channel || '';
        const reportKey = button.dataset.reportKey || '';
        const recipientSelectId = button.dataset.recipientSelectId || '';
        const recipientSelect = recipientSelectId ? document.getElementById(recipientSelectId) : null;
        const opensWhatsappWindow = channel === 'whatsapp' || channel === 'whatsapp-group';
        const needsRecipient = !!recipientSelectId && (channel === 'email' || channel === 'whatsapp');
        if (!channel || !reportKey) {{
          await appAlert('Invio report non disponibile.', {{
            title: 'Invio report',
            variant: 'danger',
            badge: 'Errore',
          }});
          return false;
        }}
        if (needsRecipient && (!recipientSelect || !recipientSelect.value)) {{
          await appAlert('Seleziona prima un destinatario valido.', {{
            title: 'Invio report',
            variant: 'warning',
            badge: 'Controllo dati',
          }});
          return false;
        }}

        let pendingWhatsappWindow = null;
        let popupBlocked = false;
        if (opensWhatsappWindow) {{
          closeWhatsappShareWindow();
          pendingWhatsappWindow = window.open('', 'oratorioCarloAcutisWhatsappShare');
          if (pendingWhatsappWindow) {{
            whatsappShareWindow = pendingWhatsappWindow;
            try {{
              pendingWhatsappWindow.document.write('<!doctype html><title>Preparazione WhatsApp</title><body style="font-family:Arial,sans-serif;padding:24px;">Preparazione messaggio WhatsApp...</body>');
            }} catch (error) {{
              console.warn('Finestra temporanea WhatsApp non disponibile', error);
            }}
          }}
        }}
        if (opensWhatsappWindow && !pendingWhatsappWindow) {{
          popupBlocked = true;
        }}

        const url = new URL('/api/report-share', window.location.origin);
        const params = new URLSearchParams(window.location.search);
        const toolbar = button.closest('.report-toolbar');
        const searchInput = toolbar ? toolbar.querySelector('.report-search input') : null;
        if (searchInput) {{
          if (searchInput.value) {{
            params.set('search', searchInput.value);
          }} else {{
            params.delete('search');
          }}
        }}
        params.set('report_key', reportKey);
        if (recipientSelect && recipientSelect.value) {{
          params.set('recipient_id', recipientSelect.value);
        }} else {{
          params.delete('recipient_id');
        }}
        params.set('channel', channel);
        url.search = params.toString();

        try {{
          const response = await fetch(url.toString(), {{
            headers: {{ 'Accept': 'application/json' }}
          }});
          const payload = await response.json();
          if (!response.ok || !payload || !payload.ok || !payload.url) {{
            if (pendingWhatsappWindow) {{
              try {{
                pendingWhatsappWindow.close();
              }} catch (closeError) {{
                console.warn('Chiusura finestra WhatsApp non disponibile', closeError);
              }}
            if (whatsappShareWindow === pendingWhatsappWindow) {{
                whatsappShareWindow = null;
              }}
            }}
            await appAlert((payload && payload.error) || 'Invio report non disponibile.', {{
              title: 'Invio report',
              variant: 'danger',
              badge: 'Errore',
            }});
            return false;
          }}

          if (payload.channel === 'email') {{
            window.location.href = payload.url;
            return false;
          }}

          if (payload.channel === 'whatsapp-group') {{
            if (payload.message) {{
              try {{
                if (navigator.clipboard && navigator.clipboard.writeText) {{
                  await navigator.clipboard.writeText(payload.message);
                }}
              }} catch (error) {{
                console.warn('Copia del messaggio gruppo WhatsApp non disponibile', error);
              }}
            }}
            whatsappShareWindow = pendingWhatsappWindow || whatsappShareWindow;
            if (whatsappShareWindow) {{
              whatsappShareWindow.location.href = payload.url;
            }} else if (popupBlocked) {{
              window.location.href = payload.url;
            }} else {{
              whatsappShareWindow = window.open(payload.url, 'oratorioCarloAcutisWhatsappShare');
            }}
            if (whatsappShareWindow) {{
              try {{
                whatsappShareWindow.focus();
              }} catch (error) {{
                console.warn('Focus finestra WhatsApp non disponibile', error);
              }}
            }}
            return false;
          }}

          whatsappShareWindow = pendingWhatsappWindow || whatsappShareWindow;
          if (whatsappShareWindow) {{
            whatsappShareWindow.location.href = payload.url;
          }} else if (popupBlocked) {{
            window.location.href = payload.url;
          }} else {{
            closeWhatsappShareWindow();
            whatsappShareWindow = window.open(payload.url, 'oratorioCarloAcutisWhatsappShare');
          }}
          if (whatsappShareWindow) {{
            try {{
              whatsappShareWindow.focus();
            }} catch (error) {{
              console.warn('Focus finestra WhatsApp non disponibile', error);
            }}
          }}
          return false;
        }} catch (error) {{
          if (pendingWhatsappWindow) {{
            try {{
              pendingWhatsappWindow.close();
            }} catch (closeError) {{
              console.warn('Chiusura finestra WhatsApp non disponibile', closeError);
            }}
            if (whatsappShareWindow === pendingWhatsappWindow) {{
              whatsappShareWindow = null;
            }}
          }}
          console.warn('Invio report non disponibile', error);
          await appAlert('Invio report non disponibile.', {{
            title: 'Invio report',
            variant: 'danger',
            badge: 'Errore',
          }});
          return false;
        }}
      }}
    </script>
  </body>
</html>"""
    return document.encode("utf-8")


def redirect(
    start_response,
    path: str,
    *,
    ok: str | None = None,
    err: str | None = None,
    extra_query: dict[str, str] | None = None,
    extra_headers: list[tuple[str, str]] | None = None,
):
    params = dict(extra_query or {})
    if ok:
        params["ok"] = ok
    if err:
        params["err"] = err
    parsed = urlsplit(path)
    existing_query = {
        key: ",".join(values)
        for key, values in parse_qs(parsed.query, keep_blank_values=True).items()
    }
    existing_query.update(params)
    query_string = urlencode(existing_query)
    location = urlunsplit(("", "", parsed.path, query_string, parsed.fragment))
    headers = [("Location", location)]
    if extra_headers:
        headers.extend(extra_headers)
    start_response("303 See Other", headers)
    return [b""]


def friendly_db_error(error: Exception) -> str:
    message = str(error)
    replacements = {
        "UNIQUE constraint failed: associati.codice_associato": "Esiste gia un associato con questo codice.",
        "UNIQUE constraint failed: associati.codice_fiscale": "Esiste gia un associato con questo codice fiscale.",
        "UNIQUE constraint failed: tesseramenti_annuali.associato_id, tesseramenti_annuali.anno_sociale": "Questo associato ha gia un tesseramento per l'anno indicato.",
        "UNIQUE constraint failed: pagamenti_campi_estivi.iscrizione_campo_id": "Per questa iscrizione al Campo estivo esiste gia un pagamento una tantum.",
        "UNIQUE constraint failed: pagamenti_eventi.iscrizione_evento_id": "Per questa iscrizione evento esiste gia un pagamento una tantum.",
        "UNIQUE constraint failed: iscrizioni_campi_estivi.associato_id, iscrizioni_campi_estivi.campo_estivo_id": "Questo associato risulta gia iscritto al Campo estivo dell'anno selezionato.",
        "UNIQUE constraint failed: iscrizioni_eventi.associato_id, iscrizioni_eventi.evento_id": "Questo associato risulta gia iscritto all'evento selezionato.",
        "UNIQUE constraint failed: rate_corsi_mensili.iscrizione_corso_id, rate_corsi_mensili.anno, rate_corsi_mensili.mese": "Per questa iscrizione corso esiste gia una quota mensile per anno e mese indicati.",
        "UNIQUE constraint failed: corsi.codice_corso": "Esiste gia un corso con questo codice.",
        "UNIQUE constraint failed: eventi.codice_evento": "Esiste gia un evento con questo codice.",
        "UNIQUE constraint failed: campi_estivi.codice_campo": "Esiste gia un Campo estivo con questo codice.",
        "UNIQUE constraint failed: tipologie_corsi.codice_tipologia": "Esiste gia una tipologia corso con questo codice.",
        "UNIQUE constraint failed: tipologie_corsi.nome": "Esiste gia una tipologia corso con questo nome.",
        "UNIQUE constraint failed: utenti_accesso.username": "Esiste gia un utente con questo username.",
        "FOREIGN KEY constraint failed": "Operazione non possibile: esistono dati collegati che devono essere gestiti prima.",
    }
    for needle, replacement in replacements.items():
        if needle in message:
            return replacement
    return message


def associati_options() -> list[sqlite3.Row]:
    rows = fetch_all(
        """
        SELECT
            id,
            codice_associato,
            nome,
            cognome,
            data_nascita,
            COALESCE(telefono, '') AS telefono,
            COALESCE(email, '') AS email
        FROM associati
        ORDER BY cognome, nome
        """
    )
    options = []
    for row in rows:
        full_name = plain_text(f"{row['cognome']} {row['nome']}")
        named_label = label_with_age(full_name, row["data_nascita"])
        contact_parts = []
        if row["telefono"]:
            contact_parts.append(row["telefono"])
        if row["email"]:
            contact_parts.append(row["email"])
        contact_tail = f" - {' - '.join(contact_parts)}" if contact_parts else ""
        options.append(
            {
                "id": row["id"],
                "label": f"{row['codice_associato']} - {named_label}",
                "autocomplete_label": f"{row['codice_associato']} - {named_label}{contact_tail}",
                "search_text": plain_text(
                    f"{row['codice_associato']} {row['cognome']} {row['nome']} {row['telefono']} {row['email']}"
                ).lower(),
            }
        )
    return options


def report_recipient_rows() -> list[dict[str, str]]:
    placeholders = ", ".join("?" for _ in REPORT_RECIPIENT_CARICHE)
    rows = fetch_all(
        f"""
        SELECT
            id,
            nome,
            cognome,
            data_nascita,
            carica,
            COALESCE(email, '') AS email,
            COALESCE(telefono, '') AS telefono
        FROM associati
        WHERE carica IN ({placeholders})
        ORDER BY cognome, nome
        """,
        tuple(REPORT_RECIPIENT_CARICHE),
    )
    role_order = {value: index for index, value in enumerate(REPORT_RECIPIENT_CARICHE)}
    recipients: list[dict[str, str]] = []
    for row in rows:
        full_name = plain_text(f"{row['cognome']} {row['nome']}")
        named_label = label_with_age(full_name, row["data_nascita"])
        contact_parts = []
        if row["email"]:
            contact_parts.append(row["email"])
        if row["telefono"]:
            contact_parts.append(row["telefono"])
        contact_tail = f" - {' - '.join(contact_parts)}" if contact_parts else ""
        recipients.append(
            {
                "id": str(row["id"]),
                "label": f"{row['carica']} - {named_label}",
                "autocomplete_label": f"{row['carica']} - {named_label}{contact_tail}",
                "search_text": plain_text(
                    f"{row['carica']} {row['cognome']} {row['nome']} {row['email']} {row['telefono']}"
                ).lower(),
                "email": row["email"],
                "whatsapp_phone": clean_phone_number(row["telefono"]),
                "carica": row["carica"],
            }
        )
    recipients.sort(key=lambda item: (role_order.get(item["carica"], 999), item["label"]))
    return recipients


def report_recipient_options(selected: str | None = None) -> str:
    return render_select_options(
        report_recipient_rows(),
        selected,
        blank_label="Seleziona destinatario...",
        data_keys=["search_text", "autocomplete_label", "email", "whatsapp_phone", "carica"],
    )


def activity_log_user_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT DISTINCT
            username AS id,
            username AS label
        FROM registro_attivita
        WHERE COALESCE(username, '') <> ''
        ORDER BY username
        """
    )


def activity_log_associato_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT DISTINCT
            associato_id AS id,
            trim(COALESCE(associato_codice, '') || CASE
                WHEN COALESCE(associato_codice, '') <> '' AND COALESCE(associato_nominativo, '') <> '' THEN ' - '
                ELSE ''
            END || COALESCE(associato_nominativo, '')) AS label,
            lower(trim(COALESCE(associato_codice, '') || ' ' || COALESCE(associato_nominativo, ''))) AS search_text,
            trim(COALESCE(associato_codice, '') || CASE
                WHEN COALESCE(associato_codice, '') <> '' AND COALESCE(associato_nominativo, '') <> '' THEN ' - '
                ELSE ''
            END || COALESCE(associato_nominativo, '')) AS autocomplete_label
        FROM registro_attivita
        WHERE associato_id IS NOT NULL
          AND COALESCE(associato_nominativo, '') <> ''
        ORDER BY label
        """
    )


def activity_log_attivita_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT DISTINCT
            descrizione_attivita AS id,
            descrizione_attivita AS label
        FROM registro_attivita
        WHERE COALESCE(descrizione_attivita, '') <> ''
        ORDER BY descrizione_attivita
        """
    )


def can_manage_cariche(current_user: dict[str, object] | None) -> bool:
    return bool(current_user and current_user.get("is_admin"))


def available_cariche(current_user: dict[str, object] | None) -> tuple[str, ...]:
    return CARICA_VALUES if can_manage_cariche(current_user) else ("Associato",)


def resolved_carica_value(
    form_data: dict[str, str],
    current_user: dict[str, object] | None,
    *,
    existing_value: str = "Associato",
) -> str:
    if not can_manage_cariche(current_user):
        return existing_value if existing_value in CARICA_VALUES else "Associato"
    selected = normalized(form_data, "carica", existing_value) or "Associato"
    if selected not in CARICA_VALUES:
        raise ValueError("La carica selezionata non e valida.")
    return selected


def associato_carica_field(selected: str, current_user: dict[str, object] | None) -> str:
    current_value = selected if selected in CARICA_VALUES else "Associato"
    if can_manage_cariche(current_user):
        return select_field("Carica", "carica", carica_options(current_value, current_user))
    return readonly_field("Carica", current_value)


def metodi_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT id, nome AS label
        FROM metodi_pagamento
        WHERE attivo = 1
        ORDER BY nome
        """
    )


def preferred_metodo_pagamento_id(metodi: list[sqlite3.Row]) -> str:
    for row in metodi:
        if plain_text(row["label"]).lower() == "contanti":
            return str(row["id"])
    return str(metodi[0]["id"]) if metodi else ""


def quote_predefinite_options(area: str) -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            id,
            replace(printf('%.2f', importo), ',', '.') AS importo,
            descrizione || ' - ' || replace(printf('%.2f', importo), '.', ',') || ' EUR' AS label
        FROM quote_predefinite
        WHERE area = ? AND attiva = 1
        ORDER BY descrizione, id
        """,
        (area,),
    )


def quote_predefinite_rows(area: str) -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT id, descrizione, importo, attiva, COALESCE(note, '') AS note
        FROM quote_predefinite
        WHERE area = ?
        ORDER BY descrizione, id
        """,
        (area,),
    )


def tipologie_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT id, codice_tipologia || ' - ' || nome AS label
        FROM tipologie_corsi
        WHERE attiva = 1
        ORDER BY nome
        """
    )


def corsi_options() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            c.id,
            c.nome AS label,
            replace(printf('%.2f', c.quota_mensile_standard), ',', '.') AS importo
        FROM corsi c
        WHERE c.attivo = 1
        ORDER BY c.nome
        """
    )


def pagamenti_multi_area_course_options(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            ic.id,
            ic.associato_id,
            substr(COALESCE(NULLIF(ic.data_inizio, ''), ic.data_iscrizione, ''), 1, 7) AS start_competenza,
            c.nome || ' - decorrenza ' || substr(COALESCE(NULLIF(ic.data_inizio, ''), ic.data_iscrizione, ''), 1, 7) AS label
        FROM iscrizioni_corsi ic
        JOIN corsi c ON c.id = ic.corso_id
        WHERE ic.stato_iscrizione = 'Attiva'
          AND {iscrizione_corso_year_relevance_sql('ic')}
        ORDER BY c.nome, start_competenza, ic.id
        """,
        iscrizione_corso_year_relevance_params(work_year),
    )


def corsi_report_options(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            c.id,
            c.nome AS label
        FROM corsi c
        WHERE c.attivo = 1
          AND {corso_year_relevance_sql('c')}
        ORDER BY c.nome
        """,
        corso_year_relevance_params(work_year),
    )


def iscrizioni_corsi_options() -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            ic.id,
            {associato_display_sql('a')} || ' - ' || c.nome AS label
        FROM iscrizioni_corsi ic
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        ORDER BY ic.id DESC
        """
    )


def rate_corsi_options() -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            r.id,
            {associato_display_sql('a')} || ' - ' || c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS label
        FROM rate_corsi_mensili r
        JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        ORDER BY r.anno DESC, r.mese DESC, a.cognome, a.nome
        """
    )


def rate_corsi_aperte_options(work_year: int, associato_id: str | None = None) -> list[sqlite3.Row]:
    params: list[object] = [work_year]
    where_clause = "WHERE v.anno = ? AND v.saldo_residuo > 0"
    if associato_id and associato_id.isdigit():
        where_clause += " AND v.associato_id = ?"
        params.append(int(associato_id))
    return fetch_all(
        f"""
        SELECT
            v.id,
            v.associato_id,
            replace(printf('%.2f', v.saldo_residuo), ',', '.') AS residuo,
            v.associato || ' - ' || v.corso || ' - ' || v.competenza || ' - residuo ' ||
            replace(printf('%.2f', v.saldo_residuo), '.', ',') || ' EUR' AS label
        FROM v_rate_corsi_saldo v
        {where_clause}
        ORDER BY v.associato, v.corso, v.anno, v.mese
        """,
        tuple(params),
    )


def tesseramenti_options() -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            t.id,
            {associato_display_sql('a')} || ' - anno ' || t.anno_sociale AS label
        FROM tesseramenti_annuali t
        JOIN associati a ON a.id = t.associato_id
        ORDER BY t.anno_sociale DESC, a.cognome, a.nome
        """
    )


def tesseramenti_aperti_options(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            v.id,
            replace(printf('%.2f', v.saldo_residuo), ',', '.') AS residuo,
            v.associato || ' - anno ' || v.anno_sociale || ' - residuo ' ||
            replace(printf('%.2f', v.saldo_residuo), '.', ',') || ' EUR' AS label
        FROM v_tesseramenti_saldo v
        WHERE v.anno_sociale = ? AND v.saldo_residuo > 0
        ORDER BY v.associato
        """,
        (work_year,),
    )


def campi_estivi_options() -> list[sqlite3.Row]:
    return campi_estivi_options_for_year(None)


def campi_estivi_options_for_year(work_year: int | None) -> list[sqlite3.Row]:
    params: tuple[object, ...] = ()
    where_clause = "WHERE attivo = 1"
    if work_year is not None:
        where_clause += " AND anno = ?"
        params = (work_year,)
    return fetch_all(
        f"""
        SELECT
            id,
            anno,
            replace(printf('%.2f', quota_partecipazione_standard), ',', '.') AS quota_standard,
            'Campo estivo ' || anno AS label
        FROM campi_estivi
        {where_clause}
        ORDER BY anno DESC, id DESC
        """,
        params,
    )


def estate_record_for_year(work_year: int) -> sqlite3.Row | None:
    return fetch_one(
        """
        SELECT id, anno, quota_partecipazione_standard, attivo
        FROM campi_estivi
        WHERE anno = ?
        ORDER BY id
        LIMIT 1
        """,
        (work_year,),
    )


def ensure_estate_record(connection: sqlite3.Connection, work_year: int, standard_fee: str | None = None) -> int:
    existing_id = connection.execute(
        "SELECT id FROM campi_estivi WHERE anno = ? ORDER BY id LIMIT 1",
        (work_year,),
    ).fetchone()
    if existing_id:
        if standard_fee is not None:
            connection.execute(
                "UPDATE campi_estivi SET quota_partecipazione_standard = ?, attivo = 1 WHERE id = ?",
                (standard_fee, existing_id["id"]),
            )
        return int(existing_id["id"])

    progressive_number = reserve_progressive_number(connection, "campi_estivi")
    connection.execute(
        """
        INSERT INTO campi_estivi (
            numero_progressivo, codice_campo, nome, anno, data_inizio, data_fine, quota_partecipazione_standard, attivo
        ) VALUES (?, ?, ?, ?, ?, ?, ?, 1)
        """,
        (
            progressive_number,
            format_progressive_code("campi_estivi", progressive_number),
            f"{ESTATE_LABEL} {work_year}",
            work_year,
            f"{work_year}-06-01",
            f"{work_year}-06-30",
            standard_fee or "0",
        ),
    )
    return int(connection.execute("SELECT last_insert_rowid()").fetchone()[0])


def iscrizioni_campi_options() -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            ice.id,
            {associato_display_sql('a')} || ' - ' || ESTATE_LABEL || ' ' || ce.anno AS label
        FROM iscrizioni_campi_estivi ice
        JOIN associati a ON a.id = ice.associato_id
        JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
        WHERE ce.anno = ?
        ORDER BY ice.id DESC
        """,
        (work_year,),
    )


def iscrizioni_campi_aperte_options(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            v.id,
            replace(printf('%.2f', v.saldo_residuo), ',', '.') AS residuo,
            v.associato || ' - anno ' || v.anno || ' - residuo ' ||
            replace(printf('%.2f', v.saldo_residuo), '.', ',') || ' EUR' AS label
        FROM v_campi_estivi_saldo v
        WHERE v.anno = ? AND v.saldo_residuo > 0
        ORDER BY v.associato
        """,
        (work_year,),
    )


def eventi_options(work_year: int | None = None) -> list[sqlite3.Row]:
    params: tuple[object, ...] = ()
    where_clause = "WHERE attivo = 1"
    if work_year is not None:
        where_clause += " AND substr(COALESCE(data_evento, ''), 1, 4) = ?"
        params = (str(work_year),)
    return fetch_all(
        f"""
        SELECT
            id,
            codice_evento || ' - ' || nome AS label,
            replace(printf('%.2f', quota_partecipazione_standard), ',', '.') AS importo
        FROM eventi
        {where_clause}
        ORDER BY data_evento DESC, nome
        """,
        params,
    )


def iscrizioni_eventi_options() -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        SELECT
            ie.id,
            {associato_display_sql('a')} || ' - ' || e.nome AS label
        FROM iscrizioni_eventi ie
        JOIN associati a ON a.id = ie.associato_id
        JOIN eventi e ON e.id = ie.evento_id
        ORDER BY ie.id DESC
        """
    )


def iscrizioni_eventi_aperte_options(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            v.id,
            replace(printf('%.2f', v.saldo_residuo), ',', '.') AS residuo,
            v.associato || ' - ' || v.evento || ' - residuo ' ||
            replace(printf('%.2f', v.saldo_residuo), '.', ',') || ' EUR' AS label
        FROM v_eventi_saldo v
        WHERE substr(v.data_evento, 1, 4) = ? AND v.saldo_residuo > 0
        ORDER BY v.evento, v.associato
        """,
        (str(work_year),),
    )


def scadenze_multi_area_options(work_year: int, associato_id: int | None = None) -> list[sqlite3.Row]:
    clauses = ["1 = 1"]
    params: list[object] = [work_year, work_year, work_year, str(work_year)]
    if associato_id is not None:
        clauses.append("associato_id = ?")
        params.append(int(associato_id))
    return fetch_all(
        f"""
        SELECT
            kind || ':' || source_id AS id,
            associato_id,
            area,
            riferimento,
            scadenza,
            replace(printf('%.2f', saldo_residuo), ',', '.') AS residuo,
            area || ' - ' || riferimento || ' - ' || COALESCE(scadenza, '') || ' - residuo ' ||
            replace(printf('%.2f', saldo_residuo), '.', ',') || ' EUR' AS label
        FROM (
            SELECT
                'tesseramenti' AS kind,
                id AS source_id,
                associato_id,
                'Tesseramento annuale' AS area,
                'Anno ' || anno_sociale AS riferimento,
                COALESCE(data_scadenza, data_tesseramento) AS scadenza,
                saldo_residuo
            FROM v_tesseramenti_saldo
            WHERE anno_sociale = ? AND saldo_residuo > 0

            UNION ALL

            SELECT
                'corsi-rate' AS kind,
                id AS source_id,
                associato_id,
                'Corso - quota mensile' AS area,
                corso || ' ' || competenza AS riferimento,
                COALESCE(data_scadenza, printf('%04d-%02d-01', anno, mese)) AS scadenza,
                saldo_residuo
            FROM v_rate_corsi_saldo
            WHERE anno = ? AND saldo_residuo > 0

            UNION ALL

            SELECT
                'campi-estivi' AS kind,
                id AS source_id,
                associato_id,
                'Campo estivo' AS area,
                campo_estivo AS riferimento,
                COALESCE(data_inizio, data_iscrizione) AS scadenza,
                saldo_residuo
            FROM v_campi_estivi_saldo
            WHERE anno = ? AND saldo_residuo > 0

            UNION ALL

            SELECT
                'eventi' AS kind,
                id AS source_id,
                associato_id,
                'Evento' AS area,
                evento AS riferimento,
                data_evento AS scadenza,
                saldo_residuo
            FROM v_eventi_saldo
            WHERE substr(data_evento, 1, 4) = ? AND saldo_residuo > 0
        ) scadenze
        WHERE {' AND '.join(clauses)}
        ORDER BY scadenza, area, riferimento
        """,
        tuple(params),
    )


def scadenze_multi_area_payload_rows(work_year: int, associato_id: int | None = None) -> list[dict[str, str]]:
    return [
        {
            "id": str(row["id"]),
            "associato_id": str(row["associato_id"]),
            "label": plain_text(row["label"]),
            "residuo": str(row["residuo"]),
        }
        for row in scadenze_multi_area_options(work_year, associato_id)
    ]


def generate_multi_area_future_course_payload(
    work_year: int,
    associato_id: int,
    iscrizione_corso_id: int,
    fine_competenza: str,
) -> dict[str, object]:
    end_year, end_month = parse_year_month_value(fine_competenza, "Competenza finale")
    if end_year != work_year:
        raise ValueError("Puoi generare quote future del corso solo entro l'anno di lavoro selezionato.")

    iscrizione = fetch_one(
        """
        SELECT
            ic.id,
            ic.associato_id,
            c.nome AS corso,
            substr(COALESCE(NULLIF(ic.data_inizio, ''), ic.data_iscrizione, ''), 1, 7) AS start_competenza
        FROM iscrizioni_corsi ic
        JOIN corsi c ON c.id = ic.corso_id
        WHERE ic.id = ?
          AND ic.associato_id = ?
          AND ic.stato_iscrizione = 'Attiva'
        """,
        (iscrizione_corso_id, associato_id),
    )
    if iscrizione is None:
        raise ValueError("Iscrizione corso non valida per l'associato selezionato.")

    start_year, start_month = parse_year_month_value(str(iscrizione["start_competenza"]), "Decorrenza corso")
    effective_start_year, effective_start_month = (start_year, start_month)
    if (effective_start_year, effective_start_month) < (work_year, 1):
        effective_start_year, effective_start_month = work_year, 1
    if (end_year, end_month) < (effective_start_year, effective_start_month):
        raise ValueError("La competenza finale deve essere uguale o successiva alla decorrenza del corso.")

    with get_connection() as connection:
        ensure_course_rates_for_enrollment_range(
            connection,
            iscrizione_corso_id,
            effective_start_year,
            effective_start_month,
            end_year,
            end_month,
            note="Quota generata automaticamente da Pagamenti multi-area",
        )
        selected_rows = connection.execute(
            """
            SELECT
                'corsi-rate:' || r.id AS token
            FROM rate_corsi_mensili r
            LEFT JOIN (
                SELECT rata_corso_id, SUM(importo) AS totale_pagato
                FROM pagamenti_rate_corsi
                GROUP BY rata_corso_id
            ) pagamenti ON pagamenti.rata_corso_id = r.id
            WHERE r.iscrizione_corso_id = ?
              AND r.anno = ?
              AND (r.anno < ? OR (r.anno = ? AND r.mese <= ?))
              AND (r.importo_dovuto - COALESCE(pagamenti.totale_pagato, 0)) > 0
            ORDER BY r.anno, r.mese, r.id
            """,
            (iscrizione_corso_id, work_year, end_year, end_year, end_month),
        ).fetchall()
        connection.commit()

    selected_tokens = [str(row["token"]) for row in selected_rows]
    target_label = f"{month_label(end_month)} {end_year}"
    return {
        "ok": True,
        "message": (
            f"Quote del corso {plain_text(iscrizione['corso'])} aggiornate fino a {target_label}. "
            "Le mensilita aperte del corso sono state selezionate automaticamente."
        ),
        "selected_scadenze": selected_tokens,
        "options": scadenze_multi_area_payload_rows(work_year, associato_id),
    }


def parse_scadenza_multi_area_token(token: str) -> tuple[str, int]:
    if ":" not in token:
        raise ValueError("Selezione scadenza non valida.")
    kind, raw_id = token.split(":", 1)
    if kind not in {"tesseramenti", "corsi-rate", "campi-estivi", "eventi"} or not raw_id.isdigit():
        raise ValueError("Selezione scadenza non valida.")
    return kind, int(raw_id)


def load_multi_area_scadenze(tokens: list[str]) -> list[dict]:
    parsed = [parse_scadenza_multi_area_token(token) for token in tokens]
    grouped_ids: dict[str, list[int]] = {}
    for kind, source_id in parsed:
        grouped_ids.setdefault(kind, []).append(source_id)

    resolved: dict[str, dict] = {}

    def store_rows(rows: list[sqlite3.Row], kind: str) -> None:
        for row in rows:
            token = f"{kind}:{row['source_id']}"
            resolved[token] = dict(row)

    for kind, ids in grouped_ids.items():
        placeholders = ",".join("?" for _ in ids)
        if kind == "tesseramenti":
            rows = fetch_all(
                f"""
                SELECT
                    id AS source_id,
                    associato_id,
                    'Tesseramento annuale' AS area,
                    'Anno ' || anno_sociale AS riferimento,
                    COALESCE(data_scadenza, data_tesseramento) AS scadenza,
                    importo_dovuto,
                    importo_pagato,
                    saldo_residuo
                FROM v_tesseramenti_saldo
                WHERE id IN ({placeholders})
                """,
                tuple(ids),
            )
        elif kind == "corsi-rate":
            rows = fetch_all(
                f"""
                SELECT
                    id AS source_id,
                    associato_id,
                    'Corso - quota mensile' AS area,
                    corso || ' ' || competenza AS riferimento,
                    COALESCE(data_scadenza, printf('%04d-%02d-01', anno, mese)) AS scadenza,
                    importo_dovuto,
                    importo_pagato,
                    saldo_residuo
                FROM v_rate_corsi_saldo
                WHERE id IN ({placeholders})
                """,
                tuple(ids),
            )
        elif kind == "campi-estivi":
            rows = fetch_all(
                f"""
                SELECT
                    id AS source_id,
                    associato_id,
                    'Campo estivo' AS area,
                    campo_estivo AS riferimento,
                    COALESCE(data_inizio, data_iscrizione) AS scadenza,
                    importo_dovuto,
                    importo_pagato,
                    saldo_residuo
                FROM v_campi_estivi_saldo
                WHERE id IN ({placeholders})
                """,
                tuple(ids),
            )
        else:
            rows = fetch_all(
                f"""
                SELECT
                    id AS source_id,
                    associato_id,
                    'Evento' AS area,
                    evento AS riferimento,
                    data_evento AS scadenza,
                    importo_dovuto,
                    importo_pagato,
                    saldo_residuo
                FROM v_eventi_saldo
                WHERE id IN ({placeholders})
                """,
                tuple(ids),
            )
        store_rows(rows, kind)

    ordered_rows: list[dict] = []
    for token in tokens:
        row = resolved.get(token)
        if row is None:
            raise ValueError("Una o piu scadenze selezionate non sono piu disponibili.")
        row["kind"], row["source_id"] = parse_scadenza_multi_area_token(token)
        ordered_rows.append(row)
    return ordered_rows


def multi_area_receipts_rows(work_year: int) -> list[sqlite3.Row]:
    return fetch_all(
        f"""
        WITH pagamenti AS (
            SELECT
                pt.gruppo_ricevuta,
                pt.data_pagamento,
                pt.importo,
                {associato_display_sql('a')} AS associato
            FROM pagamenti_tesseramenti pt
            JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
            JOIN associati a ON a.id = t.associato_id
            WHERE pt.gruppo_ricevuta LIKE 'MGR-%' AND t.anno_sociale = ?

            UNION ALL

            SELECT
                prc.gruppo_ricevuta,
                prc.data_pagamento,
                prc.importo,
                {associato_display_sql('a')} AS associato
            FROM pagamenti_rate_corsi prc
            JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
            JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
            JOIN associati a ON a.id = ic.associato_id
            WHERE prc.gruppo_ricevuta LIKE 'MGR-%' AND r.anno = ?

            UNION ALL

            SELECT
                pce.gruppo_ricevuta,
                pce.data_pagamento,
                pce.importo,
                {associato_display_sql('a')} AS associato
            FROM pagamenti_campi_estivi pce
            JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
            JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
            JOIN associati a ON a.id = ice.associato_id
            WHERE pce.gruppo_ricevuta LIKE 'MGR-%' AND ce.anno = ?

            UNION ALL

            SELECT
                pe.gruppo_ricevuta,
                pe.data_pagamento,
                pe.importo,
                {associato_display_sql('a')} AS associato
            FROM pagamenti_eventi pe
            JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
            JOIN eventi e ON e.id = ie.evento_id
            JOIN associati a ON a.id = ie.associato_id
            WHERE pe.gruppo_ricevuta LIKE 'MGR-%' AND substr(COALESCE(e.data_evento, ''), 1, 4) = ?
        )
        SELECT
            gruppo_ricevuta,
            associato,
            data_pagamento,
            COUNT(*) AS numero_scadenze,
            COALESCE(SUM(importo), 0) AS importo_totale
        FROM pagamenti
        GROUP BY gruppo_ricevuta, associato, data_pagamento
        ORDER BY data_pagamento DESC, gruppo_ricevuta DESC
        """,
        (work_year, work_year, work_year, str(work_year)),
    )


def delete_multi_area_group(group_code: str) -> None:
    if not group_code:
        raise ValueError("Pagamento multi-area non valido.")
    with get_connection() as connection:
        connection.execute("DELETE FROM pagamenti_tesseramenti WHERE gruppo_ricevuta = ?", (group_code,))
        connection.execute("DELETE FROM pagamenti_rate_corsi WHERE gruppo_ricevuta = ?", (group_code,))
        connection.execute("DELETE FROM pagamenti_campi_estivi WHERE gruppo_ricevuta = ?", (group_code,))
        connection.execute("DELETE FROM pagamenti_eventi WHERE gruppo_ricevuta = ?", (group_code,))
        connection.commit()


def boolean_options(selected: str = "1") -> str:
    return render_static_options([("1", "Si"), ("0", "No")], selected, blank_label=None)


def associato_status_options(selected: str = "Attivo") -> str:
    return render_static_options(
        [("Attivo", "Attivo"), ("Sospeso", "Sospeso"), ("Dimesso", "Dimesso")],
        selected,
        blank_label=None,
    )


def sesso_options(selected: str = "M") -> str:
    return render_static_options(
        [("M", "Maschio"), ("F", "Femmina")],
        selected,
        blank_label=None,
    )


def carica_options(selected: str = "Associato", current_user: dict[str, object] | None = None) -> str:
    return render_static_options(
        [(value, value) for value in available_cariche(current_user)],
        selected,
        blank_label=None,
    )


def corso_enrollment_status_options(selected: str = "Attiva") -> str:
    return render_static_options(
        [("Attiva", "Attiva"), ("Sospesa", "Sospesa"), ("Chiusa", "Chiusa")],
        selected,
        blank_label=None,
    )


def camp_enrollment_status_options(selected: str = "Iscritto") -> str:
    return render_static_options(
        [("Iscritto", "Iscritto"), ("Lista attesa", "Lista attesa"), ("Annullato", "Annullato")],
        selected,
        blank_label=None,
    )


def event_enrollment_status_options(selected: str = "Iscritto") -> str:
    return render_static_options(
        [("Iscritto", "Iscritto"), ("Confermato", "Confermato"), ("Annullato", "Annullato")],
        selected,
        blank_label=None,
    )


def edit_path(entity_key: str, record_id: object, query_params: dict[str, str]) -> str:
    return with_query(f"/modifica/{entity_key}/{record_id}", current_page_query(query_params))


def render_associato_edit_fields(row: sqlite3.Row, current_user: dict[str, object] | None) -> list[str]:
    return [
        readonly_field(
            "Numero progressivo",
            str(row["numero_progressivo"] or ""),
        ),
        input_field("Codice associato", "codice_associato", value=row["codice_associato"] or "", required_field=True),
        input_field("Nome", "nome", value=row["nome"] or "", required_field=True),
        input_field("Cognome", "cognome", value=row["cognome"] or "", required_field=True),
        input_field(
            "Codice fiscale",
            "codice_fiscale",
            value=row["codice_fiscale"] or "",
            attrs={
                "maxlength": "16",
                "data-codice-fiscale": "true",
                "autocomplete": "off",
                "autocapitalize": "characters",
                "spellcheck": "false",
            },
        ),
        input_field("Data nascita", "data_nascita", input_type="date", value=row["data_nascita"] or ""),
        select_field("Sesso", "sesso", sesso_options(row["sesso"] or "M")),
        input_field("Comune di nascita", "comune_nascita", value=row["comune_nascita"] or ""),
        input_field("Provincia di nascita", "provincia_nascita", value=row["provincia_nascita"] or "", attrs={"maxlength": "2", "autocapitalize": "characters"}),
        input_field("Telefono", "telefono", value=row["telefono"] or ""),
        input_field("Email", "email", input_type="email", value=row["email"] or ""),
        input_field("Indirizzo", "indirizzo", value=row["indirizzo"] or "", wide=True),
        input_field("CAP", "cap", value=row["cap"] or "", attrs={"maxlength": "5", "inputmode": "numeric"}),
        input_field("Citta", "citta", value=row["citta"] or ""),
        input_field("Provincia", "provincia", value=row["provincia"] or "", attrs={"maxlength": "2", "autocapitalize": "characters"}),
        input_field(
            "Data prima iscrizione",
            "data_prima_iscrizione",
            input_type="date",
            value=row["data_prima_iscrizione"] or "",
            required_field=True,
        ),
        select_field(
            "Stato associato",
            "stato_associato",
            associato_status_options(row["stato_associato"] or "Attivo"),
        ),
        associato_carica_field(row["carica"] or "Associato", current_user),
        textarea_field("Note", "note", value=row["note"] or ""),
    ]


CRUD_CONFIG = {
    "associati": {
        "page_title": "Modifica associato",
        "page_subtitle": "Aggiorna anagrafica e contatti del socio.",
        "return_path": "/maschere/associati",
        "return_query": {"vista": "dati"},
        "fetch_query": "SELECT * FROM associati WHERE id = ?",
        "update_sql": """
            UPDATE associati
            SET codice_associato = ?, nome = ?, cognome = ?, codice_fiscale = ?, data_nascita = ?,
                sesso = ?, comune_nascita = ?, provincia_nascita = ?, carica = ?, email = ?, telefono = ?, indirizzo = ?, cap = ?, citta = ?, provincia = ?,
                data_prima_iscrizione = ?, stato_associato = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "codice_associato", "Codice associato"),
            required(form_data, "nome", "Nome"),
            required(form_data, "cognome", "Cognome"),
            optional(form_data, "codice_fiscale"),
            optional(form_data, "data_nascita"),
            normalized(form_data, "sesso", "M") or "M",
            optional(form_data, "comune_nascita"),
            optional(form_data, "provincia_nascita"),
            normalized(form_data, "carica", "Associato") or "Associato",
            optional(form_data, "email"),
            optional(form_data, "telefono"),
            optional(form_data, "indirizzo"),
            optional(form_data, "cap"),
            optional(form_data, "citta"),
            optional(form_data, "provincia"),
            required(form_data, "data_prima_iscrizione", "Data prima iscrizione"),
            normalized(form_data, "stato_associato", "Attivo") or "Attivo",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: readonly_field(
                "Numero progressivo",
                str(row["numero_progressivo"] or ""),
            ),
            lambda row: input_field("Codice associato", "codice_associato", value=row["codice_associato"] or "", required_field=True),
            lambda row: input_field("Nome", "nome", value=row["nome"] or "", required_field=True),
            lambda row: input_field("Cognome", "cognome", value=row["cognome"] or "", required_field=True),
            lambda row: input_field(
                "Codice fiscale",
                "codice_fiscale",
                value=row["codice_fiscale"] or "",
                attrs={
                    "maxlength": "16",
                    "data-codice-fiscale": "true",
                    "autocomplete": "off",
                    "autocapitalize": "characters",
                    "spellcheck": "false",
                },
            ),
            lambda row: input_field("Data nascita", "data_nascita", input_type="date", value=row["data_nascita"] or ""),
            lambda row: select_field("Sesso", "sesso", sesso_options(row["sesso"] or "M")),
            lambda row: input_field("Comune di nascita", "comune_nascita", value=row["comune_nascita"] or ""),
            lambda row: input_field("Provincia di nascita", "provincia_nascita", value=row["provincia_nascita"] or "", attrs={"maxlength": "2", "autocapitalize": "characters"}),
            lambda row: input_field("Telefono", "telefono", value=row["telefono"] or ""),
            lambda row: input_field("Email", "email", input_type="email", value=row["email"] or ""),
            lambda row: input_field("Indirizzo", "indirizzo", value=row["indirizzo"] or "", wide=True),
            lambda row: input_field("CAP", "cap", value=row["cap"] or "", attrs={"maxlength": "5", "inputmode": "numeric"}),
            lambda row: input_field("Citta", "citta", value=row["citta"] or ""),
            lambda row: input_field("Provincia", "provincia", value=row["provincia"] or "", attrs={"maxlength": "2", "autocapitalize": "characters"}),
            lambda row: input_field(
                "Data prima iscrizione",
                "data_prima_iscrizione",
                input_type="date",
                value=row["data_prima_iscrizione"] or "",
                required_field=True,
            ),
            lambda row: select_field(
                "Stato associato",
                "stato_associato",
                associato_status_options(row["stato_associato"] or "Attivo"),
            ),
            lambda row: select_field("Carica", "carica", carica_options(row["carica"] or "Associato")),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM associati WHERE id = ?",
        "delete_prompt": "Eliminare questo associato? Verranno eliminati anche tesseramenti, iscrizioni e pagamenti collegati.",
        "success_update": "Associato aggiornato.",
        "success_delete": "Associato eliminato.",
    },
    "tesseramenti_annuali": {
        "page_title": "Modifica tesseramento",
        "page_subtitle": "Aggiorna anno sociale, importi e scadenze.",
        "return_path": "/maschere/tesseramenti",
        "fetch_query": "SELECT * FROM tesseramenti_annuali WHERE id = ?",
        "update_sql": """
            UPDATE tesseramenti_annuali
            SET associato_id = ?, anno_sociale = ?, data_tesseramento = ?, importo_dovuto = ?, data_scadenza = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "associato_id", "Associato"),
            required(form_data, "anno_sociale", "Anno sociale"),
            required(form_data, "data_tesseramento", "Data tesseramento"),
            required(form_data, "importo_dovuto", "Importo dovuto"),
            optional(form_data, "data_scadenza"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Associato", "associato_id", render_associato_options(associati_options(), str(row["associato_id"])), required_field=True, wide=True, searchable=True),
            lambda row: input_field("Anno sociale", "anno_sociale", input_type="number", value=str(row["anno_sociale"] or ""), required_field=True, minimum="2000"),
            lambda row: input_field("Data tesseramento", "data_tesseramento", input_type="date", value=row["data_tesseramento"] or "", required_field=True),
            lambda row: input_field("Importo dovuto", "importo_dovuto", input_type="number", value=str(row["importo_dovuto"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: input_field("Data scadenza", "data_scadenza", input_type="date", value=row["data_scadenza"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM tesseramenti_annuali WHERE id = ?",
        "delete_prompt": "Eliminare questo tesseramento e i pagamenti collegati?",
        "success_update": "Tesseramento aggiornato.",
        "success_delete": "Tesseramento eliminato.",
    },
    "pagamenti_tesseramenti": {
        "page_title": "Modifica pagamento tesseramento",
        "page_subtitle": "Correggi data, importo o metodo del pagamento.",
        "return_path": "/maschere/tesseramenti",
        "fetch_query": "SELECT * FROM pagamenti_tesseramenti WHERE id = ?",
        "update_sql": """
            UPDATE pagamenti_tesseramenti
            SET tesseramento_id = ?, data_pagamento = ?, importo = ?, metodo_pagamento_id = ?, riferimento = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "tesseramento_id", "Tesseramento"),
            required(form_data, "data_pagamento", "Data pagamento"),
            required(form_data, "importo", "Importo"),
            optional(form_data, "metodo_pagamento_id"),
            optional(form_data, "riferimento"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Tesseramento", "tesseramento_id", render_select_options(tesseramenti_options(), str(row["tesseramento_id"])), required_field=True, wide=True),
            lambda row: input_field("Data pagamento", "data_pagamento", input_type="date", value=row["data_pagamento"] or "", required_field=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0.01", required_field=True),
            lambda row: select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi_options(), str(row["metodo_pagamento_id"] or "")), wide=True),
            lambda row: input_field("Riferimento", "riferimento", value=row["riferimento"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM pagamenti_tesseramenti WHERE id = ?",
        "delete_prompt": "Eliminare questo pagamento del tesseramento?",
        "success_update": "Pagamento tesseramento aggiornato.",
        "success_delete": "Pagamento tesseramento eliminato.",
    },
    "tipologie_corsi": {
        "page_title": "Modifica tipologia corso",
        "page_subtitle": "Aggiorna codice, nome e stato della tipologia.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM tipologie_corsi WHERE id = ?",
        "update_sql": """
            UPDATE tipologie_corsi
            SET codice_tipologia = ?, nome = ?, descrizione = ?, attiva = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "codice_tipologia", "Codice tipologia"),
            required(form_data, "nome", "Nome tipologia"),
            optional(form_data, "descrizione"),
            normalized(form_data, "attiva", "1") or "1",
        ),
        "fields": [
            lambda row: input_field("Codice tipologia", "codice_tipologia", value=row["codice_tipologia"] or "", required_field=True),
            lambda row: input_field("Nome tipologia", "nome", value=row["nome"] or "", required_field=True),
            lambda row: select_field("Attiva", "attiva", boolean_options(str(row["attiva"] if row["attiva"] is not None else 1))),
            lambda row: textarea_field("Descrizione", "descrizione", value=row["descrizione"] or ""),
        ],
        "delete_sql": "DELETE FROM tipologie_corsi WHERE id = ?",
        "delete_prompt": "Eliminare questa tipologia corso?",
        "success_update": "Tipologia corso aggiornata.",
        "success_delete": "Tipologia corso eliminata.",
    },
    "quote_tesseramenti": {
        "page_title": "Modifica quota tesseramento",
        "page_subtitle": "Aggiorna descrizione, importo e stato della quota predefinita.",
        "return_path": "/maschere/tesseramenti",
        "return_query": {"vista": "dati"},
        "fetch_query": "SELECT * FROM quote_predefinite WHERE id = ? AND area = 'tesseramenti'",
        "update_sql": """
            UPDATE quote_predefinite
            SET descrizione = ?, importo = ?, attiva = ?, note = ?
            WHERE id = ? AND area = 'tesseramenti'
        """,
        "build_params": lambda form_data: (
            required(form_data, "descrizione", "Descrizione"),
            required(form_data, "importo", "Importo"),
            normalized(form_data, "attiva", "1") or "1",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: input_field("Descrizione", "descrizione", value=row["descrizione"] or "", required_field=True, wide=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: select_field("Attiva", "attiva", boolean_options(str(row["attiva"] if row["attiva"] is not None else 1))),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM quote_predefinite WHERE id = ? AND area = 'tesseramenti'",
        "delete_prompt": "Eliminare questa quota predefinita di tesseramento?",
        "success_update": "Quota tesseramento aggiornata.",
        "success_delete": "Quota tesseramento eliminata.",
    },
    "quote_campi_estivi": {
        "page_title": "Modifica quota Campo estivo",
        "page_subtitle": "Aggiorna descrizione, importo e stato della quota predefinita.",
        "return_path": "/maschere/campi-estivi",
        "return_query": {"vista": "dati"},
        "fetch_query": "SELECT * FROM quote_predefinite WHERE id = ? AND area = 'campi-estivi'",
        "update_sql": """
            UPDATE quote_predefinite
            SET descrizione = ?, importo = ?, attiva = ?, note = ?
            WHERE id = ? AND area = 'campi-estivi'
        """,
        "build_params": lambda form_data: (
            required(form_data, "descrizione", "Descrizione"),
            required(form_data, "importo", "Importo"),
            normalized(form_data, "attiva", "1") or "1",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: input_field("Descrizione", "descrizione", value=row["descrizione"] or "", required_field=True, wide=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: select_field("Attiva", "attiva", boolean_options(str(row["attiva"] if row["attiva"] is not None else 1))),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM quote_predefinite WHERE id = ? AND area = 'campi-estivi'",
        "delete_prompt": "Eliminare questa quota predefinita di Campo estivo?",
        "success_update": "Quota Campo estivo aggiornata.",
        "success_delete": "Quota Campo estivo eliminata.",
    },
    "corsi": {
        "page_title": "Modifica corso",
        "page_subtitle": "Aggiorna anagrafica corso, quota mensile e organizzazione.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM corsi WHERE id = ?",
        "update_sql": """
            UPDATE corsi
            SET codice_corso = ?, nome = ?, descrizione = ?, quota_iscrizione_standard = 0,
                quota_mensile_standard = ?, sede = ?, giorno_settimana = ?, orario = ?, attivo = ?, note = ?,
                tipologia_corso_id = NULL
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "codice_corso", "Codice corso"),
            required(form_data, "nome", "Nome corso"),
            optional(form_data, "descrizione"),
            normalized(form_data, "quota_mensile_standard", "0"),
            optional(form_data, "sede"),
            optional(form_data, "giorno_settimana"),
            optional(form_data, "orario"),
            normalized(form_data, "attivo", "1") or "1",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: readonly_field("Numero progressivo", str(row["numero_progressivo"] or "")),
            lambda row: input_field("Codice corso", "codice_corso", value=row["codice_corso"] or "", required_field=True),
            lambda row: input_field("Nome corso", "nome", value=row["nome"] or "", required_field=True),
            lambda row: input_field("Quota mensile standard", "quota_mensile_standard", input_type="number", value=str(row["quota_mensile_standard"] or ""), step="0.01", minimum="0"),
            lambda row: input_field("Sede", "sede", value=row["sede"] or ""),
            lambda row: input_field("Giorno settimana", "giorno_settimana", value=row["giorno_settimana"] or ""),
            lambda row: input_field("Orario", "orario", value=row["orario"] or ""),
            lambda row: select_field("Attivo", "attivo", boolean_options(str(row["attivo"] if row["attivo"] is not None else 1))),
            lambda row: textarea_field("Descrizione", "descrizione", value=row["descrizione"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM corsi WHERE id = ?",
        "delete_prompt": "Eliminare questo corso? Le iscrizioni collegate devono essere eliminate prima.",
        "success_update": "Corso aggiornato.",
        "success_delete": "Corso eliminato.",
    },
    "iscrizioni_corsi": {
        "page_title": "Modifica iscrizione corso",
        "page_subtitle": "Aggiorna partecipazione e quota mensile del corso.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM iscrizioni_corsi WHERE id = ?",
        "update_sql": """
            UPDATE iscrizioni_corsi
            SET associato_id = ?, corso_id = ?, data_iscrizione = ?, data_inizio = ?, data_fine = ?,
                quota_mensile = ?, stato_iscrizione = ?, note = ?, quota_iscrizione = 0
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "associato_id", "Associato"),
            required(form_data, "corso_id", "Corso"),
            required(form_data, "data_iscrizione", "Data iscrizione"),
            optional(form_data, "data_inizio"),
            optional(form_data, "data_fine"),
            required(form_data, "quota_mensile", "Quota mensile"),
            normalized(form_data, "stato_iscrizione", "Attiva") or "Attiva",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Associato", "associato_id", render_associato_options(associati_options(), str(row["associato_id"])), required_field=True, wide=True, searchable=True),
            lambda row: select_field("Corso", "corso_id", render_select_options(corsi_options(), str(row["corso_id"])), required_field=True, wide=True),
            lambda row: input_field("Data iscrizione", "data_iscrizione", input_type="date", value=row["data_iscrizione"] or "", required_field=True),
            lambda row: input_field("Data inizio", "data_inizio", input_type="date", value=row["data_inizio"] or ""),
            lambda row: input_field("Data fine", "data_fine", input_type="date", value=row["data_fine"] or ""),
            lambda row: input_field("Quota mensile", "quota_mensile", input_type="number", value=str(row["quota_mensile"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: select_field("Stato iscrizione", "stato_iscrizione", corso_enrollment_status_options(row["stato_iscrizione"] or "Attiva")),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM iscrizioni_corsi WHERE id = ?",
        "delete_prompt": "Eliminare questa iscrizione corso e tutti i pagamenti collegati?",
        "success_update": "Iscrizione corso aggiornata.",
        "success_delete": "Iscrizione corso eliminata.",
    },
    "pagamenti_iscrizioni_corsi": {
        "page_title": "Modifica pagamento iscrizione corso",
        "page_subtitle": "Aggiorna il versamento iniziale del corso.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM pagamenti_iscrizioni_corsi WHERE id = ?",
        "update_sql": """
            UPDATE pagamenti_iscrizioni_corsi
            SET iscrizione_corso_id = ?, data_pagamento = ?, importo = ?, metodo_pagamento_id = ?, riferimento = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "iscrizione_corso_id", "Iscrizione corso"),
            required(form_data, "data_pagamento", "Data pagamento"),
            required(form_data, "importo", "Importo"),
            optional(form_data, "metodo_pagamento_id"),
            optional(form_data, "riferimento"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Iscrizione corso", "iscrizione_corso_id", render_select_options(iscrizioni_corsi_options(), str(row["iscrizione_corso_id"])), required_field=True, wide=True),
            lambda row: input_field("Data pagamento", "data_pagamento", input_type="date", value=row["data_pagamento"] or "", required_field=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0.01", required_field=True),
            lambda row: select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi_options(), str(row["metodo_pagamento_id"] or "")), wide=True),
            lambda row: input_field("Riferimento", "riferimento", value=row["riferimento"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM pagamenti_iscrizioni_corsi WHERE id = ?",
        "delete_prompt": "Eliminare questo pagamento di iscrizione corso?",
        "success_update": "Pagamento iscrizione corso aggiornato.",
        "success_delete": "Pagamento iscrizione corso eliminato.",
    },
    "rate_corsi_mensili": {
        "page_title": "Modifica quota mensile corso",
        "page_subtitle": "Aggiorna competenza, scadenza e importo della rata.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM rate_corsi_mensili WHERE id = ?",
        "update_sql": """
            UPDATE rate_corsi_mensili
            SET iscrizione_corso_id = ?, anno = ?, mese = ?, importo_dovuto = ?, data_scadenza = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "iscrizione_corso_id", "Iscrizione corso"),
            required(form_data, "anno", "Anno"),
            required(form_data, "mese", "Mese"),
            required(form_data, "importo_dovuto", "Importo dovuto"),
            optional(form_data, "data_scadenza"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Iscrizione corso", "iscrizione_corso_id", render_select_options(iscrizioni_corsi_options(), str(row["iscrizione_corso_id"])), required_field=True, wide=True),
            lambda row: input_field("Anno", "anno", input_type="number", value=str(row["anno"] or ""), required_field=True, minimum="2000"),
            lambda row: input_field("Mese", "mese", input_type="number", value=str(row["mese"] or ""), required_field=True, minimum="1"),
            lambda row: input_field("Importo dovuto", "importo_dovuto", input_type="number", value=str(row["importo_dovuto"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: input_field("Data scadenza", "data_scadenza", input_type="date", value=row["data_scadenza"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM rate_corsi_mensili WHERE id = ?",
        "delete_prompt": "Eliminare questa quota mensile e i pagamenti collegati?",
        "success_update": "Quota mensile aggiornata.",
        "success_delete": "Quota mensile eliminata.",
    },
    "pagamenti_rate_corsi": {
        "page_title": "Modifica pagamento quota mensile",
        "page_subtitle": "Aggiorna il pagamento della rata mensile.",
        "return_path": "/maschere/corsi",
        "fetch_query": "SELECT * FROM pagamenti_rate_corsi WHERE id = ?",
        "update_sql": """
            UPDATE pagamenti_rate_corsi
            SET rata_corso_id = ?, data_pagamento = ?, importo = ?, metodo_pagamento_id = ?, riferimento = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "rata_corso_id", "Quota mensile"),
            required(form_data, "data_pagamento", "Data pagamento"),
            required(form_data, "importo", "Importo"),
            optional(form_data, "metodo_pagamento_id"),
            optional(form_data, "riferimento"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Quota mensile", "rata_corso_id", render_select_options(rate_corsi_options(), str(row["rata_corso_id"])), required_field=True, wide=True),
            lambda row: input_field("Data pagamento", "data_pagamento", input_type="date", value=row["data_pagamento"] or "", required_field=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0.01", required_field=True),
            lambda row: select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi_options(), str(row["metodo_pagamento_id"] or "")), wide=True),
            lambda row: input_field("Riferimento", "riferimento", value=row["riferimento"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM pagamenti_rate_corsi WHERE id = ?",
        "delete_prompt": "Eliminare questo pagamento della quota mensile?",
        "success_update": "Pagamento quota mensile aggiornato.",
        "success_delete": "Pagamento quota mensile eliminato.",
    },
    "campi_estivi": {
        "page_title": "Modifica Campo estivo",
        "page_subtitle": "Aggiorna i dati interni del Campo estivo.",
        "return_path": "/maschere/campi-estivi",
        "fetch_query": "SELECT * FROM campi_estivi WHERE id = ?",
        "update_sql": """
            UPDATE campi_estivi
            SET codice_campo = ?, nome = ?, anno = ?, data_inizio = ?, data_fine = ?, sede = ?,
                quota_partecipazione_standard = ?, posti_massimi = ?, descrizione = ?, attivo = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "codice_campo", "Codice campo"),
            required(form_data, "nome", "Nome"),
            required(form_data, "anno", "Anno"),
            required(form_data, "data_inizio", "Data inizio"),
            required(form_data, "data_fine", "Data fine"),
            optional(form_data, "sede"),
            normalized(form_data, "quota_partecipazione_standard", "0"),
            optional(form_data, "posti_massimi"),
            optional(form_data, "descrizione"),
            normalized(form_data, "attivo", "1") or "1",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: readonly_field("Numero progressivo", str(row["numero_progressivo"] or "")),
            lambda row: input_field("Codice campo", "codice_campo", value=row["codice_campo"] or "", required_field=True),
            lambda row: input_field("Nome", "nome", value=row["nome"] or "", required_field=True),
            lambda row: input_field("Anno", "anno", input_type="number", value=str(row["anno"] or ""), required_field=True, minimum="2000"),
            lambda row: input_field("Data inizio", "data_inizio", input_type="date", value=row["data_inizio"] or "", required_field=True),
            lambda row: input_field("Data fine", "data_fine", input_type="date", value=row["data_fine"] or "", required_field=True),
            lambda row: input_field("Sede", "sede", value=row["sede"] or ""),
            lambda row: input_field("Quota standard", "quota_partecipazione_standard", input_type="number", value=str(row["quota_partecipazione_standard"] or ""), step="0.01", minimum="0"),
            lambda row: input_field("Posti massimi", "posti_massimi", input_type="number", value=str(row["posti_massimi"] or "") if row["posti_massimi"] is not None else "", minimum="1"),
            lambda row: select_field("Attivo", "attivo", boolean_options(str(row["attivo"] if row["attivo"] is not None else 1))),
            lambda row: textarea_field("Descrizione", "descrizione", value=row["descrizione"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM campi_estivi WHERE id = ?",
        "delete_prompt": "Eliminare questo Campo estivo e tutti i dati collegati?",
        "success_update": "Campo estivo aggiornato.",
        "success_delete": "Campo estivo eliminato.",
    },
    "iscrizioni_campi_estivi": {
        "page_title": "Modifica iscrizione Campo estivo",
        "page_subtitle": "Aggiorna partecipante, stato e quota di partecipazione.",
        "return_path": "/maschere/campi-estivi",
        "fetch_query": """
            SELECT ice.*, ce.anno AS anno_estate
            FROM iscrizioni_campi_estivi ice
            JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
            WHERE ice.id = ?
        """,
        "update_sql": """
            UPDATE iscrizioni_campi_estivi
            SET associato_id = ?, data_iscrizione = ?, quota_partecipazione = ?, stato_iscrizione = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "associato_id", "Associato"),
            required(form_data, "data_iscrizione", "Data iscrizione"),
            required(form_data, "quota_partecipazione", "Quota partecipazione"),
            normalized(form_data, "stato_iscrizione", "Iscritto") or "Iscritto",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Associato", "associato_id", render_associato_options(associati_options(), str(row["associato_id"])), required_field=True, wide=True, searchable=True),
            lambda row: readonly_field("Anno", str(row["anno_estate"] or "")),
            lambda row: input_field("Data iscrizione", "data_iscrizione", input_type="date", value=row["data_iscrizione"] or "", required_field=True),
            lambda row: input_field("Quota partecipazione", "quota_partecipazione", input_type="number", value=str(row["quota_partecipazione"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: select_field("Stato iscrizione", "stato_iscrizione", camp_enrollment_status_options(row["stato_iscrizione"] or "Iscritto")),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM iscrizioni_campi_estivi WHERE id = ?",
        "delete_prompt": "Eliminare questa iscrizione al Campo estivo e il pagamento collegato?",
        "success_update": "Iscrizione Campo estivo aggiornata.",
        "success_delete": "Iscrizione Campo estivo eliminata.",
    },
    "pagamenti_campi_estivi": {
        "page_title": "Modifica pagamento Campo estivo",
        "page_subtitle": "Aggiorna il pagamento una tantum del Campo estivo.",
        "return_path": "/maschere/campi-estivi",
        "fetch_query": "SELECT * FROM pagamenti_campi_estivi WHERE id = ?",
        "update_sql": """
            UPDATE pagamenti_campi_estivi
            SET iscrizione_campo_id = ?, data_pagamento = ?, importo = ?, metodo_pagamento_id = ?, riferimento = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "iscrizione_campo_id", "Iscrizione Campo estivo"),
            required(form_data, "data_pagamento", "Data pagamento"),
            required(form_data, "importo", "Importo"),
            optional(form_data, "metodo_pagamento_id"),
            optional(form_data, "riferimento"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Iscrizione Campo estivo", "iscrizione_campo_id", render_select_options(iscrizioni_campi_options(), str(row["iscrizione_campo_id"])), required_field=True, wide=True),
            lambda row: input_field("Data pagamento", "data_pagamento", input_type="date", value=row["data_pagamento"] or "", required_field=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0.01", required_field=True),
            lambda row: select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi_options(), str(row["metodo_pagamento_id"] or "")), wide=True),
            lambda row: input_field("Riferimento", "riferimento", value=row["riferimento"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM pagamenti_campi_estivi WHERE id = ?",
        "delete_prompt": "Eliminare questo pagamento del Campo estivo?",
        "success_update": "Pagamento Campo estivo aggiornato.",
        "success_delete": "Pagamento Campo estivo eliminato.",
    },
    "eventi": {
        "page_title": "Modifica evento",
        "page_subtitle": "Aggiorna dati, quota e stato dell'evento.",
        "return_path": "/maschere/eventi",
        "fetch_query": "SELECT * FROM eventi WHERE id = ?",
        "update_sql": """
            UPDATE eventi
            SET codice_evento = ?, nome = ?, tipologia = ?, data_evento = ?, luogo = ?,
                quota_partecipazione_standard = ?, posti_massimi = ?, descrizione = ?, attivo = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "codice_evento", "Codice evento"),
            required(form_data, "nome", "Nome evento"),
            optional(form_data, "tipologia"),
            required(form_data, "data_evento", "Data evento"),
            optional(form_data, "luogo"),
            normalized(form_data, "quota_partecipazione_standard", "0"),
            optional(form_data, "posti_massimi"),
            optional(form_data, "descrizione"),
            normalized(form_data, "attivo", "1") or "1",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: readonly_field("Numero progressivo", str(row["numero_progressivo"] or "")),
            lambda row: input_field("Codice evento", "codice_evento", value=row["codice_evento"] or "", required_field=True),
            lambda row: input_field("Nome evento", "nome", value=row["nome"] or "", required_field=True),
            lambda row: input_field("Tipologia", "tipologia", value=row["tipologia"] or ""),
            lambda row: input_field("Data evento", "data_evento", input_type="date", value=row["data_evento"] or "", required_field=True),
            lambda row: input_field("Luogo", "luogo", value=row["luogo"] or ""),
            lambda row: input_field("Quota standard", "quota_partecipazione_standard", input_type="number", value=str(row["quota_partecipazione_standard"] or ""), step="0.01", minimum="0"),
            lambda row: input_field("Posti massimi", "posti_massimi", input_type="number", value=str(row["posti_massimi"] or "") if row["posti_massimi"] is not None else "", minimum="1"),
            lambda row: select_field("Attivo", "attivo", boolean_options(str(row["attivo"] if row["attivo"] is not None else 1))),
            lambda row: textarea_field("Descrizione", "descrizione", value=row["descrizione"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM eventi WHERE id = ?",
        "delete_prompt": "Eliminare questo evento e tutte le iscrizioni collegate?",
        "success_update": "Evento aggiornato.",
        "success_delete": "Evento eliminato.",
    },
    "iscrizioni_eventi": {
        "page_title": "Modifica iscrizione evento",
        "page_subtitle": "Aggiorna partecipante, quota e stato dell'evento.",
        "return_path": "/maschere/eventi",
        "fetch_query": "SELECT * FROM iscrizioni_eventi WHERE id = ?",
        "update_sql": """
            UPDATE iscrizioni_eventi
            SET associato_id = ?, evento_id = ?, data_iscrizione = ?, quota_partecipazione = ?, stato_iscrizione = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "associato_id", "Associato"),
            required(form_data, "evento_id", "Evento"),
            required(form_data, "data_iscrizione", "Data iscrizione"),
            required(form_data, "quota_partecipazione", "Quota partecipazione"),
            normalized(form_data, "stato_iscrizione", "Iscritto") or "Iscritto",
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Associato", "associato_id", render_associato_options(associati_options(), str(row["associato_id"])), required_field=True, wide=True, searchable=True),
            lambda row: select_field("Evento", "evento_id", render_select_options(eventi_options(), str(row["evento_id"])), required_field=True, wide=True),
            lambda row: input_field("Data iscrizione", "data_iscrizione", input_type="date", value=row["data_iscrizione"] or "", required_field=True),
            lambda row: input_field("Quota partecipazione", "quota_partecipazione", input_type="number", value=str(row["quota_partecipazione"] or ""), step="0.01", minimum="0", required_field=True),
            lambda row: select_field("Stato iscrizione", "stato_iscrizione", event_enrollment_status_options(row["stato_iscrizione"] or "Iscritto")),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM iscrizioni_eventi WHERE id = ?",
        "delete_prompt": "Eliminare questa iscrizione evento e il pagamento collegato?",
        "success_update": "Iscrizione evento aggiornata.",
        "success_delete": "Iscrizione evento eliminata.",
    },
    "pagamenti_eventi": {
        "page_title": "Modifica pagamento evento",
        "page_subtitle": "Aggiorna il pagamento una tantum dell'evento.",
        "return_path": "/maschere/eventi",
        "fetch_query": "SELECT * FROM pagamenti_eventi WHERE id = ?",
        "update_sql": """
            UPDATE pagamenti_eventi
            SET iscrizione_evento_id = ?, data_pagamento = ?, importo = ?, metodo_pagamento_id = ?, riferimento = ?, note = ?
            WHERE id = ?
        """,
        "build_params": lambda form_data: (
            required(form_data, "iscrizione_evento_id", "Iscrizione evento"),
            required(form_data, "data_pagamento", "Data pagamento"),
            required(form_data, "importo", "Importo"),
            optional(form_data, "metodo_pagamento_id"),
            optional(form_data, "riferimento"),
            optional(form_data, "note"),
        ),
        "fields": [
            lambda row: select_field("Iscrizione evento", "iscrizione_evento_id", render_select_options(iscrizioni_eventi_options(), str(row["iscrizione_evento_id"])), required_field=True, wide=True),
            lambda row: input_field("Data pagamento", "data_pagamento", input_type="date", value=row["data_pagamento"] or "", required_field=True),
            lambda row: input_field("Importo", "importo", input_type="number", value=str(row["importo"] or ""), step="0.01", minimum="0.01", required_field=True),
            lambda row: select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi_options(), str(row["metodo_pagamento_id"] or "")), wide=True),
            lambda row: input_field("Riferimento", "riferimento", value=row["riferimento"] or ""),
            lambda row: textarea_field("Note", "note", value=row["note"] or ""),
        ],
        "delete_sql": "DELETE FROM pagamenti_eventi WHERE id = ?",
        "delete_prompt": "Eliminare questo pagamento evento?",
        "success_update": "Pagamento evento aggiornato.",
        "success_delete": "Pagamento evento eliminato.",
    },
}


def get_crud_config(entity_key: str) -> dict:
    config = CRUD_CONFIG.get(entity_key)
    if config is None:
        raise KeyError(entity_key)
    return config


def access_users_rows() -> list[sqlite3.Row]:
    return fetch_all(
        """
        SELECT
            id,
            username,
            COALESCE(email_recupero, '') AS email_recupero,
            CASE WHEN is_admin = 1 THEN 'Amministratore' ELSE 'Utente' END AS profilo,
            CASE WHEN attivo = 1 THEN 'Attivo' ELSE 'Disattivato' END AS stato,
            creato_il,
            COALESCE(ultimo_accesso, '') AS ultimo_accesso
        FROM utenti_accesso
        ORDER BY is_admin DESC, username
        """
    )


def login_page(query_params: dict[str, str]) -> bytes:
    next_target = login_destination(query_params.get("next"))
    first_access = bootstrap_admin_required()
    if first_access:
        title = "Primo accesso"
        subtitle = "Crea l'utente amministratore che potra poi generare gli altri utenti operativi."
        action = "/azioni/accesso/primo-admin"
        button_label = "Crea amministratore"
        intro_html = '<span class="auth-mode">Configurazione iniziale</span>'
        recovery_html = ""
    else:
        title = "Accesso"
        subtitle = "Inserisci username e password per entrare nel gestionale."
        action = "/azioni/accesso/login"
        button_label = "Accedi"
        intro_html = '<span class="auth-mode">Area riservata</span>'
        recovery_html = f"""
        <div class="auth-support-actions">
          <a class="button ghost" href="{esc(with_query('/recupera-password', {'next': next_target} if next_target != '/' else {}))}">Recupera password</a>
        </div>
        """

    form = form_card(
        title,
        subtitle,
        action,
        "".join(
            [
                input_field(
                    "Username",
                    "username",
                    required_field=True,
                    wide=True,
                    attrs={"autocomplete": "username"},
                ),
                input_field(
                    "Password",
                    "password",
                    input_type="password",
                    required_field=True,
                    wide=True,
                    revealable=True,
                    attrs={"autocomplete": "current-password" if not first_access else "new-password"},
                ),
                input_field(
                    "Conferma password",
                    "password_conferma",
                    input_type="password",
                    required_field=first_access,
                    wide=True,
                    revealable=True,
                    attrs={"autocomplete": "new-password"},
                ) if first_access else "",
            ]
        ),
        button_label,
        hidden_fields={"next": next_target},
    )
    content = f"""
    <div class="auth-panel-stack">
      <div class="auth-intro-card">
        {intro_html}
        <h2>{esc(title)}</h2>
        <p>{esc(subtitle)}</p>
      </div>
      {form}
      {recovery_html}
    </div>
    """
    return public_page(title, content, query_params)


def recover_password_page(query_params: dict[str, str]) -> bytes:
    next_target = login_destination(query_params.get("next"))
    form = form_card(
        "Recupera password",
        "Inserisci username ed email di recupero per impostare una nuova password del tuo utente standard.",
        "/azioni/accesso/recupera-password",
        "".join(
            [
                input_field("Username", "username", required_field=True, wide=True, attrs={"autocomplete": "username"}),
                input_field("Email recupero", "email_recupero", input_type="email", required_field=True, wide=True, attrs={"autocomplete": "email"}),
                input_field("Nuova password", "password", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
                input_field("Conferma nuova password", "password_conferma", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
            ]
        ),
        "Aggiorna password",
        hidden_fields={"next": next_target},
    )
    content = f"""
    <div class="auth-panel-stack">
      <div class="auth-intro-card">
        <span class="auth-mode">Recupero credenziali</span>
        <h2>Recupera password</h2>
        <p>Per ragioni di sicurezza la password attuale non viene mostrata: qui puoi solo impostarne una nuova dopo verifica.</p>
      </div>
      {form}
      <div class="auth-support-actions">
        <a class="button ghost" href="{esc(with_query('/login', {'next': next_target} if next_target != '/' else {}))}">Torna al login</a>
      </div>
    </div>
    """
    return public_page("Recupera password", content, query_params)


def access_profile_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    if current_user is None:
        raise KeyError("current_user")
    user_row = access_user_row(int(current_user["id"]))
    if user_row is None:
        raise KeyError("current_user")

    info_card = table_card(
        "Dati accesso",
        "Le password restano protette e non sono visualizzabili. Da questa area puoi aggiornarle in sicurezza.",
        [user_row],
        [
            ("username", "Username"),
            ("email_recupero", "Email recupero"),
            ("ultimo_accesso", "Ultimo accesso"),
            ("aggiornato_il", "Ultimo aggiornamento"),
        ],
    )
    change_password_form = form_card(
        "Cambia password",
        "Inserisci la password attuale e poi la nuova password.",
        "/azioni/accesso/cambia-password",
        "".join(
            [
                input_field("Password attuale", "password_attuale", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "current-password"}),
                input_field("Nuova password", "password", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
                input_field("Conferma nuova password", "password_conferma", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
            ]
        ),
        "Aggiorna password",
        hidden_fields=work_year_query(query_params),
    )
    content = f"""
    <div class="cards-grid">
      {change_password_form}
    </div>
    {info_card}
    """
    return page("Profilo accesso", "/maschere/accesso", content, query_params, current_user)


def utenti_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    users = access_users_rows()
    create_form = form_card(
        "Nuovo utente",
        "Crea un utente operativo. Gli utenti creati da qui non possono amministrare altri account.",
        "/azioni/utenti/crea",
        "".join(
            [
                readonly_field("Profilo", "Utente standard"),
                input_field("Username", "username", required_field=True, attrs={"autocomplete": "off"}),
                input_field("Password", "password", input_type="password", required_field=True, revealable=True, attrs={"autocomplete": "new-password"}),
                input_field("Conferma password", "password_conferma", input_type="password", required_field=True, revealable=True, attrs={"autocomplete": "new-password"}),
                input_field("Email recupero", "email_recupero", input_type="email", required_field=True, wide=True, attrs={"autocomplete": "off"}),
            ]
        ),
        "Salva utente",
        hidden_fields=work_year_query(query_params),
    )
    content = f"""
    <div class="cards-grid">
      {create_form}
    </div>
    {table_card(
        "Utenti registrati",
        "Le password non sono leggibili ne esportabili: l'amministratore puo pero reimpostarle e disattivare gli utenti standard.",
        users,
        [
            ("username", "Username"),
            ("email_recupero", "Email recupero"),
            ("profilo", "Profilo"),
            ("stato", "Stato"),
            ("creato_il", "Creato il"),
            ("ultimo_accesso", "Ultimo accesso"),
            (
                "id",
                "Azioni",
                lambda value, row: ""
                if row["profilo"] == "Amministratore"
                else action_links_html(
                    extra_links=[
                        (
                            with_query(f"/maschere/utenti/gestione/{value}", work_year_query(query_params)),
                            "Gestisci",
                        )
                    ]
                ),
            ),
        ],
    )}
    """
    return page("Utenti", "/maschere/utenti", content, query_params, current_user)


def gestione_utente_page(
    user_id: int,
    query_params: dict[str, str],
    current_user: dict[str, object] | None = None,
) -> bytes:
    if not current_user or not current_user.get("is_admin"):
        raise KeyError(user_id)
    user_row = access_user_row(user_id)
    if user_row is None or user_row["is_admin"]:
        raise KeyError(user_id)

    info_table = table_card(
        "Dati utente",
        "L'account puo essere aggiornato, disattivato o ricevere una nuova password, ma la password attuale resta protetta.",
        [user_row],
        [
            ("username", "Username"),
            ("email_recupero", "Email recupero"),
            ("attivo", "Attivo", lambda value, row: "Si" if int(value or 0) == 1 else "No"),
            ("creato_il", "Creato il"),
            ("ultimo_accesso", "Ultimo accesso"),
        ],
    )
    update_form = form_card(
        "Anagrafica utente",
        "Aggiorna username ed email di recupero del profilo standard.",
        f"/azioni/utenti/aggiorna/{user_id}",
        "".join(
            [
                input_field("Username", "username", required_field=True, value=user_row["username"] or ""),
                input_field("Email recupero", "email_recupero", input_type="email", required_field=True, wide=True, value=user_row["email_recupero"] or ""),
            ]
        ),
        "Salva dati",
        hidden_fields=work_year_query(query_params),
    )
    password_form = form_card(
        "Reimposta password",
        "Inserisci una nuova password per l'utente standard. La password attuale non e visualizzabile.",
        f"/azioni/utenti/password/{user_id}",
        "".join(
            [
                input_field("Nuova password", "password", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
                input_field("Conferma nuova password", "password_conferma", input_type="password", required_field=True, wide=True, revealable=True, attrs={"autocomplete": "new-password"}),
            ]
        ),
        "Aggiorna password",
        hidden_fields=work_year_query(query_params),
    )
    status_action = "riattiva" if int(user_row["attivo"] or 0) == 0 else "disattiva"
    status_label = "Riattiva utente" if status_action == "riattiva" else "Disattiva utente"
    status_form = f"""
    <section class="card">
      <div class="card-head">
        <h2>Stato utente</h2>
        <p>Puoi annullare l'utente standard disattivandolo. Se e disattivato, non potra accedere finche non lo riattivi.</p>
      </div>
      <form method="post" action="{esc(f'/azioni/utenti/stato/{user_id}')}" class="form-grid">
        {hidden_fields_html({**work_year_query(query_params), 'azione_stato': status_action})}
        <div class="form-actions">
          <button type="submit" class="button ghost">{esc(status_label)}</button>
        </div>
      </form>
    </section>
    """
    back_url = with_query("/maschere/utenti", work_year_query(query_params))
    content = f"""
    <section class="hero">
      <div>
        <span class="eyebrow">Gestione utenti</span>
        <h2>{esc(user_row['username'])}</h2>
        <p>Qui puoi modificare solo gli utenti standard. L'account amministratore resta gestito dal proprio profilo accesso.</p>
      </div>
      <div class="hero-actions">
        <a class="button ghost" href="{esc(back_url)}">Torna agli utenti</a>
      </div>
    </section>
    <div class="cards-grid">
      {update_form}
      {password_form}
    </div>
    {status_form}
    {info_table}
    """
    return page("Gestione utente", "/maschere/utenti", content, query_params, current_user)


def render_crud_edit_page(
    entity_key: str,
    record_id: int,
    query_params: dict[str, str],
    current_user: dict[str, object] | None = None,
) -> bytes:
    if entity_key in LOCKED_CRUD_ENTITIES:
        raise KeyError(entity_key)
    config = get_crud_config(entity_key)
    row = fetch_one(config["fetch_query"], (record_id,))
    if row is None:
        raise KeyError(entity_key)

    fields_html = (
        "".join(render_associato_edit_fields(row, current_user))
        if entity_key == "associati"
        else "".join(field(row) for field in config["fields"])
    )

    form = form_card(
        config["page_title"],
        config["page_subtitle"],
        f"/azioni/crud/aggiorna/{entity_key}/{record_id}",
        fields_html,
        "Salva modifiche",
        hidden_fields=current_page_query(query_params),
    )
    delete_block = f"""
    <section class="card screen-only">
      <div class="card-head">
        <h2>Eliminazione record</h2>
        <p>Usa questa azione solo se vuoi rimuovere definitivamente il dato selezionato.</p>
      </div>
      {delete_action_form(
          f"/azioni/crud/elimina/{entity_key}/{record_id}",
          config["delete_prompt"],
          extra_fields=current_page_query(query_params),
      )}
    </section>
    """
    back_url = with_query(config["return_path"], current_page_query(query_params))
    content = f"""
    <section class="hero">
      <div>
        <span class="eyebrow">Gestione dati</span>
        <h2>{esc(config["page_title"])}</h2>
        <p>Torna alla maschera principale dopo aver salvato o eliminato il record.</p>
      </div>
      <div class="hero-actions">
        <a class="button ghost" href="{esc(back_url)}">Torna alla maschera</a>
      </div>
    </section>
    {form}
    {delete_block}
    """
    return page(config["page_title"], config["return_path"], content, query_params, current_user)


def handle_crud_update(
    entity_key: str,
    record_id: int,
    form_data: dict[str, str],
    start_response,
    current_user: dict[str, object] | None = None,
):
    if entity_key in LOCKED_CRUD_ENTITIES:
        raise ValueError(LOCKED_CRUD_ENTITIES[entity_key])
    config = get_crud_config(entity_key)
    if entity_key == "associati":
        existing_row = fetch_one("SELECT carica FROM associati WHERE id = ?", (record_id,))
        if existing_row is None:
            raise ValueError("Associato non trovato.")
        params = (
            required(form_data, "codice_associato", "Codice associato"),
            required(form_data, "nome", "Nome"),
            required(form_data, "cognome", "Cognome"),
            optional(form_data, "codice_fiscale"),
            optional(form_data, "data_nascita"),
            normalized(form_data, "sesso", "M") or "M",
            optional(form_data, "comune_nascita"),
            optional(form_data, "provincia_nascita"),
            resolved_carica_value(form_data, current_user, existing_value=existing_row["carica"] or "Associato"),
            optional(form_data, "email"),
            optional(form_data, "telefono"),
            optional(form_data, "indirizzo"),
            optional(form_data, "cap"),
            optional(form_data, "citta"),
            optional(form_data, "provincia"),
            required(form_data, "data_prima_iscrizione", "Data prima iscrizione"),
            normalized(form_data, "stato_associato", "Attivo") or "Attivo",
            optional(form_data, "note"),
            record_id,
        )
    else:
        params = config["build_params"](form_data) + (record_id,)
    execute(config["update_sql"], params)
    extra_query = work_year_query_from_form(form_data)
    extra_query.update(config.get("return_query", {}))
    return redirect(
        start_response,
        config["return_path"],
        ok=config["success_update"],
        extra_query=extra_query,
    )


def handle_crud_delete(entity_key: str, record_id: int, form_data: dict[str, str], start_response):
    if entity_key in LOCKED_CRUD_ENTITIES:
        raise ValueError(LOCKED_CRUD_ENTITIES[entity_key])
    config = get_crud_config(entity_key)
    execute(config["delete_sql"], (record_id,))
    extra_query = work_year_query_from_form(form_data)
    extra_query.update(config.get("return_query", {}))
    return redirect(
        start_response,
        config["return_path"],
        ok=config["success_delete"],
        extra_query=extra_query,
    )


def course_generation_reminder_block(query_params: dict[str, str]) -> str:
    return ""


def dashboard_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    work_year = current_work_year(query_params)
    date_from, date_to = year_start_end(work_year)
    stats = [
        stat_card(
            f"Associati {work_year}",
            fetch_scalar(
                f"""
                SELECT COUNT(*)
                FROM associati a
                WHERE {associato_year_relevance_sql('a')}
                """,
                associato_year_relevance_params(work_year),
            ) or 0,
            with_query("/maschere/associati", data_view_query(query_params)),
        ),
        stat_card(
            f"Partecipanti corsi {work_year}",
            fetch_scalar(
                f"""
                SELECT COUNT(*)
                FROM iscrizioni_corsi ic
                WHERE {iscrizione_corso_year_relevance_sql('ic')}
                """,
                iscrizione_corso_year_relevance_params(work_year),
            ) or 0,
            with_query("/maschere/corsi", data_view_query(query_params)),
        ),
        stat_card(
            f"Partecipanti campo estivo {work_year}",
            fetch_scalar(
                """
                SELECT COUNT(*)
                FROM iscrizioni_campi_estivi ice
                JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
                WHERE ce.anno = ?
                """,
                (work_year,),
            ) or 0,
            with_query("/maschere/campi-estivi", data_view_query(query_params)),
        ),
        stat_card(
            f"Partecipanti eventi {work_year}",
            fetch_scalar(
                """
                SELECT COUNT(*)
                FROM iscrizioni_eventi ie
                JOIN eventi e ON e.id = ie.evento_id
                WHERE substr(e.data_evento, 1, 4) = ?
                """,
                (str(work_year),),
            ) or 0,
            with_query("/maschere/eventi", data_view_query(query_params)),
        ),
    ]

    associati_summary = fetch_all(
        posizione_associati_query(limit=20),
        posizione_associati_params(work_year),
    )
    associati_summary_columns = [
        ("codice_associato", "Codice"),
        (
            "associato",
            "Associato",
            lambda value, row: report_link(value, f"/report/associato/{row['associato_id']}", work_year_query(query_params)),
        ),
        ("stato_associato", "Stato"),
        ("totale_dovuto", "Totale dovuto", lambda value, _: money(value)),
        ("totale_pagato", "Totale pagato", lambda value, _: money(value)),
        ("saldo_residuo", "Saldo residuo", lambda value, _: money(value)),
    ]

    content = f"""
    <section class="stat-grid">
      {''.join(stats)}
    </section>
    {dashboard_associati_search_toolbar()}
    {table_card(
        f"Posizione associati anno {work_year}",
        "Clicca sul nome per aprire il dettaglio con incassi e scadenze del singolo associato.",
        associati_summary,
        associati_summary_columns,
        empty_message="Nessun associato disponibile.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(associati_summary, associati_summary_columns),
    )}
    {dashboard_charts(query_params)}
    """
    return page("Dashboard", "/", content, query_params, current_user)


def consiglio_direttivo_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    work_year = current_work_year(query_params)
    placeholders = ", ".join("?" for _ in CONSIGLIO_DIRETTIVO_ORDER)
    rows = fetch_all(
        f"""
        SELECT
            id,
            numero_progressivo,
            codice_associato,
            {associato_base_name_sql('')} AS associato,
            data_nascita,
            carica,
            COALESCE(telefono, '') AS telefono,
            COALESCE(email, '') AS email,
            COALESCE(citta, '') AS citta,
            stato_associato
        FROM associati a
        WHERE carica IN ({placeholders})
          AND {associato_year_relevance_sql('a')}
        ORDER BY
            CASE carica
                WHEN 'Presidente' THEN 1
                WHEN 'Vice Presidente' THEN 2
                WHEN 'Tesoriere' THEN 3
                WHEN 'Segretario' THEN 4
                WHEN 'Consigliere' THEN 5
                WHEN 'Consigliere spirituale' THEN 6
                ELSE 99
            END,
            cognome,
            nome
        """,
        (*CONSIGLIO_DIRETTIVO_ORDER, *associato_year_relevance_params(work_year)),
    )
    grouped_rows = {role: [] for role in CONSIGLIO_DIRETTIVO_ORDER}
    for row in rows:
        grouped_rows[row["carica"]].append(row)

    cards_html = []
    for role in CONSIGLIO_DIRETTIVO_ORDER:
        role_rows = grouped_rows.get(role, [])
        members_html = ""
        if role_rows:
            member_cards: list[str] = []
            for row in role_rows:
                metadata: list[str] = []
                display_name = esc(label_with_age(row["associato"], row["data_nascita"]))
                recapiti_parts: list[str] = []
                if row["telefono"]:
                    recapiti_parts.append(
                        f'<span class="direttivo-member-inline-item"><strong>Telefono</strong><span>{esc(row["telefono"])}</span></span>'
                    )
                if row["email"]:
                    recapiti_parts.append(
                        f'<span class="direttivo-member-inline-item"><strong>Email</strong><span>{esc(row["email"])}</span></span>'
                    )
                if recapiti_parts:
                    metadata.append(
                        f'<div class="direttivo-member-fact direttivo-member-fact-contacts is-wide"><div class="direttivo-member-inline">{"".join(recapiti_parts)}</div></div>'
                    )
                actions_html = ""
                if can_manage_cariche(current_user):
                    actions_html = (
                        f'<div class="direttivo-actions">'
                        f'{action_links_html(edit_href=edit_path("associati", row["id"], query_params), extra_fields=work_year_query(query_params))}'
                        f"</div>"
                    )
                member_cards.append(
                    f"""
                    <article class="direttivo-member">
                      <div class="direttivo-member-copy">
                        <header class="direttivo-member-header">
                          <h3>{display_name}</h3>
                        </header>
                        <div class="direttivo-member-meta">
                          {''.join(metadata) if metadata else '<div class="direttivo-member-fact direttivo-member-fact-contacts is-empty is-wide"><div class="direttivo-member-empty-note">Nessun recapito inserito</div></div>'}
                        </div>
                        {actions_html}
                      </div>
                    </article>
                    """
                )
            members_html = "".join(member_cards)
        else:
            members_html = '<div class="empty-state">Nessun associato con questa carica.</div>'

        cards_html.append(
            f"""
            <section class="card direttivo-role-card">
              <div class="direttivo-role-head">
                <h2>{esc(CONSIGLIO_DIRETTIVO_LABELS.get(role, role))}</h2>
                <div class="direttivo-role-accent"></div>
              </div>
              <div class="direttivo-role-body">
                {members_html}
              </div>
            </section>
            """
        )

    content = f"""
    <div class="cards-grid consiglio-direttivo-grid">
      {''.join(cards_html)}
    </div>
    """
    return page("Consiglio Direttivo", "/maschere/consiglio-direttivo", content, query_params, current_user)


def associati_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    work_year = current_work_year(query_params)
    associati = fetch_all(
        f"""
        SELECT
            id,
            numero_progressivo,
            codice_associato,
            {associato_display_sql('')} AS associato,
            data_nascita,
            carica,
            telefono,
            email,
            citta,
            stato_associato
        FROM associati a
        WHERE {associato_year_relevance_sql('a')}
        ORDER BY cognome, nome
        """,
        associato_year_relevance_params(work_year),
    )

    associati_form = form_card(
        "Nuovo associato",
        "Inserisci l'anagrafica di base del socio.",
        "/azioni/associati/crea",
        "".join(
            [
                readonly_field("Codice associato assegnato", peek_next_progressive_code("associati"), wide=True),
                input_field("Nome", "nome", required_field=True),
                input_field("Cognome", "cognome", required_field=True),
                input_field(
                    "Codice fiscale",
                    "codice_fiscale",
                    attrs={
                        "maxlength": "16",
                        "data-codice-fiscale": "true",
                        "autocomplete": "off",
                        "autocapitalize": "characters",
                        "spellcheck": "false",
                    },
                ),
                input_field("Data nascita", "data_nascita", input_type="date"),
                select_field("Sesso", "sesso", sesso_options("M")),
                input_field("Comune di nascita", "comune_nascita"),
                input_field("Provincia di nascita", "provincia_nascita", attrs={"maxlength": "2", "autocapitalize": "characters"}),
                input_field("Telefono", "telefono"),
                input_field("Email", "email", input_type="email"),
                input_field("Indirizzo", "indirizzo", wide=True),
                input_field("CAP", "cap", attrs={"maxlength": "5", "inputmode": "numeric"}),
                input_field("Citta", "citta"),
                input_field("Provincia", "provincia", attrs={"maxlength": "2", "autocapitalize": "characters"}),
                input_field(
                    "Data prima iscrizione",
                    "data_prima_iscrizione",
                    input_type="date",
                    value=date.today().isoformat(),
                    required_field=True,
                ),
                select_field(
                    "Stato associato",
                    "stato_associato",
                    associato_status_options("Attivo"),
                ),
                associato_carica_field("Associato", current_user),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva associato",
        hidden_fields=work_year_query(query_params),
    )

    forms_html = f"""
    <div class="cards-grid">
      {associati_form}
    </div>
    """ if not data_only else ""

    tables_html = f"""
    {table_card(
        "Associati registrati",
        "Elenco completo dell'anagrafica.",
        associati,
        [
            ("codice_associato", "Codice"),
            ("associato", "Associato"),
            ("carica", "Carica"),
            ("data_nascita", "Data nascita"),
            ("telefono", "Telefono"),
            ("email", "Email"),
            ("citta", "Citta"),
            ("stato_associato", "Stato"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("associati", value, query_params),
                    delete_action=f"/azioni/crud/elimina/associati/{value}",
                    delete_prompt="Eliminare questo associato e tutti i dati collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        table_class="search-table",
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/associati", query_params, "Apri dati associati")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati associati" if data_only else "Associati", "/maschere/associati", content, query_params, current_user)


def tesseramenti_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    year = str(current_work_year(query_params))
    associati = associati_options()
    metodi = metodi_options()
    metodo_predefinito = preferred_metodo_pagamento_id(metodi)
    tesseramenti = tesseramenti_aperti_options(int(year))
    quote_tesseramenti = quote_predefinite_options("tesseramenti")
    quota_tesseramento_default = quote_tesseramenti[0] if len(quote_tesseramenti) == 1 else None
    quota_tesseramento_default_id = str(quota_tesseramento_default["id"]) if quota_tesseramento_default else None
    quota_tesseramento_default_importo = str(quota_tesseramento_default["importo"]) if quota_tesseramento_default else ""
    tesseramento_scadenza_default = f"{year}-12-31"
    quote_tesseramenti_table = quote_predefinite_rows("tesseramenti")
    tesseramenti_table = fetch_all(
        f"""
        SELECT
            t.id,
            {associato_display_sql('a')} AS associato,
            t.anno_sociale,
            t.data_tesseramento,
            t.importo_dovuto,
            t.data_scadenza
        FROM tesseramenti_annuali t
        JOIN associati a ON a.id = t.associato_id
        WHERE t.anno_sociale = ?
        ORDER BY a.cognome, a.nome
        """,
        (int(year),),
    )
    payments_table = fetch_all(
        f"""
        SELECT
            pt.id,
            {associato_display_sql('a')} AS associato,
            t.anno_sociale,
            pt.data_pagamento,
            pt.importo,
            COALESCE(mp.nome, '') AS metodo_pagamento,
            COALESCE(pt.riferimento, '') AS riferimento
        FROM pagamenti_tesseramenti pt
        JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
        JOIN associati a ON a.id = t.associato_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id
        WHERE t.anno_sociale = ?
        ORDER BY pt.data_pagamento DESC, associato
        """,
        (int(year),),
    )
    saldo = fetch_all(
        """
        SELECT associato, anno_sociale, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
        FROM v_tesseramenti_saldo
        WHERE anno_sociale = ?
        ORDER BY anno_sociale DESC, associato
        """,
        (int(year),),
    )

    create_form = form_card(
        "Nuovo tesseramento",
        "Registra la quota annuale per un associato richiamando, se vuoi, una quota predefinita.",
        "/azioni/tesseramenti/crea",
        "".join(
            [
                select_field("Associato", "associato_id", render_associato_options(associati), required_field=True, wide=True, searchable=True),
                select_field(
                    "Quota tesseramento",
                    "quota_predefinita_id",
                    render_select_options(quote_tesseramenti, selected=quota_tesseramento_default_id, data_keys=["importo"]),
                    wide=True,
                    element_id="tesseramento-quota-select",
                    attrs={"data-amount-target": "tesseramento-importo-dovuto"},
                ),
                input_field(
                    "Anno sociale",
                    "anno_sociale",
                    input_type="number",
                    value=year,
                    required_field=True,
                    minimum="2000",
                    element_id="tesseramento-anno-sociale",
                    attrs={"data-year-end-target": "tesseramento-data-scadenza"},
                ),
                input_field("Data tesseramento", "data_tesseramento", input_type="date", value=date.today().isoformat(), required_field=True),
                input_field(
                    "Importo dovuto",
                    "importo_dovuto",
                    input_type="number",
                    step="0.01",
                    minimum="0",
                    required_field=True,
                    value=quota_tesseramento_default_importo,
                    element_id="tesseramento-importo-dovuto",
                ),
                input_field(
                    "Data scadenza",
                    "data_scadenza",
                    input_type="date",
                    value=tesseramento_scadenza_default,
                    element_id="tesseramento-data-scadenza",
                ),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva tesseramento",
        hidden_fields=work_year_query(query_params),
        form_attrs={
            "data-payment-flow": "tesseramento",
            "data-payment-amount-field": "importo_dovuto",
            "data-payment-method-default": metodo_predefinito,
            "data-payment-prompt-title": "Conferma tesseramento",
            "data-payment-prompt-message": "Vuoi procedere subito anche al pagamento del tesseramento?",
            "data-payment-prompt-yes": "Si, registra anche il pagamento",
            "data-payment-prompt-no": "No, solo tesseramento",
            "data-payment-dialog-title": "Pagamento tesseramento",
            "data-payment-dialog-message": "Conferma i dati del pagamento del tesseramento.",
            "data-payment-dialog-confirm": "Registra pagamento e genera ricevuta",
        },
    )

    forms_html = f"""
    <div class="cards-grid">
      {create_form}
    </div>
    """ if not data_only else ""

    tables_html = f"""
    {table_card(
        f"Tesseramenti anno {year}",
        "Anagrafica dei tesseramenti inseriti nell'anno di lavoro selezionato.",
        tesseramenti_table,
        [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("data_tesseramento", "Data"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("data_scadenza", "Scadenza"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("tesseramenti_annuali", value, query_params),
                    delete_action=f"/azioni/crud/elimina/tesseramenti_annuali/{value}",
                    delete_prompt="Eliminare questo tesseramento e i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        empty_message="Nessun tesseramento presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(tesseramenti_table, [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("data_tesseramento", "Data"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("data_scadenza", "Scadenza"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("tesseramenti_annuali", value, query_params),
                    delete_action=f"/azioni/crud/elimina/tesseramenti_annuali/{value}",
                    delete_prompt="Eliminare questo tesseramento e i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ]),
    )}
    {table_card(
        f"Pagamenti tesseramenti anno {year}",
        "Incassi registrati sulle quote annuali.",
        payments_table,
        [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("data_pagamento", "Data pagamento"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("metodo_pagamento", "Metodo"),
            ("riferimento", "Riferimento"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("pagamenti_tesseramenti", value, query_params),
                    extra_links=[(receipt_link("tesseramenti", value, query_params), "Ricevuta")],
                    delete_action=f"/azioni/crud/elimina/pagamenti_tesseramenti/{value}",
                    delete_prompt="Eliminare questo pagamento del tesseramento?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        empty_message="Nessun pagamento tesseramento presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(payments_table, [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("data_pagamento", "Data pagamento"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("metodo_pagamento", "Metodo"),
            ("riferimento", "Riferimento"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("pagamenti_tesseramenti", value, query_params),
                    extra_links=[(receipt_link("tesseramenti", value, query_params), "Ricevuta")],
                    delete_action=f"/azioni/crud/elimina/pagamenti_tesseramenti/{value}",
                    delete_prompt="Eliminare questo pagamento del tesseramento?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ]),
    )}
    {table_card(
        f"Situazione tesseramenti anno {year}",
        "Vista con dovuto, pagato e saldo residuo.",
        saldo,
        [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
        table_class="search-table",
        summary_rows=summary_rows_for_table(saldo, [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ]),
    )}
    {table_card(
        "Quota tesseramento",
        "Quota predefinita disponibile per la registrazione del tesseramento.",
        quote_tesseramenti_table,
        [
            ("descrizione", "Descrizione"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("attiva", "Attiva"),
            ("note", "Note"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("quote_tesseramenti", value, query_params),
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        empty_message="Nessuna quota tesseramento presente.",
        table_class="search-table",
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/tesseramenti", query_params, "Apri dati tesseramenti")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati tesseramenti" if data_only else "Tesseramenti", "/maschere/tesseramenti", content, query_params, current_user)


def corsi_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    work_year = current_work_year(query_params)
    associati = associati_options()
    corsi = corsi_options()
    metodi = metodi_options()
    metodo_predefinito = preferred_metodo_pagamento_id(metodi)
    open_rate_options = rate_corsi_aperte_options(work_year)
    corsi_table = fetch_all(
        f"""
        SELECT
            c.id,
            c.numero_progressivo,
            c.codice_corso,
            c.nome AS corso,
            c.quota_mensile_standard,
            c.giorno_settimana,
            c.orario,
            c.attivo
        FROM corsi c
        WHERE {corso_year_relevance_sql('c')}
        ORDER BY corso
        """,
        corso_year_relevance_params(work_year),
    )
    iscrizioni_table = fetch_all(
        f"""
        SELECT
            ic.id,
            {associato_display_sql('a')} AS associato,
            c.nome AS corso,
            ic.data_iscrizione,
            ic.data_inizio,
            ic.data_fine,
            ic.quota_mensile,
            ic.stato_iscrizione
        FROM iscrizioni_corsi ic
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        WHERE {iscrizione_corso_year_relevance_sql('ic')}
        ORDER BY ic.id DESC
        """,
        iscrizione_corso_year_relevance_params(work_year),
    )
    rate_table = fetch_all(
        """
        SELECT
            id,
            corso,
            associato,
            competenza,
            importo_dovuto,
            importo_pagato,
            saldo_residuo,
            stato_pagamento
        FROM v_rate_corsi_saldo
        WHERE anno = ?
        ORDER BY anno DESC, mese DESC, corso, associato
        """,
        (work_year,),
    )
    rate_payments_table = fetch_all(
        f"""
        SELECT
            prc.id,
            COALESCE(prc.gruppo_ricevuta, '') AS gruppo_ricevuta,
            {associato_display_sql('a')} AS associato,
            c.nome AS corso,
            printf('%04d-%02d', r.anno, r.mese) AS competenza,
            prc.data_pagamento,
            prc.importo,
            COALESCE(mp.nome, '') AS metodo_pagamento,
            COALESCE(prc.riferimento, '') AS riferimento
        FROM pagamenti_rate_corsi prc
        JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
        JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id
        WHERE r.anno = ?
        ORDER BY prc.data_pagamento DESC, associato
        """,
        (work_year,),
    )
    create_course_form = form_card(
        "Nuovo corso",
        "Crea un corso e definisci la quota mensile di riferimento.",
        "/azioni/corsi/crea",
        "".join(
            [
                readonly_field("Codice corso assegnato", peek_next_progressive_code("corsi")),
                input_field("Nome corso", "nome", required_field=True),
                input_field("Quota mensile", "quota_mensile_standard", input_type="number", step="0.01", minimum="0"),
                input_field("Giorno settimana", "giorno_settimana", placeholder="Lunedi"),
                input_field("Orario", "orario", placeholder="17:00-18:00"),
                textarea_field("Descrizione", "descrizione"),
            ]
        ),
        "Salva corso",
        hidden_fields=work_year_query(query_params),
    )

    create_enrollment_form = form_card(
        "Iscrizione corso",
        "Collega un associato a un corso e definisci la quota mensile da applicare.",
        "/azioni/corsi/iscrizione",
        "".join(
            [
                select_field("Associato", "associato_id", render_associato_options(associati), required_field=True, wide=True, searchable=True),
                select_field(
                    "Corso",
                    "corso_id",
                    render_select_options(corsi, data_keys=["importo"]),
                    required_field=True,
                    wide=True,
                    element_id="corso-iscrizione-select",
                    attrs={"data-amount-target": "corso-iscrizione-quota"},
                ),
                input_field("Data iscrizione", "data_iscrizione", input_type="date", value=date.today().isoformat(), required_field=True),
                input_field("Data inizio", "data_inizio", input_type="date"),
                input_field(
                    "Quota mensile",
                    "quota_mensile",
                    input_type="number",
                    step="0.01",
                    minimum="0",
                    element_id="corso-iscrizione-quota",
                ),
                select_field(
                    "Stato iscrizione",
                    "stato_iscrizione",
                    corso_enrollment_status_options("Attiva"),
                ),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva iscrizione corso",
        hidden_fields=work_year_query(query_params),
        form_attrs={
            "data-payment-flow": "corso",
            "data-payment-amount-field": "quota_mensile",
            "data-payment-method-default": metodo_predefinito,
            "data-payment-prompt-title": "Conferma iscrizione corso",
            "data-payment-prompt-message": "Dopo il salvataggio verra generata automaticamente la quota mensile del mese di iscrizione. Vuoi procedere subito anche al pagamento?",
            "data-payment-prompt-yes": "Si, procedi al pagamento",
            "data-payment-prompt-no": "No, solo iscrizione",
            "data-payment-dialog-title": "Pagamento quote corso",
            "data-payment-dialog-message": "Scegli se pagare solo il mese di iscrizione oppure anche le mensilita future.",
            "data-payment-dialog-confirm": "Registra pagamento e genera ricevuta",
        },
    )

    forms_html = f"""
    <div class="cards-grid cards-stack">
      {create_enrollment_form}
      {create_course_form}
    </div>
    """ if not data_only else ""

    tables_html = f"""
    {table_card(
        "Corsi registrati",
        "Catalogo corsi con quota mensile e organizzazione.",
        corsi_table,
        [
            ("numero_progressivo", "N."),
            ("codice_corso", "Codice"),
            ("corso", "Corso"),
            ("quota_mensile_standard", "Mensile", lambda value, _: money(value)),
            ("giorno_settimana", "Giorno"),
            ("orario", "Orario"),
            ("attivo", "Attivo"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("corsi", value, query_params),
                    delete_action=f"/azioni/crud/elimina/corsi/{value}",
                    delete_prompt="Eliminare questo corso? Le iscrizioni collegate devono essere eliminate prima.",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        table_class="search-table",
    )}
    {table_card(
        "Iscrizioni corsi",
        "Elenco delle iscrizioni corsi con quota mensile e stato.",
        iscrizioni_table,
        [
            ("associato", "Associato"),
            ("corso", "Corso"),
            ("data_iscrizione", "Data iscrizione"),
            ("data_inizio", "Inizio"),
            ("data_fine", "Fine"),
            ("quota_mensile", "Quota mensile", lambda value, _: money(value)),
            ("stato_iscrizione", "Stato"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("iscrizioni_corsi", value, query_params),
                    delete_action=f"/azioni/crud/elimina/iscrizioni_corsi/{value}",
                    delete_prompt="Eliminare questa iscrizione corso e tutti i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        table_class="search-table",
        summary_rows=summary_rows_for_table(iscrizioni_table, [
            ("associato", "Associato"),
            ("corso", "Corso"),
            ("data_iscrizione", "Data iscrizione"),
            ("data_inizio", "Inizio"),
            ("data_fine", "Fine"),
            ("quota_mensile", "Quota mensile", lambda value, _: money(value)),
            ("stato_iscrizione", "Stato"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("iscrizioni_corsi", value, query_params),
                    delete_action=f"/azioni/crud/elimina/iscrizioni_corsi/{value}",
                    delete_prompt="Eliminare questa iscrizione corso e tutti i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ]),
    )}
    {table_card(
        f"Quote mensili corsi anno {work_year}",
        "Vista operativa per associato, competenza e stato pagamento.",
        rate_table,
        [
            ("corso", "Corso"),
            ("associato", "Associato"),
            ("competenza", "Competenza"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("rate_corsi_mensili", value, query_params),
                    delete_action=f"/azioni/crud/elimina/rate_corsi_mensili/{value}",
                    delete_prompt="Eliminare questa quota mensile e i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        empty_message="Nessuna quota mensile presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(rate_table, [
            ("corso", "Corso"),
            ("associato", "Associato"),
            ("competenza", "Competenza"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("rate_corsi_mensili", value, query_params),
                    delete_action=f"/azioni/crud/elimina/rate_corsi_mensili/{value}",
                    delete_prompt="Eliminare questa quota mensile e i pagamenti collegati?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ]),
    )}
    {table_card(
        f"Pagamenti quote mensili anno {work_year}",
        "Incassi registrati sulle quote mensili dei corsi.",
        rate_payments_table,
        [
            ("associato", "Associato"),
            ("corso", "Corso"),
            ("competenza", "Competenza"),
            ("data_pagamento", "Data pagamento"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("metodo_pagamento", "Metodo"),
            ("riferimento", "Riferimento"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("pagamenti_rate_corsi", value, query_params),
                    extra_links=[(
                        grouped_receipt_link(row["gruppo_ricevuta"], query_params)
                        if row["gruppo_ricevuta"]
                        else receipt_link("corsi-rata", value, query_params),
                        "Ricevuta",
                    )],
                    delete_action=f"/azioni/crud/elimina/pagamenti_rate_corsi/{value}",
                    delete_prompt="Eliminare questo pagamento della quota mensile?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ],
        empty_message="Nessun pagamento mensile presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(rate_payments_table, [
            ("associato", "Associato"),
            ("corso", "Corso"),
            ("competenza", "Competenza"),
            ("data_pagamento", "Data pagamento"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("metodo_pagamento", "Metodo"),
            ("riferimento", "Riferimento"),
            (
                "id",
                "Azioni",
                lambda value, row: action_links_html(
                    edit_href=edit_path("pagamenti_rate_corsi", value, query_params),
                    extra_links=[(
                        grouped_receipt_link(row["gruppo_ricevuta"], query_params)
                        if row["gruppo_ricevuta"]
                        else receipt_link("corsi-rata", value, query_params),
                        "Ricevuta",
                    )],
                    delete_action=f"/azioni/crud/elimina/pagamenti_rate_corsi/{value}",
                    delete_prompt="Eliminare questo pagamento della quota mensile?",
                    extra_fields=work_year_query(query_params),
                ),
            ),
        ]),
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/corsi", query_params, "Apri dati corsi")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati corsi" if data_only else "Corsi", "/maschere/corsi", content, query_params, current_user)


def campi_estivi_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    work_year = current_work_year(query_params)
    associati = associati_options()
    iscrizioni = iscrizioni_campi_aperte_options(work_year)
    metodi = metodi_options()
    metodo_predefinito = preferred_metodo_pagamento_id(metodi)
    quote_campo = quote_predefinite_options("campi-estivi")
    quote_campo_table = quote_predefinite_rows("campi-estivi")
    iscrizioni_table = fetch_all(
        f"""
        SELECT
            ice.id,
            ce.anno,
            {associato_display_sql('a')} AS associato,
            ice.data_iscrizione,
            ice.quota_partecipazione,
            ice.stato_iscrizione
        FROM iscrizioni_campi_estivi ice
        JOIN associati a ON a.id = ice.associato_id
        JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
        WHERE ce.anno = ?
        ORDER BY ice.id DESC
        """,
        (work_year,),
    )
    pagamenti_table = fetch_all(
        f"""
        SELECT
            pce.id,
            ce.anno,
            {associato_display_sql('a')} AS associato,
            pce.data_pagamento,
            pce.importo,
            COALESCE(mp.nome, '') AS metodo_pagamento,
            COALESCE(pce.riferimento, '') AS riferimento
        FROM pagamenti_campi_estivi pce
        JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
        JOIN associati a ON a.id = ice.associato_id
        JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id
        WHERE ce.anno = ?
        ORDER BY pce.data_pagamento DESC, associato
        """,
        (work_year,),
    )
    saldo = fetch_all(
        """
        SELECT anno, associato, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
        FROM v_campi_estivi_saldo
        WHERE anno = ?
        ORDER BY associato
        """,
        (work_year,),
    )

    quote_form = form_card(
        "Nuova quota Campo estivo",
        "Inserisci una quota predefinita con descrizione e importo da richiamare in fase di iscrizione.",
        "/azioni/quote/crea",
        "".join(
            [
                input_field("Descrizione", "descrizione", required_field=True, wide=True),
                input_field("Importo", "importo", input_type="number", step="0.01", minimum="0", required_field=True),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva quota",
        hidden_fields={**work_year_query(query_params), "area": "campi-estivi"},
    )

    enrollment_form = form_card(
        "Iscrizione Campo estivo",
        "Registra il partecipante e proponi la quota scelta, sempre modificabile.",
        "/azioni/campi-estivi/iscrizione",
        "".join(
            [
                readonly_field("Anno di lavoro", str(work_year)),
                select_field("Associato", "associato_id", render_associato_options(associati), required_field=True, wide=True, searchable=True),
                select_field(
                    "Quota predefinita",
                    "quota_predefinita_id",
                    render_select_options(quote_campo, data_keys=["importo"]),
                    wide=True,
                    element_id="campo-quota-select",
                    attrs={"data-amount-target": "campo-estivo-quota"},
                ),
                input_field("Data iscrizione", "data_iscrizione", input_type="date", value=date.today().isoformat(), required_field=True),
                input_field(
                    "Importo",
                    "quota_partecipazione",
                    input_type="number",
                    step="0.01",
                    minimum="0",
                    required_field=True,
                    element_id="campo-estivo-quota",
                    wide=True,
                ),
                select_field(
                    "Stato iscrizione",
                    "stato_iscrizione",
                    camp_enrollment_status_options("Iscritto"),
                ),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva iscrizione Campo estivo",
        hidden_fields=work_year_query(query_params),
        form_attrs={
            "data-payment-flow": "campo-estivo",
            "data-payment-amount-field": "quota_partecipazione",
            "data-payment-method-default": metodo_predefinito,
            "data-payment-prompt-title": "Conferma iscrizione Campo estivo",
            "data-payment-prompt-message": "Vuoi procedere subito anche al pagamento dell'iscrizione al Campo estivo?",
            "data-payment-prompt-yes": "Si, registra anche il pagamento",
            "data-payment-prompt-no": "No, solo iscrizione",
            "data-payment-dialog-title": "Pagamento Campo estivo",
            "data-payment-dialog-message": "Conferma i dati del pagamento del Campo estivo.",
            "data-payment-dialog-confirm": "Registra pagamento e genera ricevuta",
        },
    )

    forms_html = f"""
    <div class="cards-grid cards-stack">
      {enrollment_form}
      {quote_form}
    </div>
    """ if not data_only else ""

    quote_campo_columns = [
        ("descrizione", "Descrizione"),
        ("importo", "Importo", lambda value, _: money(value)),
        ("attiva", "Attiva"),
        ("note", "Note"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("quote_campi_estivi", value, query_params),
                delete_action=f"/azioni/crud/elimina/quote_campi_estivi/{value}",
                delete_prompt="Eliminare questa quota predefinita di Campo estivo?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    iscrizioni_columns = [
        ("anno", "Anno"),
        ("associato", "Associato"),
        ("data_iscrizione", "Data iscrizione"),
        ("quota_partecipazione", "Quota", lambda value, _: money(value)),
        ("stato_iscrizione", "Stato"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("iscrizioni_campi_estivi", value, query_params),
                delete_action=f"/azioni/crud/elimina/iscrizioni_campi_estivi/{value}",
                delete_prompt="Eliminare questa iscrizione al Campo estivo e il pagamento collegato?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    pagamenti_columns = [
        ("anno", "Anno"),
        ("associato", "Associato"),
        ("data_pagamento", "Data"),
        ("importo", "Importo", lambda value, _: money(value)),
        ("metodo_pagamento", "Metodo"),
        ("riferimento", "Riferimento"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("pagamenti_campi_estivi", value, query_params),
                extra_links=[(receipt_link("campi-estivi", value, query_params), "Ricevuta")],
                delete_action=f"/azioni/crud/elimina/pagamenti_campi_estivi/{value}",
                delete_prompt="Eliminare questo pagamento del Campo estivo?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    saldo_columns = [
        ("anno", "Anno"),
        ("associato", "Associato"),
        ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
        ("importo_pagato", "Pagato", lambda value, _: money(value)),
        ("saldo_residuo", "Residuo", lambda value, _: money(value)),
        ("stato_pagamento", "Stato"),
    ]

    tables_html = f"""
    {table_card(
        "Iscrizioni Campo estivo",
        "Partecipanti registrati al Campo estivo dell'anno selezionato.",
        iscrizioni_table,
        iscrizioni_columns,
        table_class="search-table",
        summary_rows=summary_rows_for_table(iscrizioni_table, iscrizioni_columns),
    )}
    {table_card(
        "Pagamenti Campo estivo",
        "Pagamenti una tantum registrati per il Campo estivo.",
        pagamenti_table,
        pagamenti_columns,
        table_class="search-table",
        summary_rows=summary_rows_for_table(pagamenti_table, pagamenti_columns),
    )}
    {table_card(
        f"Situazione Campo estivo anno {work_year}",
        "Elenco iscritti con saldo della quota di partecipazione.",
        saldo,
        saldo_columns,
        empty_message="Nessun movimento presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(saldo, saldo_columns),
    )}
    {table_card(
        "Quote Campo estivo",
        "Elenco delle quote predefinite disponibili per le iscrizioni al Campo estivo.",
        quote_campo_table,
        quote_campo_columns,
        empty_message="Nessuna quota predefinita presente.",
        table_class="search-table",
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/campi-estivi", query_params, "Apri dati Campo estivo")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati Campo estivo" if data_only else ESTATE_LABEL, "/maschere/campi-estivi", content, query_params, current_user)


def eventi_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    work_year = current_work_year(query_params)
    associati = associati_options()
    eventi = eventi_options(work_year)
    iscrizioni = iscrizioni_eventi_aperte_options(work_year)
    metodi = metodi_options()
    metodo_predefinito = preferred_metodo_pagamento_id(metodi)

    eventi_table = fetch_all(
        """
        SELECT
            id,
            numero_progressivo,
            codice_evento,
            nome,
            tipologia,
            data_evento,
            quota_partecipazione_standard,
            attivo
        FROM eventi
        WHERE substr(COALESCE(data_evento, ''), 1, 4) = ?
        ORDER BY data_evento DESC, nome
        """,
        (str(work_year),),
    )
    iscrizioni_table = fetch_all(
        f"""
        SELECT
            ie.id,
            e.nome AS evento,
            {associato_display_sql('a')} AS associato,
            ie.data_iscrizione,
            ie.quota_partecipazione,
            ie.stato_iscrizione
        FROM iscrizioni_eventi ie
        JOIN associati a ON a.id = ie.associato_id
        JOIN eventi e ON e.id = ie.evento_id
        WHERE substr(COALESCE(e.data_evento, ''), 1, 4) = ?
        ORDER BY ie.id DESC
        """,
        (str(work_year),),
    )
    pagamenti_table = fetch_all(
        f"""
        SELECT
            pe.id,
            e.nome AS evento,
            {associato_display_sql('a')} AS associato,
            pe.data_pagamento,
            pe.importo,
            COALESCE(mp.nome, '') AS metodo_pagamento,
            COALESCE(pe.riferimento, '') AS riferimento
        FROM pagamenti_eventi pe
        JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
        JOIN associati a ON a.id = ie.associato_id
        JOIN eventi e ON e.id = ie.evento_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id
        WHERE substr(COALESCE(e.data_evento, ''), 1, 4) = ?
        ORDER BY pe.data_pagamento DESC, associato
        """,
        (str(work_year),),
    )
    saldo = fetch_all(
        """
        SELECT evento, tipologia, data_evento, associato, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
        FROM v_eventi_saldo
        WHERE substr(data_evento, 1, 4) = ?
        ORDER BY data_evento DESC, evento, associato
        """,
        (str(work_year),),
    )

    create_event_form = form_card(
        "Nuovo evento",
        "Crea eventi diversi con quota di partecipazione una tantum.",
        "/azioni/eventi/crea",
        "".join(
            [
                readonly_field("Codice evento assegnato", peek_next_progressive_code("eventi")),
                input_field("Nome evento", "nome", required_field=True),
                input_field("Data evento", "data_evento", input_type="date", required_field=True),
                input_field("Luogo", "luogo"),
                input_field("Quota dovuta", "quota_partecipazione_standard", input_type="number", step="0.01", minimum="0"),
                textarea_field("Descrizione", "descrizione"),
            ]
        ),
        "Salva evento",
        hidden_fields=work_year_query(query_params),
    )

    enrollment_form = form_card(
        "Iscrizione evento",
        "Registra i partecipanti all'evento e la loro quota.",
        "/azioni/eventi/iscrizione",
        "".join(
            [
                select_field("Associato", "associato_id", render_associato_options(associati), required_field=True, wide=True, searchable=True),
                select_field(
                    "Evento",
                    "evento_id",
                    render_select_options(eventi, data_keys=["importo"]),
                    required_field=True,
                    wide=True,
                    element_id="evento-iscrizione-select",
                    attrs={"data-amount-target": "evento-iscrizione-quota"},
                ),
                input_field("Data iscrizione", "data_iscrizione", input_type="date", value=date.today().isoformat(), required_field=True),
                input_field(
                    "Quota dovuta",
                    "quota_partecipazione",
                    input_type="number",
                    step="0.01",
                    minimum="0",
                    required_field=True,
                    element_id="evento-iscrizione-quota",
                ),
                select_field(
                    "Stato iscrizione",
                    "stato_iscrizione",
                    event_enrollment_status_options("Iscritto"),
                ),
                textarea_field("Note", "note"),
            ]
        ),
        "Salva iscrizione evento",
        hidden_fields=work_year_query(query_params),
        form_attrs={
            "data-payment-flow": "evento",
            "data-payment-amount-field": "quota_partecipazione",
            "data-payment-method-default": metodo_predefinito,
            "data-payment-prompt-title": "Conferma iscrizione evento",
            "data-payment-prompt-message": "Vuoi procedere subito anche al pagamento dell'iscrizione all'evento?",
            "data-payment-prompt-yes": "Si, registra anche il pagamento",
            "data-payment-prompt-no": "No, solo iscrizione",
            "data-payment-dialog-title": "Pagamento evento",
            "data-payment-dialog-message": "Conferma i dati del pagamento dell'evento.",
            "data-payment-dialog-confirm": "Registra pagamento e genera ricevuta",
        },
    )

    forms_html = f"""
    <div class="cards-grid cards-stack">
      {enrollment_form}
      {create_event_form}
    </div>
    """ if not data_only else ""

    eventi_columns = [
        ("numero_progressivo", "N."),
        ("codice_evento", "Codice"),
        ("nome", "Evento"),
        ("data_evento", "Data"),
        ("quota_partecipazione_standard", "Quota dovuta", lambda value, _: money(value)),
        ("attivo", "Attivo"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("eventi", value, query_params),
                delete_action=f"/azioni/crud/elimina/eventi/{value}",
                delete_prompt="Eliminare questo evento e tutte le iscrizioni collegate?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    iscrizioni_columns = [
        ("evento", "Evento"),
        ("associato", "Associato"),
        ("data_iscrizione", "Data iscrizione"),
        ("quota_partecipazione", "Quota dovuta", lambda value, _: money(value)),
        ("stato_iscrizione", "Stato"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("iscrizioni_eventi", value, query_params),
                delete_action=f"/azioni/crud/elimina/iscrizioni_eventi/{value}",
                delete_prompt="Eliminare questa iscrizione evento e il pagamento collegato?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    pagamenti_columns = [
        ("evento", "Evento"),
        ("associato", "Associato"),
        ("data_pagamento", "Data"),
        ("importo", "Importo", lambda value, _: money(value)),
        ("metodo_pagamento", "Metodo"),
        ("riferimento", "Riferimento"),
        (
            "id",
            "Azioni",
            lambda value, row: action_links_html(
                edit_href=edit_path("pagamenti_eventi", value, query_params),
                extra_links=[(receipt_link("eventi", value, query_params), "Ricevuta")],
                delete_action=f"/azioni/crud/elimina/pagamenti_eventi/{value}",
                delete_prompt="Eliminare questo pagamento evento?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]
    saldo_columns = [
        ("evento", "Evento"),
        ("data_evento", "Data"),
        ("associato", "Associato"),
        ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
        ("importo_pagato", "Pagato", lambda value, _: money(value)),
        ("saldo_residuo", "Residuo", lambda value, _: money(value)),
        ("stato_pagamento", "Stato"),
    ]

    tables_html = f"""
    {table_card(
        "Eventi registrati",
        "Anagrafica completa degli eventi.",
        eventi_table,
        eventi_columns,
        table_class="search-table",
        summary_rows=summary_rows_for_table(eventi_table, eventi_columns),
    )}
    {table_card(
        "Iscrizioni eventi",
        "Partecipanti registrati agli eventi.",
        iscrizioni_table,
        iscrizioni_columns,
        table_class="search-table",
        summary_rows=summary_rows_for_table(iscrizioni_table, iscrizioni_columns),
    )}
    {table_card(
        "Pagamenti eventi",
        "Pagamenti una tantum registrati per gli eventi.",
        pagamenti_table,
        pagamenti_columns,
        table_class="search-table",
        summary_rows=summary_rows_for_table(pagamenti_table, pagamenti_columns),
    )}
    {table_card(
        f"Situazione eventi anno {work_year}",
        "Partecipanti, quota una tantum e stato del pagamento.",
        saldo,
        saldo_columns,
        empty_message="Nessun movimento presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(saldo, saldo_columns),
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/eventi", query_params, "Apri dati eventi")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati eventi" if data_only else "Eventi", "/maschere/eventi", content, query_params, current_user)


def pagamenti_multi_area_page(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    data_only = normalized(query_params, "vista", "") == "dati"
    work_year = current_work_year(query_params)
    associati = associati_options()
    metodi = metodi_options()
    metodo_predefinito = preferred_metodo_pagamento_id(metodi)
    scadenze = scadenze_multi_area_options(work_year)
    future_course_options = pagamenti_multi_area_course_options(work_year)
    ricevute_table = multi_area_receipts_rows(work_year)
    default_until_month = (
        f"{work_year}-{date.today().month:02d}"
        if work_year == date.today().year
        else f"{work_year}-12"
    )

    payment_form = form_card(
        "Pagamento multi-area",
        "Seleziona un associato e registra in un unico pagamento scadenze aperte di tesseramenti, quote mensili corsi, Campo estivo ed eventi.",
        "/azioni/pagamenti-multi-area/crea",
        "".join(
            [
                readonly_field("Anno di lavoro", str(work_year)),
                select_field(
                    "Associato",
                    "associato_id",
                    render_associato_options(associati),
                    required_field=True,
                    wide=True,
                    searchable=True,
                    element_id="multi-area-associato",
                ),
                select_field(
                    "Corso per quote future",
                    "multi_area_iscrizione_corso_id",
                    render_select_options(future_course_options, data_keys=["associato_id", "start_competenza"]),
                    wide=True,
                    element_id="multi-area-course-enrollment",
                    attrs={"data-work-year": str(work_year)},
                ),
                input_field(
                    "Genera fino alla competenza",
                    "multi_area_competenza_fine",
                    input_type="month",
                    value=default_until_month,
                    wide=True,
                    element_id="multi-area-course-until",
                ),
                (
                    '<div class="field wide multi-area-course-actions">'
                    '<span>Quote future corso</span>'
                    '<div class="multi-area-course-actions-row">'
                    '<button type="button" class="button secondary" id="multi-area-course-generate">Aggiungi quote future corso</button>'
                    "</div>"
                    '<p id="multi-area-course-feedback" class="multi-area-course-feedback">'
                    "Seleziona associato e corso, indica fino a quale mensilita vuoi arrivare e il gestionale selezionera in basso le relative quote."
                    "</p>"
                    "</div>"
                ),
                multi_select_field(
                    "Scadenze da saldare",
                    "scadenza_id",
                    render_select_options_multi(scadenze, data_keys=["associato_id", "residuo"]),
                    required_field=True,
                    wide=True,
                    size=12,
                    element_id="multi-area-scadenze",
                ),
                input_field("Data pagamento", "data_pagamento", input_type="date", value=date.today().isoformat(), required_field=True),
                input_field(
                    "Importo da registrare",
                    "importo",
                    input_type="number",
                    step="0.01",
                    minimum="0.01",
                    required_field=True,
                    element_id="multi-area-amount",
                ),
                select_field("Metodo", "metodo_pagamento_id", render_select_options(metodi, metodo_predefinito), required_field=True),
                input_field("Riferimento", "riferimento"),
                textarea_field("Note", "note"),
            ]
        ),
        "Registra pagamento e ricevuta unica",
        hidden_fields=work_year_query(query_params),
    )

    forms_html = f"""
    <div class="cards-grid">
      {payment_form}
    </div>
    """ if not data_only else ""

    ricevute_columns = [
        ("gruppo_ricevuta", "Codice ricevuta"),
        ("associato", "Associato"),
        ("data_pagamento", "Data pagamento"),
        ("numero_scadenze", "Scadenze"),
        ("importo_totale", "Importo", lambda value, _: money(value)),
        (
            "gruppo_ricevuta",
            "Azioni",
            lambda value, row: action_links_html(
                extra_links=[(with_query(f"/ricevute/multi-area-gruppo/{value}", work_year_query(query_params)), "Ricevuta")],
                delete_action=f"/azioni/pagamenti-multi-area/elimina/{value}",
                delete_prompt="Eliminare questo pagamento multi-area e ripristinare le scadenze saldate?",
                extra_fields=work_year_query(query_params),
            ),
        ),
    ]

    tables_html = f"""
    {table_card(
        f"Ricevute multi-area anno {work_year}",
        "Storico dei pagamenti unici registrati su scadenze provenienti da aree diverse.",
        ricevute_table,
        ricevute_columns,
        empty_message="Nessuna ricevuta multi-area presente per l'anno selezionato.",
        table_class="search-table",
        summary_rows=summary_rows_for_table(ricevute_table, ricevute_columns),
    )}
    """

    content = f"""
    {view_mode_switch("/maschere/pagamenti-multi-area", query_params, "Apri dati pagamenti")}
    {forms_html}
    {data_view_search_toolbar() if data_only else ""}
    {tables_html if data_only else ""}
    """
    return page("Dati pagamenti" if data_only else "Pagamenti", "/maschere/pagamenti-multi-area", content, query_params, current_user)


def with_query(path: str, query_params: dict[str, str]) -> str:
    active_params = {key: value for key, value in query_params.items() if value}
    if not active_params:
        return path
    return f"{path}?{urlencode(active_params)}"


def slugify(value: str) -> str:
    cleaned = []
    previous_dash = False
    for character in value.lower():
        if character.isalnum():
            cleaned.append(character)
            previous_dash = False
        elif not previous_dash:
            cleaned.append("-")
            previous_dash = True
    return "".join(cleaned).strip("-") or "report"


def report_display_value(row: sqlite3.Row, column: tuple) -> str:
    key = column[0]
    formatter = column[2] if len(column) > 2 else None
    value = row[key]
    if callable(formatter):
        return plain_text(formatter(value, row))
    return plain_text(value)


def report_share_detail_columns(columns: list[tuple]) -> list[tuple]:
    excluded_labels = {"Azioni", "Ricevuta"}
    excluded_keys = {"id", "payment_id"}
    return [
        column
        for column in columns
        if len(column) >= 2
        and column[1] not in excluded_labels
        and column[0] not in excluded_keys
    ]


def report_share_row_line(row: sqlite3.Row, columns: list[tuple], index: int) -> str:
    parts: list[str] = []
    for column in report_share_detail_columns(columns):
        value = report_display_value(row, column)
        if not value:
            continue
        parts.append(f"{column[1]}: {value}")
    if not parts:
        return f"{index}. Riga {index}"
    return f"{index}. " + " | ".join(parts)


def report_search_term(query_params: dict[str, str]) -> str:
    return normalized(query_params, "search", "")


def definition_with_search(definition: dict, query_params: dict[str, str]) -> dict:
    search_term = report_search_term(query_params)
    if not search_term:
        return definition
    filters = list(definition.get("filters", []))
    filters.append({"label": "Ricerca testo", "value": search_term})
    enriched = dict(definition)
    enriched["filters"] = filters
    return enriched


def filter_report_rows(rows: list[sqlite3.Row], columns: list[tuple], search_term: str) -> list[sqlite3.Row]:
    needle = search_term.lower().strip()
    if not needle:
        return rows
    return [
        row
        for row in rows
        if needle in " ".join(report_display_value(row, column).lower() for column in columns)
    ]


def participants_target_options(area: str, work_year: int) -> list[sqlite3.Row]:
    if area == "corsi":
        return fetch_all(
            f"""
            SELECT c.id, c.nome AS label
            FROM corsi c
            WHERE {corso_year_relevance_sql('c')}
            ORDER BY c.nome
            """,
            corso_year_relevance_params(work_year),
        )
    if area == "campi-estivi":
        return fetch_all(
            """
            SELECT id, 'Campo estivo ' || anno AS label
            FROM campi_estivi
            WHERE anno = ?
            ORDER BY id
            """,
            (work_year,),
        )
    if area == "eventi":
        return fetch_all(
            """
            SELECT id, nome || ' - ' || data_evento AS label
            FROM eventi
            WHERE substr(data_evento, 1, 4) = ?
            ORDER BY data_evento, nome
            """,
            (str(work_year),),
        )
    return []


def build_participants_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    area = normalized(query_params, "area", "corsi")
    target_id = normalized(query_params, "target_id", "")
    area_options = render_static_options(
        [("corsi", "Corso"), ("campi-estivi", ESTATE_LABEL), ("eventi", "Evento")],
        area,
        blank_label=None,
    )
    target_rows = participants_target_options(area, work_year)
    target_options = render_select_options(target_rows, target_id, "Seleziona...")
    current_query = dict(work_year_query(query_params))
    current_query["area"] = area
    if target_id:
        current_query["target_id"] = target_id

    lead_html = f"""
    <section class="card compact screen-only">
      <div class="card-head">
        <h2>Filtro partecipanti</h2>
        <p>Seleziona il corso, il Campo estivo o l'evento di cui vuoi visualizzare e stampare i partecipanti.</p>
      </div>
      <form method="get" action="/report/partecipanti" class="form-grid">
        <input type="hidden" name="anno_lavoro" value="{esc(work_year)}">
        {select_field("Area", "area", area_options, required_field=True, attrs={"onchange": "const target = this.form.querySelector('[name=target_id]'); if (target) target.value = ''; this.form.submit();"})}
        {select_field("Elemento", "target_id", target_options, required_field=True, wide=True)}
        <div class="form-actions">
          <button type="submit" class="button">Apri partecipanti</button>
        </div>
      </form>
    </section>
    """

    base_definition = {
        "title": "Partecipanti attività",
        "current_path": "/report/partecipanti",
        "subtitle": "Visualizzazione e stampa partecipanti per corso, Campo estivo o evento.",
        "query": """
            SELECT '' AS codice_associato, '' AS associato, '' AS telefono, '' AS email, '' AS data_iscrizione, '' AS stato_iscrizione, 0 AS quota
            WHERE 1 = 0
        """,
        "params": (),
        "sheet_name": "Partecipanti attivita",
        "export_name": "partecipanti.xlsx",
        "lead_html": lead_html,
        "filters": [{"label": "Anno di lavoro", "value": str(work_year)}],
        "columns": [
            ("codice_associato", "Codice"),
            ("associato", "Associato"),
            ("telefono", "Telefono"),
            ("email", "Email"),
            ("data_iscrizione", "Data iscrizione"),
            ("stato_iscrizione", "Stato"),
            ("quota", "Quota", lambda value, _: money(value)),
        ],
    }

    if not target_id.isdigit():
        return base_definition

    if area == "corsi":
        course = fetch_one(
            f"""
            SELECT nome
            FROM corsi c
            WHERE c.id = ?
              AND {corso_year_relevance_sql('c')}
            """,
            (int(target_id), *corso_year_relevance_params(work_year)),
        )
        if course is None:
            return base_definition
        base_definition.update(
            {
                "subtitle": f"Partecipanti iscritti al corso {course['nome']}.",
                "query": f"""
                    SELECT
                        a.codice_associato,
                        {associato_display_sql('a')} AS associato,
                        COALESCE(a.telefono, '') AS telefono,
                        COALESCE(a.email, '') AS email,
                        ic.data_iscrizione,
                        ic.stato_iscrizione,
                        ic.quota_mensile AS quota
                    FROM iscrizioni_corsi ic
                    JOIN associati a ON a.id = ic.associato_id
                    WHERE ic.corso_id = ?
                      AND {iscrizione_corso_year_relevance_sql('ic')}
                    ORDER BY a.cognome, a.nome
                """,
                "params": (int(target_id), *iscrizione_corso_year_relevance_params(work_year)),
                "export_name": f"partecipanti_corso_{slugify(course['nome'])}.xlsx",
                "filters": [
                    {"label": "Anno di lavoro", "value": str(work_year)},
                    {"label": "Corso", "value": course["nome"]},
                ],
            }
        )
        return base_definition

    if area == "campi-estivi":
        camp = fetch_one(
            """
            SELECT nome, anno
            FROM campi_estivi
            WHERE id = ?
              AND anno = ?
            """,
            (int(target_id), work_year),
        )
        if camp is None:
            return base_definition
        base_definition.update(
            {
                "subtitle": f"Partecipanti iscritti al Campo estivo {camp['anno']}.",
                "query": f"""
                    SELECT
                        a.codice_associato,
                        {associato_display_sql('a')} AS associato,
                        COALESCE(a.telefono, '') AS telefono,
                        COALESCE(a.email, '') AS email,
                        ice.data_iscrizione,
                        ice.stato_iscrizione,
                        ice.quota_partecipazione AS quota
                    FROM iscrizioni_campi_estivi ice
                    JOIN associati a ON a.id = ice.associato_id
                    WHERE ice.campo_estivo_id = ?
                    ORDER BY a.cognome, a.nome
                """,
                "params": (int(target_id),),
                "export_name": f"partecipanti_campo_estivo_{camp['anno']}.xlsx",
                "filters": [
                    {"label": "Anno di lavoro", "value": str(work_year)},
                    {"label": ESTATE_LABEL, "value": f"{ESTATE_LABEL} {camp['anno']}"},
                ],
            }
        )
        return base_definition

    if area == "eventi":
        event = fetch_one(
            """
            SELECT nome, data_evento
            FROM eventi
            WHERE id = ?
              AND substr(COALESCE(data_evento, ''), 1, 4) = ?
            """,
            (int(target_id), str(work_year)),
        )
        if event is None:
            return base_definition
        base_definition.update(
            {
                "subtitle": f"Partecipanti iscritti all'evento {event['nome']}.",
                "query": f"""
                    SELECT
                        a.codice_associato,
                        {associato_display_sql('a')} AS associato,
                        COALESCE(a.telefono, '') AS telefono,
                        COALESCE(a.email, '') AS email,
                        ie.data_iscrizione,
                        ie.stato_iscrizione,
                        ie.quota_partecipazione AS quota
                    FROM iscrizioni_eventi ie
                    JOIN associati a ON a.id = ie.associato_id
                    JOIN eventi e ON e.id = ie.evento_id
                    WHERE ie.evento_id = ?
                      AND substr(COALESCE(e.data_evento, ''), 1, 4) = ?
                    ORDER BY a.cognome, a.nome
                """,
                "params": (int(target_id), str(work_year)),
                "export_name": f"partecipanti_evento_{slugify(event['nome'])}.xlsx",
                "filters": [
                    {"label": "Anno di lavoro", "value": str(work_year)},
                    {"label": "Evento", "value": f"{event['nome']} ({event['data_evento']})"},
                ],
            }
        )
    return base_definition


def incassi_dataset_sql() -> str:
    return f"""
        SELECT
            'Tesseramento annuale' AS area,
            t.anno_sociale AS anno_riferimento,
            pt.data_pagamento,
            pt.importo,
            mp.nome AS metodo_pagamento,
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            'Anno ' || t.anno_sociale AS riferimento,
            'tesseramenti' AS payment_type,
            pt.id AS payment_id,
            '' AS gruppo_ricevuta
        FROM pagamenti_tesseramenti pt
        JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
        JOIN associati a ON a.id = t.associato_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id

        UNION ALL

        SELECT
            'Corso - quota mensile' AS area,
            r.anno AS anno_riferimento,
            prc.data_pagamento,
            prc.importo,
            mp.nome AS metodo_pagamento,
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS riferimento,
            'corsi-rata' AS payment_type,
            prc.id AS payment_id,
            COALESCE(prc.gruppo_ricevuta, '') AS gruppo_ricevuta
        FROM pagamenti_rate_corsi prc
        JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
        JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id

        UNION ALL

        SELECT
            'Campo estivo' AS area,
            ce.anno AS anno_riferimento,
            pce.data_pagamento,
            pce.importo,
            mp.nome AS metodo_pagamento,
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            'Campo estivo ' || ce.anno AS riferimento,
            'campi-estivi' AS payment_type,
            pce.id AS payment_id,
            '' AS gruppo_ricevuta
        FROM pagamenti_campi_estivi pce
        JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
        JOIN associati a ON a.id = ice.associato_id
        JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id

        UNION ALL

        SELECT
            'Evento' AS area,
            CAST(substr(COALESCE(e.data_evento, ''), 1, 4) AS INTEGER) AS anno_riferimento,
            pe.data_pagamento,
            pe.importo,
            mp.nome AS metodo_pagamento,
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            e.nome AS riferimento,
            'eventi' AS payment_type,
            pe.id AS payment_id,
            '' AS gruppo_ricevuta
        FROM pagamenti_eventi pe
        JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
        JOIN associati a ON a.id = ie.associato_id
        JOIN eventi e ON e.id = ie.evento_id
        LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id
    """


def lookup_label(table_name: str, record_id: str, label_query: str, default: str = "") -> str:
    if not record_id.isdigit():
        return default
    row = fetch_one(label_query, (int(record_id),))
    if row is None:
        return default
    return row[0]


def report_link(label: str, path: str, query_params: dict[str, str]) -> str:
    return f'<a href="{esc(with_query(path, query_params))}">{esc(label)}</a>'


def format_money_plain(value: object) -> str:
    return money(value).replace(" EUR", "")


def chart_card(title: str, subtitle: str, chart_html: str, legend_html: str = "") -> str:
    return f"""
    <section class="card chart-card">
      <div class="card-head">
        <h2>{esc(title)}</h2>
        <p>{esc(subtitle)}</p>
      </div>
      <div class="chart-wrap">
        {chart_html}
        {legend_html}
      </div>
    </section>
    """


def pie_chart_svg(paid_total: float, due_total: float, incassi_url: str, scadenze_url: str) -> str:
    total = max(paid_total + due_total, 0.0)
    if total <= 0:
        return '<div class="empty-state">Nessun movimento disponibile per il grafico.</div>'

    circumference = 2 * 3.14159 * 72
    paid_ratio = paid_total / total
    paid_dash = circumference * paid_ratio
    due_dash = max(circumference - paid_dash, 0)
    svg = f"""
    <svg class="chart-svg pie-chart" viewBox="0 0 220 220" role="img" aria-label="Grafico a torta incassi e scadenze">
      <circle cx="110" cy="110" r="72" fill="none" stroke="#ffe6ce" stroke-width="26" class="chart-clickable" ondblclick="window.location.href='{esc(scadenze_url)}'"><title>Scadenze: {esc(money(due_total))}</title></circle>
      <circle cx="110" cy="110" r="72" fill="none" stroke="#ef7f1a" stroke-width="26" stroke-linecap="round"
              stroke-dasharray="{paid_dash:.2f} {due_dash:.2f}" transform="rotate(-90 110 110)" class="chart-clickable" ondblclick="window.location.href='{esc(incassi_url)}'"><title>Incassi: {esc(money(paid_total))}</title></circle>
      <circle cx="110" cy="110" r="48" fill="#ffffff"></circle>
      <text x="110" y="92" text-anchor="middle" class="chart-number chart-clickable" ondblclick="window.location.href='{esc(incassi_url)}'">{esc(format_money_plain(paid_total))}<title>Incassi: {esc(money(paid_total))}</title></text>
      <text x="110" y="112" text-anchor="middle" class="chart-label chart-clickable" ondblclick="window.location.href='{esc(incassi_url)}'">Incassi<title>Incassi: {esc(money(paid_total))}</title></text>
      <text x="110" y="132" text-anchor="middle" class="chart-subtle chart-clickable" ondblclick="window.location.href='{esc(scadenze_url)}'">Scadenze {esc(format_money_plain(due_total))}<title>Scadenze: {esc(money(due_total))}</title></text>
    </svg>
    """
    legend = """
    <div class="chart-legend">
      <span><i class="legend-dot incassi"></i> Incassi</span>
      <span><i class="legend-dot scadenze"></i> Scadenze</span>
    </div>
    """
    return svg + legend


def bar_chart_svg(rows: list[tuple[str, float, float, str, str]]) -> str:
    if not rows:
        return '<div class="empty-state">Nessun movimento disponibile per il grafico.</div>'

    max_value = max(max(incassi, scadenze) for _, incassi, scadenze, _, _ in rows) or 1
    column_step = 120
    width = max(420, len(rows) * column_step + 40)
    height = 260
    bars = []
    labels = []
    base_y = 210
    for index, (label, incassi, scadenze, incassi_url, scadenze_url) in enumerate(rows):
        x = 40 + index * column_step
        incassi_height = (incassi / max_value) * 140 if max_value else 0
        scadenze_height = (scadenze / max_value) * 140 if max_value else 0
        bars.append(
            f'<rect x="{x}" y="{base_y - incassi_height:.1f}" width="32" height="{incassi_height:.1f}" rx="8" fill="#ef7f1a" '
            f'class="chart-clickable" ondblclick="window.location.href=\'{esc(incassi_url)}\'"><title>{esc(label)} - Incassi: {esc(money(incassi))}</title></rect>'
        )
        bars.append(
            f'<rect x="{x + 40}" y="{base_y - scadenze_height:.1f}" width="32" height="{scadenze_height:.1f}" rx="8" fill="#ffd2a8" '
            f'class="chart-clickable" ondblclick="window.location.href=\'{esc(scadenze_url)}\'"><title>{esc(label)} - Scadenze: {esc(money(scadenze))}</title></rect>'
        )
        labels.append(f'<text x="{x + 36}" y="236" text-anchor="middle" class="chart-axis">{esc(label)}<title>{esc(label)} - Totale area: {esc(money(incassi + scadenze))}</title></text>')

    total_incassi = sum(incassi for _, incassi, _, _, _ in rows)
    total_scadenze = sum(scadenze for _, _, scadenze, _, _ in rows)
    grand_total = total_incassi + total_scadenze

    return f"""
    <div class="bar-chart-wrap">
      <svg class="chart-svg bar-chart" viewBox="0 0 {width} {height}" role="img" aria-label="Grafico a barre incassi e scadenze per area">
        <line x1="24" y1="{base_y}" x2="{width - 16}" y2="{base_y}" stroke="#d8c1aa" stroke-width="1.5"></line>
        {''.join(bars)}
        {''.join(labels)}
      </svg>
      <div class="chart-legend">
        <span><i class="legend-dot incassi"></i> Incassi</span>
        <span><i class="legend-dot scadenze"></i> Scadenze</span>
      </div>
      <div class="chart-summary-strip">
        <article class="chart-summary-card incassi">
          <span>Totale incassi</span>
          <strong>{esc(money(total_incassi))}</strong>
        </article>
        <article class="chart-summary-card scadenze">
          <span>Totale scadenze</span>
          <strong>{esc(money(total_scadenze))}</strong>
        </article>
        <article class="chart-summary-card totale">
          <span>Totale complessivo</span>
          <strong>{esc(money(grand_total))}</strong>
        </article>
      </div>
    </div>
    """


def dashboard_charts(query_params: dict[str, str]) -> str:
    work_year = current_work_year(query_params)
    date_from, date_to = year_start_end(work_year)
    incassi_total_url = with_query(
        "/report/incassi",
        {"anno_lavoro": str(work_year), "date_from": date_from, "date_to": date_to},
    )
    scadenze_total_url = with_query("/report/scadenze", {"anno_lavoro": str(work_year)})
    incassi_total = float(
        fetch_scalar(
            f"""
            SELECT COALESCE(SUM(importo), 0)
            FROM ({incassi_dataset_sql()}) incassi
            WHERE anno_riferimento = ?
              AND data_pagamento BETWEEN ? AND ?
            """,
            (work_year, date_from, date_to),
        )
        or 0
    )
    scadenze_total = float(
        fetch_scalar(
            """
            SELECT COALESCE(SUM(saldo_residuo), 0)
            FROM v_scadenze_da_incassare
            WHERE substr(COALESCE(scadenza, ''), 1, 4) = ?
              AND area <> 'Corso - iscrizione'
            """,
            (str(work_year),),
        )
        or 0
    )
    area_rows = fetch_all(
        f"""
        WITH incassi AS (
            SELECT area, COALESCE(SUM(importo), 0) AS totale_incassi
            FROM ({incassi_dataset_sql()}) dati
            WHERE anno_riferimento = ?
              AND data_pagamento BETWEEN ? AND ?
            GROUP BY area
        ),
        scadenze AS (
            SELECT
                CASE WHEN area = 'Estate' THEN 'Campo estivo' ELSE area END AS area,
                COALESCE(SUM(saldo_residuo), 0) AS totale_scadenze
            FROM v_scadenze_da_incassare
            WHERE substr(COALESCE(scadenza, ''), 1, 4) = ?
              AND area <> 'Corso - iscrizione'
            GROUP BY 1
        ),
        aree AS (
            SELECT area FROM incassi
            UNION
            SELECT area FROM scadenze
        )
        SELECT
            area,
            COALESCE((SELECT totale_incassi FROM incassi WHERE incassi.area = aree.area), 0) AS totale_incassi,
            COALESCE((SELECT totale_scadenze FROM scadenze WHERE scadenze.area = aree.area), 0) AS totale_scadenze
        FROM aree
        ORDER BY area
        """,
        (work_year, date_from, date_to, str(work_year)),
    )
    chart_rows = [
        (
            row["area"],
            float(row["totale_incassi"] or 0),
            float(row["totale_scadenze"] or 0),
            with_query(
                "/report/incassi",
                {
                    "anno_lavoro": str(work_year),
                    "date_from": date_from,
                    "date_to": date_to,
                    "area": row["area"],
                },
            ),
            with_query(
                "/report/scadenze",
                {
                    "anno_lavoro": str(work_year),
                    "area": row["area"],
                },
            ),
        )
        for row in area_rows
    ]
    return f"""
    <div class="charts-grid">
      {chart_card("Incassi e Scadenze Totali", f"Anno di lavoro {work_year}. Doppio clic sui valori per aprire il report collegato.", pie_chart_svg(incassi_total, scadenze_total, incassi_total_url, scadenze_total_url))}
      {chart_card("Incassi e Scadenze per Area", "Doppio clic su ogni barra per aprire il report incassi o scadenze dell'area selezionata.", bar_chart_svg(chart_rows))}
    </div>
    """


def associato_detail_href(associato_id: object, query_params: dict[str, str]) -> str:
    return with_query(f"/report/associato/{associato_id}", work_year_query(query_params))


def report_share_message(definition: dict, rows: list[sqlite3.Row], *, max_length: int | None = None) -> str:
    filter_lines = [f"{filter_row['label']}: {filter_row['value']}" for filter_row in definition.get("filters", [])]
    summary_parts: list[str] = []
    summary_rows = summary_rows_for_table(rows, definition["columns"])
    if summary_rows:
        summary_row = summary_rows[0]
        for index, cell in enumerate(summary_row):
            if not cell or cell == "Totali":
                continue
            summary_parts.append(f"{definition['columns'][index][1]} {cell}")

    lines = [
        f"Report {definition['title']} - {APP_NAME}",
        f"Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}",
    ]
    if filter_lines:
        lines.append("Filtri: " + " | ".join(filter_lines))
    lines.append(f"Righe visualizzate: {len(rows)}")
    if summary_parts:
        lines.append("Totali: " + " | ".join(summary_parts))

    detail_lines = [report_share_row_line(row, definition["columns"], index) for index, row in enumerate(rows, start=1)]
    footer = "Cordiali saluti,\nOratorio Carlo Acutis"
    if detail_lines:
        lines.append("Dettaglio righe visibili:")
        included = 0
        for detail_line in detail_lines:
            candidate = "\n".join(lines + [detail_line, footer])
            if max_length and len(candidate) > max_length:
                break
            lines.append(detail_line)
            included += 1
        if included < len(detail_lines):
            remaining = len(detail_lines) - included
            notice = (
                f"Dettaglio parziale per limiti del messaggio: incluse {included} righe su {len(detail_lines)}."
                if included
                else f"Dettaglio troppo esteso per il canale scelto: {remaining} righe non incluse nel messaggio."
            )
            candidate = "\n".join(lines + [notice, footer])
            if not max_length or len(candidate) <= max_length:
                lines.append(notice)
    lines.append(footer)
    return "\n".join(lines)


def build_report_share_payload(
    report_key: str,
    query_params: dict[str, str],
    recipient_id: str,
    channel: str,
) -> dict[str, str] | None:
    if channel not in {"email", "whatsapp", "whatsapp-group"}:
        raise ValueError("Canale di invio non valido.")
    recipient = None
    if channel == "whatsapp":
        recipient = next((row for row in report_recipient_rows() if row["id"] == recipient_id), None)
        if recipient is None:
            raise ValueError("Destinatario non valido.")

    definition = definition_with_search(get_report_definition(report_key, query_params), query_params)
    rows = filter_report_rows(
        fetch_all(definition["query"], definition.get("params", ())),
        definition["columns"],
        report_search_term(query_params),
    )
    subject = f"Report {definition['title']} - {APP_NAME}"
    message = report_share_message(
        definition,
        rows,
        max_length=6000 if channel == "email" else 2500,
    )

    if channel == "whatsapp-group":
        return {
            "channel": "whatsapp-group",
            "url": f"https://web.whatsapp.com/send?text={quote(message)}",
            "message": message,
        }

    if channel == "email":
        recipient_email = ""
        if recipient_id:
            recipient = next((row for row in report_recipient_rows() if row["id"] == recipient_id), None)
            if recipient is None:
                raise ValueError("Destinatario non valido.")
            if not recipient.get("email"):
                raise ValueError("Il destinatario selezionato non ha un indirizzo email registrato.")
            recipient_email = str(recipient["email"])
        return {
            "channel": "email",
            "url": f"mailto:{recipient_email}?subject={quote(subject)}&body={quote(message)}",
        }

    if not recipient.get("whatsapp_phone"):
        raise ValueError("Il destinatario selezionato non ha un numero cellulare registrato.")
    return {
        "channel": "whatsapp",
        "url": f"https://wa.me/{recipient['whatsapp_phone']}?text={quote(message)}",
    }


def report_toolbar(
    report_key: str,
    query_params: dict[str, str],
    *,
    definition: dict | None = None,
    rows: list[sqlite3.Row] | None = None,
) -> str:
    export_url = with_query(f"/export/excel/{report_key}", query_params)
    pdf_url = with_query(f"/export/pdf/{report_key}", query_params)
    search_value = report_search_term(query_params)
    email_button = ""
    whatsapp_group_button = ""
    toolbar_copy = "Esporta il report in Excel o PDF, oppure stampalo direttamente dal browser."
    if definition is not None and rows is not None:
        toolbar_copy = (
            "Esporta il report in Excel o PDF, stampalo oppure prepara l'invio via email o WhatsApp. "
            "Per WhatsApp il testo viene preparato e passato a WhatsApp Web, cosi puoi scegliere il gruppo e inviare subito."
        )
        whatsapp_group_button = (
            f'<button type="button" class="button action" '
            f'onclick="return shareReport(this)" '
            f'data-channel="whatsapp-group" data-report-key="{esc(report_key)}">'
            f'Invia WhatsApp</button>'
        )
        email_button = (
            f'<button type="button" class="button action" '
            f'onclick="return shareReport(this)" '
            f'data-channel="email" data-report-key="{esc(report_key)}">Prepara email</button>'
        )
    return f"""
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Azioni report</span>
        <p>{esc(toolbar_copy)}</p>
      </div>
      <div class="report-toolbar-actions">
        <label class="report-search">
          <span>Cerca</span>
          <input type="search" class="control" value="{esc(search_value)}" placeholder="Filtra tutte le colonne..." oninput="handleReportSearch(this)">
        </label>
        <span class="report-actions-break" aria-hidden="true"></span>
        <a class="button action" href="{esc(export_url)}" data-search-link="excel">Esporta Excel</a>
        <a class="button action" href="{esc(pdf_url)}" data-search-link="pdf">Esporta PDF</a>
        <button type="button" class="button action" onclick="window.print()">Stampa report</button>
        {email_button}
        {whatsapp_group_button}
      </div>
    </section>
    """


def build_scadenze_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    associato_id = normalized(query_params, "associato_id", "")
    area = normalized(query_params, "area", "")
    clauses = ["substr(COALESCE(scadenza, ''), 1, 4) = ?", "area <> 'Corso - iscrizione'"]
    params: list[object] = [str(work_year)]
    filters = [{"label": "Anno di lavoro", "value": str(work_year)}]
    lead_html = ""

    if associato_id.isdigit():
        clauses.append("associato_id = ?")
        params.append(int(associato_id))
        associato_label = lookup_label(
            "associati",
            associato_id,
            f"SELECT {associato_display_sql('')} AS associato FROM associati WHERE id = ?",
            "Associato selezionato",
        )
        filters.append({"label": "Associato", "value": associato_label})
        lead_html = f"""
            <section class="card compact screen-only">
              <div class="card-head">
                <h2>Filtro associato</h2>
                <p>Vista limitata all'associato selezionato dalla dashboard.</p>
              </div>
              <div class="empty-state">{esc(associato_label)}</div>
            </section>
        """

    if area:
        clauses.append("area = ?")
        params.append(area)
        filters.append({"label": "Area", "value": area})

    return {
        "title": "Scadenze da incassare",
        "current_path": "/report/scadenze",
        "subtitle": f"Vista unica con insoluti di tesseramenti, corsi, Campo estivo ed eventi per l'anno {work_year}.",
        "query": f"""
            SELECT area, riferimento, associato, scadenza, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
            FROM (
                SELECT
                    CASE WHEN area = 'Estate' THEN 'Campo estivo' ELSE area END AS area,
                    riferimento,
                    associato_id,
                    associato,
                    scadenza,
                    importo_dovuto,
                    importo_pagato,
                    saldo_residuo,
                    stato_pagamento
                FROM v_scadenze_da_incassare
            ) dati
            WHERE {' AND '.join(clauses)}
            ORDER BY scadenza, associato, area
        """,
        "params": tuple(params),
        "sheet_name": "Scadenze",
        "export_name": "scadenze_da_incassare.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("area", "Area"),
            ("riferimento", "Riferimento"),
            ("associato", "Associato"),
            ("scadenza", "Scadenza"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
    }


def build_tesseramenti_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    associato_id = normalized(query_params, "associato_id", "")
    associati = associati_options()
    clauses = ["anno_sociale = ?"]
    params: list[object] = [work_year]
    filters = [{"label": "Anno di lavoro", "value": str(work_year)}]
    if associato_id.isdigit():
        clauses.append("associato_id = ?")
        params.append(int(associato_id))
        associato_label = lookup_label(
            "associati",
            associato_id,
            f"SELECT {associato_display_sql('')} AS associato FROM associati WHERE id = ?",
            "",
        )
        if associato_label:
            filters.append({"label": "Associato", "value": associato_label})

    return {
        "title": "Situazione tesseramenti",
        "current_path": "/report/tesseramenti",
        "subtitle": f"Situazione completa dei tesseramenti per l'anno {work_year}.",
        "query": f"""
            SELECT associato, anno_sociale, data_tesseramento, data_scadenza, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
            FROM v_tesseramenti_saldo
            WHERE {' AND '.join(clauses)}
            ORDER BY associato
        """,
        "params": tuple(params),
        "sheet_name": "Tesseramenti",
        "export_name": f"situazione_tesseramenti_{work_year}.xlsx",
        "filters": filters,
        "lead_html": "",
        "columns": [
            ("associato", "Associato"),
            ("anno_sociale", "Anno"),
            ("data_tesseramento", "Data tesseramento"),
            ("data_scadenza", "Scadenza"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
    }


def build_registro_attivita_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    date_from, date_to = year_start_end(work_year)
    data_da = normalized(query_params, "data_da", date_from) or date_from
    data_a = normalized(query_params, "data_a", date_to) or date_to
    username = normalized(query_params, "username", "")
    associato_id = normalized(query_params, "associato_id", "")
    attivita = normalized(query_params, "attivita", "")

    clauses = ["date(substr(data_ora, 1, 10)) BETWEEN date(?) AND date(?)"]
    params: list[object] = [data_da, data_a]
    filters = [
        {"label": "Data da", "value": data_da},
        {"label": "Data a", "value": data_a},
    ]

    if username:
        clauses.append("username = ?")
        params.append(username)
        filters.append({"label": "Utente", "value": username})

    if associato_id.isdigit():
        clauses.append("associato_id = ?")
        params.append(int(associato_id))
        associato_label = lookup_label(
            "associati",
            associato_id,
            f"SELECT trim(codice_associato || ' - ' || {associato_display_sql('')}) AS associato FROM associati WHERE id = ?",
            "",
        )
        if associato_label:
            filters.append({"label": "Associato", "value": associato_label})

    if attivita:
        clauses.append("descrizione_attivita = ?")
        params.append(attivita)
        filters.append({"label": "Attivita", "value": attivita})

    lead_html = f"""
        <section class="card compact screen-only">
          <div class="card-head">
            <h2>Filtro registro attivita</h2>
            <p>Filtra il registro per intervallo date, utente e associato coinvolto.</p>
          </div>
          <form method="get" action="/report/registro-attivita" class="form-grid">
            <input type="hidden" name="anno_lavoro" value="{esc(work_year)}">
            {input_field("Data da", "data_da", input_type="date", value=data_da, required_field=True)}
            {input_field("Data a", "data_a", input_type="date", value=data_a, required_field=True)}
            {select_field("Utente", "username", render_select_options(activity_log_user_options(), username, blank_label="Tutti gli utenti"))}
            {select_field("Attivita", "attivita", render_select_options(activity_log_attivita_options(), attivita, blank_label="Tutte le attivita"))}
            {select_field("Associato", "associato_id", render_select_options(activity_log_associato_options(), associato_id, blank_label="Tutti gli associati", data_keys=["search_text", "autocomplete_label"]), wide=True, searchable=True, search_placeholder="Cerca associato nel registro...")}
            <div class="form-actions">
              <button type="submit" class="button">Aggiorna report</button>
            </div>
          </form>
        </section>
    """

    return {
        "title": "Registro attivita",
        "current_path": "/report/registro-attivita",
        "subtitle": "Storico operativo del gestionale con utente, associato coinvolto, dispositivo e risultato dell'attivita.",
        "query": f"""
            SELECT
                data_ora,
                username,
                trim(COALESCE(associato_codice, '') || CASE
                    WHEN COALESCE(associato_codice, '') <> '' AND COALESCE(associato_nominativo, '') <> '' THEN ' - '
                    ELSE ''
                END || COALESCE(associato_nominativo, '')) AS associato,
                nome_pc,
                categoria,
                descrizione_attivita,
                dettaglio,
                esito,
                percorso
            FROM registro_attivita
            WHERE {' AND '.join(clauses)}
            ORDER BY data_ora DESC, id DESC
        """,
        "params": tuple(params),
        "sheet_name": "Registro attivita",
        "export_name": "registro_attivita.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("data_ora", "Data ora"),
            ("username", "Utente"),
            ("associato", "Associato"),
            ("nome_pc", "Nome PC"),
            ("categoria", "Categoria"),
            ("descrizione_attivita", "Attivita"),
            ("dettaglio", "Dettaglio"),
            ("esito", "Esito"),
            ("percorso", "Percorso"),
        ],
    }


def build_corsi_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    course_id = normalized(query_params, "corso_id", "")
    comp_from = normalized(query_params, "comp_from", f"{work_year}-01-01")
    comp_to = normalized(query_params, "comp_to", f"{work_year}-12-31")
    course_name = lookup_label("corsi", course_id, "SELECT nome FROM corsi WHERE id = ?", "")
    clauses = ["anno = ?", "date(printf('%04d-%02d-01', anno, mese)) BETWEEN date(?) AND date(?)"]
    params: list[object] = [work_year, comp_from, comp_to]
    filters = [
        {"label": "Anno di lavoro", "value": str(work_year)},
        {"label": "Competenza da", "value": comp_from},
        {"label": "Competenza a", "value": comp_to},
    ]
    if course_name:
        clauses.append("corso = ?")
        params.append(course_name)
        filters.append({"label": "Corso", "value": course_name})

    lead_html = f"""
        <section class="card compact screen-only">
          <div class="card-head">
            <h2>Filtro quote mensili corsi</h2>
            <p>Filtra per corso e per competenza da data a data.</p>
          </div>
          <form method="get" action="/report/corsi" class="form-grid">
            <input type="hidden" name="anno_lavoro" value="{esc(work_year)}">
            {select_field("Corso", "corso_id", render_select_options(corsi_report_options(work_year), course_id), wide=True)}
            {input_field("Competenza dal", "comp_from", input_type="date", value=comp_from, required_field=True)}
            {input_field("Competenza al", "comp_to", input_type="date", value=comp_to, required_field=True)}
            <div class="form-actions">
              <button type="submit" class="button">Aggiorna report</button>
            </div>
          </form>
        </section>
    """

    return {
        "title": "Situazione corsi",
        "current_path": "/report/corsi",
        "subtitle": f"Vista operativa dei pagamenti mensili per competenza nell'anno {work_year}.",
        "query": f"""
            SELECT corso, associato, competenza, data_scadenza, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
            FROM v_rate_corsi_saldo
            WHERE {' AND '.join(clauses)}
            ORDER BY anno DESC, mese DESC, corso, associato
        """,
        "params": tuple(params),
        "sheet_name": "Situazione corsi",
        "export_name": "situazione_corsi.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("corso", "Corso"),
            ("associato", "Associato"),
            ("competenza", "Competenza"),
            ("data_scadenza", "Scadenza"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
    }


def build_campi_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    clauses = ["anno = ?"]
    params: list[object] = [work_year]
    filters = [{"label": "Anno di lavoro", "value": str(work_year)}]
    lead_html = ""

    return {
        "title": "Situazione campo estivo",
        "current_path": "/report/campi-estivi",
        "subtitle": f"Iscritti al Campo estivo con quota una tantum e situazione pagamenti per l'anno {work_year}.",
        "query": f"""
            SELECT anno, associato, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
            FROM v_campi_estivi_saldo
            WHERE {' AND '.join(clauses)}
            ORDER BY associato
        """,
        "params": tuple(params),
        "sheet_name": "Situazione campo estivo",
        "export_name": "situazione_campo_estivo.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("anno", "Anno"),
            ("associato", "Associato"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
    }


def build_eventi_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    evento_id = normalized(query_params, "evento_id", "")
    event_name = ""
    if evento_id.isdigit():
        event_row = fetch_one(
            """
            SELECT nome
            FROM eventi
            WHERE id = ?
              AND substr(COALESCE(data_evento, ''), 1, 4) = ?
            """,
            (int(evento_id), str(work_year)),
        )
        if event_row is not None:
            event_name = event_row["nome"]
    clauses = ["substr(data_evento, 1, 4) = ?"]
    params: list[object] = [str(work_year)]
    filters = [{"label": "Anno di lavoro", "value": str(work_year)}]
    if event_name:
        clauses.append("evento = ?")
        params.append(event_name)
        filters.append({"label": "Evento", "value": event_name})

    lead_html = f"""
        <section class="card compact screen-only">
          <div class="card-head">
            <h2>Filtro eventi</h2>
            <p>Seleziona l'evento da visualizzare nel report.</p>
          </div>
          <form method="get" action="/report/eventi" class="form-grid">
            <input type="hidden" name="anno_lavoro" value="{esc(work_year)}">
            {select_field("Evento", "evento_id", render_select_options(eventi_options(work_year), evento_id), wide=True)}
            <div class="form-actions">
              <button type="submit" class="button">Aggiorna report</button>
            </div>
          </form>
        </section>
    """

    return {
        "title": "Situazione eventi",
        "current_path": "/report/eventi",
        "subtitle": f"Partecipanti agli eventi con quota una tantum e stato pagamenti per l'anno {work_year}.",
        "query": f"""
            SELECT evento, tipologia, data_evento, associato, importo_dovuto, importo_pagato, saldo_residuo, stato_pagamento
            FROM v_eventi_saldo
            WHERE {' AND '.join(clauses)}
            ORDER BY data_evento DESC, evento, associato
        """,
        "params": tuple(params),
        "sheet_name": "Situazione eventi",
        "export_name": "situazione_eventi.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("evento", "Evento"),
            ("tipologia", "Tipologia"),
            ("data_evento", "Data"),
            ("associato", "Associato"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value)),
            ("importo_pagato", "Pagato", lambda value, _: money(value)),
            ("saldo_residuo", "Residuo", lambda value, _: money(value)),
            ("stato_pagamento", "Stato"),
        ],
    }


def build_incassi_report_definition(query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    default_from, default_to = year_start_end(work_year)
    date_from = normalized(query_params, "date_from", default_from) or default_from
    date_to = normalized(query_params, "date_to", default_to) or default_to
    area = normalized(query_params, "area", "")
    riferimento = normalized(query_params, "riferimento", "")
    associato_id = normalized(query_params, "associato_id", "")
    clauses = ["anno_riferimento = ?", "data_pagamento BETWEEN ? AND ?"]
    params: list[object] = [work_year, date_from, date_to]
    filters = [
        {"label": "Anno di lavoro", "value": str(work_year)},
        {"label": "Intervallo", "value": f"{date_from} -> {date_to}"},
    ]
    if area:
        clauses.append("area = ?")
        params.append(area)
        filters.append({"label": "Area", "value": area})
    if riferimento:
        clauses.append("LOWER(riferimento) LIKE ?")
        params.append(f"%{riferimento.lower()}%")
        filters.append({"label": "Riferimento", "value": riferimento})
    if associato_id.isdigit():
        clauses.append("associato_id = ?")
        params.append(int(associato_id))
        associato_label = lookup_label(
            "associati",
            associato_id,
            f"SELECT {associato_display_sql('')} AS associato FROM associati WHERE id = ?",
            "",
        )
        if associato_label:
            filters.append({"label": "Associato", "value": associato_label})

    lead_html = f"""
        <section class="card compact screen-only">
          <div class="card-head">
            <h2>Filtro report incassi</h2>
            <p>Filtra per intervallo, area e riferimento.</p>
          </div>
          <form method="get" action="/report/incassi" class="form-grid">
            <input type="hidden" name="anno_lavoro" value="{esc(work_year)}">
            <input type="hidden" name="associato_id" value="{esc(associato_id)}">
            {input_field("Dal", "date_from", input_type="date", value=date_from, required_field=True)}
            {input_field("Al", "date_to", input_type="date", value=date_to, required_field=True)}
            {select_field("Area", "area", render_static_options([
                ("", "Tutte le aree"),
                ("Tesseramento annuale", "Tesseramento annuale"),
                ("Corso - quota mensile", "Corso - quota mensile"),
                ("Campo estivo", ESTATE_LABEL),
                ("Evento", "Evento"),
            ], area, blank_label=None))}
            {input_field("Riferimento", "riferimento", value=riferimento, placeholder="Filtra per testo")}
            <div class="form-actions">
              <button type="submit" class="button">Aggiorna report</button>
            </div>
          </form>
        </section>
    """

    return {
        "title": "Incassi",
        "current_path": "/report/incassi",
        "subtitle": "Movimenti registrati nell'intervallo selezionato.",
        "query": f"""
            SELECT area, data_pagamento, importo, metodo_pagamento, associato, riferimento, payment_type, payment_id, gruppo_ricevuta
            FROM ({incassi_dataset_sql()}) incassi
            WHERE {' AND '.join(clauses)}
            ORDER BY data_pagamento, area, associato
        """,
        "params": tuple(params),
        "sheet_name": "Incassi",
        "export_name": f"report_incassi_{date_from}_{date_to}.xlsx",
        "filters": filters,
        "lead_html": lead_html,
        "columns": [
            ("area", "Area"),
            ("data_pagamento", "Data"),
            ("importo", "Importo", lambda value, _: money(value)),
            ("metodo_pagamento", "Metodo"),
            ("associato", "Associato"),
            ("riferimento", "Riferimento"),
            (
                "payment_id",
                "Ricevuta",
                lambda value, row: (
                    report_link(
                        "Apri",
                        f"/ricevute/corsi-rate-gruppo/{row['gruppo_ricevuta']}"
                        if row["gruppo_ricevuta"]
                        else f"/ricevute/{row['payment_type']}/{row['payment_id']}",
                        work_year_query(query_params),
                    )
                    if row["payment_id"]
                    else ""
                ),
            ),
        ],
    }


def get_report_definition(report_key: str, query_params: dict[str, str]) -> dict:
    work_year = current_work_year(query_params)
    if report_key == "registro-attivita":
        return build_registro_attivita_report_definition(query_params)

    if report_key == "tesseramenti":
        return build_tesseramenti_report_definition(query_params)

    if report_key == "scadenze":
        return build_scadenze_report_definition(query_params)

    if report_key == "corsi":
        return build_corsi_report_definition(query_params)

    if report_key == "campi-estivi":
        return build_campi_report_definition(query_params)

    if report_key == "eventi":
        return build_eventi_report_definition(query_params)

    if report_key == "incassi":
        return build_incassi_report_definition(query_params)

    if report_key == "associati":
        return {
            "title": "Posizione associati",
            "current_path": "/report/associati",
            "subtitle": "Totale dovuto, totale pagato e saldo residuo dell'anno di lavoro per associato.",
            "query": posizione_associati_query(),
            "params": posizione_associati_params(work_year),
            "sheet_name": "Associati",
            "export_name": "posizione_associati.xlsx",
            "filters": [{"label": "Anno di lavoro", "value": str(work_year)}],
            "columns": [
                ("codice_associato", "Codice"),
                (
                    "associato",
                    "Associato",
                    lambda value, row: report_link(value, f"/report/associato/{row['associato_id']}", work_year_query(query_params)),
                ),
                ("stato_associato", "Stato"),
                ("totale_dovuto", "Totale dovuto", lambda value, _: money(value)),
                ("totale_pagato", "Totale pagato", lambda value, _: money(value)),
                ("saldo_residuo", "Saldo residuo", lambda value, _: money(value)),
            ],
        }

    if report_key == "partecipanti":
        return build_participants_report_definition(query_params)

    raise KeyError(report_key)


def report_page(report_key: str, query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    definition = definition_with_search(get_report_definition(report_key, query_params), query_params)
    rows = filter_report_rows(
        fetch_all(definition["query"], definition.get("params", ())),
        definition["columns"],
        report_search_term(query_params),
    )
    summary_rows = summary_rows_for_table(rows, definition["columns"])
    content = (
        report_toolbar(report_key, query_params, definition=definition, rows=rows)
        + definition.get("lead_html", "")
        + table_card(
            definition["title"],
            definition["subtitle"],
            rows,
            definition["columns"],
            table_class="report-table",
            summary_rows=summary_rows,
        )
    )
    return page(definition["title"], definition["current_path"], content, query_params, current_user)


def report_associati(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("associati", query_params, current_user)


def report_tesseramenti(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("tesseramenti", query_params, current_user)


def report_scadenze(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("scadenze", query_params, current_user)


def report_corsi(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("corsi", query_params, current_user)


def report_campi_estivi(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("campi-estivi", query_params, current_user)


def report_eventi(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("eventi", query_params, current_user)


def report_partecipanti(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("partecipanti", query_params, current_user)


def report_incassi(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("incassi", query_params, current_user)


def report_registro_attivita(query_params: dict[str, str], current_user: dict[str, object] | None = None) -> bytes:
    return report_page("registro-attivita", query_params, current_user)


def associato_iscrizioni_rows(associato_id: int, work_year: int) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []

    tesseramenti = fetch_all(
        """
        SELECT
            'Tesseramento' AS area,
            'Anno ' || anno_sociale AS riferimento,
            data_tesseramento AS data_riferimento,
            CASE
                WHEN saldo_residuo <= 0 THEN 'Saldo completo'
                WHEN importo_pagato > 0 THEN 'Pagamento parziale'
                ELSE 'Da saldare'
            END AS stato,
            importo_dovuto AS importo
        FROM v_tesseramenti_saldo
        WHERE associato_id = ? AND anno_sociale = ?
        ORDER BY data_tesseramento DESC
        """,
        (associato_id, work_year),
    )
    rows.extend(dict(row) for row in tesseramenti)

    corsi = fetch_all(
        f"""
        SELECT
            'Corso' AS area,
            c.nome AS riferimento,
            ic.data_iscrizione AS data_riferimento,
            ic.stato_iscrizione AS stato,
            ic.quota_mensile AS importo
        FROM iscrizioni_corsi ic
        JOIN corsi c ON c.id = ic.corso_id
        WHERE ic.associato_id = ?
          AND {iscrizione_corso_year_relevance_sql('ic')}
        ORDER BY ic.data_iscrizione DESC, c.nome
        """,
        (associato_id, *iscrizione_corso_year_relevance_params(work_year)),
    )
    rows.extend(dict(row) for row in corsi)

    campi = fetch_all(
        """
        SELECT
            'Campo estivo' AS area,
            ce.nome AS riferimento,
            ice.data_iscrizione AS data_riferimento,
            ice.stato_iscrizione AS stato,
            ice.quota_partecipazione AS importo
        FROM iscrizioni_campi_estivi ice
        JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
        WHERE ice.associato_id = ?
          AND ce.anno = ?
        ORDER BY ice.data_iscrizione DESC, ce.nome
        """,
        (associato_id, work_year),
    )
    rows.extend(dict(row) for row in campi)

    eventi = fetch_all(
        """
        SELECT
            'Evento' AS area,
            e.nome AS riferimento,
            ie.data_iscrizione AS data_riferimento,
            ie.stato_iscrizione AS stato,
            ie.quota_partecipazione AS importo
        FROM iscrizioni_eventi ie
        JOIN eventi e ON e.id = ie.evento_id
        WHERE ie.associato_id = ?
          AND substr(COALESCE(e.data_evento, ''), 1, 4) = ?
        ORDER BY ie.data_iscrizione DESC, e.nome
        """,
        (associato_id, str(work_year)),
    )
    rows.extend(dict(row) for row in eventi)
    return rows


def build_associato_detail_export_payload(
    associato: sqlite3.Row,
    scoped_query: dict[str, str],
    iscrizioni_rows: list[dict[str, object]],
    scadenze_rows: list[sqlite3.Row],
    incassi_rows: list[sqlite3.Row],
) -> tuple[dict[str, object], list[dict[str, object]]]:
    work_year = current_work_year(scoped_query)
    filters = [
        {"label": "Anno di lavoro", "value": str(work_year)},
        {"label": "Associato", "value": plain_text(associato["associato"])},
        {"label": "Codice", "value": plain_text(associato["codice_associato"])},
    ]
    rows: list[dict[str, object]] = []

    for row in iscrizioni_rows:
        rows.append(
            {
                "sezione": "Iscrizioni",
                "area": row.get("area", ""),
                "riferimento": row.get("riferimento", ""),
                "data_riferimento": row.get("data_riferimento", ""),
                "stato": row.get("stato", ""),
                "importo_dovuto": row.get("importo", 0),
                "importo_pagato": "",
                "saldo_residuo": "",
                "importo": "",
            }
        )

    for row in scadenze_rows:
        rows.append(
            {
                "sezione": "Scadenze",
                "area": row["area"],
                "riferimento": row["riferimento"],
                "data_riferimento": row["scadenza"],
                "stato": row["stato_pagamento"],
                "importo_dovuto": row["importo_dovuto"],
                "importo_pagato": row["importo_pagato"],
                "saldo_residuo": row["saldo_residuo"],
                "importo": "",
            }
        )

    for row in incassi_rows:
        rows.append(
            {
                "sezione": "Incassi",
                "area": row["area"],
                "riferimento": row["riferimento"],
                "data_riferimento": row["data_pagamento"],
                "stato": row["metodo_pagamento"],
                "importo_dovuto": "",
                "importo_pagato": "",
                "saldo_residuo": "",
                "importo": row["importo"],
            }
        )

    definition = {
        "title": "Dettaglio associato",
        "subtitle": f"Riepilogo iscrizioni, scadenze e incassi di {plain_text(associato['associato'])}.",
        "sheet_name": "Dettaglio associato",
        "export_name": f"dettaglio_associato_{associato['codice_associato']}.xlsx",
        "filters": filters,
        "columns": [
            ("sezione", "Sezione"),
            ("area", "Area"),
            ("riferimento", "Riferimento"),
            ("data_riferimento", "Data"),
            ("stato", "Stato"),
            ("importo_dovuto", "Dovuto", lambda value, _: money(value) if value not in ("", None) else ""),
            ("importo_pagato", "Pagato", lambda value, _: money(value) if value not in ("", None) else ""),
            ("saldo_residuo", "Residuo", lambda value, _: money(value) if value not in ("", None) else ""),
            ("importo", "Importo", lambda value, _: money(value) if value not in ("", None) else ""),
        ],
    }
    return definition, rows


def associato_detail_toolbar(associato_id: int, scoped_query: dict[str, str]) -> str:
    excel_url = with_query(f"/export/excel/associato/{associato_id}", scoped_query)
    pdf_url = with_query(f"/export/pdf/associato/{associato_id}", scoped_query)
    return f"""
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Azioni dettaglio</span>
        <p>Esporta il dettaglio associato in Excel o PDF, oppure stampalo direttamente dal browser.</p>
      </div>
      <div class="report-toolbar-actions">
        <a class="button action" href="{esc(excel_url)}">Esporta Excel</a>
        <a class="button action" href="{esc(pdf_url)}">Esporta PDF</a>
        <button type="button" class="button action" onclick="window.print()">Stampa report</button>
      </div>
    </section>
    """


def associato_report_page(
    associato_id: int,
    query_params: dict[str, str],
    current_user: dict[str, object] | None = None,
) -> bytes:
    associato = fetch_one(
        f"""
        SELECT id, codice_associato, {associato_display_sql('')} AS associato, email, telefono, stato_associato
        FROM associati
        WHERE id = ?
        """,
        (associato_id,),
    )
    if associato is None:
        raise KeyError(associato_id)

    scoped_query = dict(work_year_query(query_params))
    scoped_query["associato_id"] = str(associato_id)
    iscrizioni_rows = associato_iscrizioni_rows(associato_id, current_work_year(scoped_query))
    incassi_definition = build_incassi_report_definition(scoped_query)
    scadenze_definition = build_scadenze_report_definition(scoped_query)
    incassi_rows = fetch_all(incassi_definition["query"], incassi_definition["params"])
    scadenze_rows = fetch_all(scadenze_definition["query"], scadenze_definition["params"])
    iscrizioni_columns = [
        ("area", "Area"),
        ("riferimento", "Riferimento"),
        ("data_riferimento", "Data iscrizione"),
        ("stato", "Stato"),
        ("importo", "Importo", lambda value, _: money(value)),
    ]
    iscrizioni_summary = summary_rows_for_table(iscrizioni_rows, iscrizioni_columns)
    incassi_summary = summary_rows_for_table(incassi_rows, incassi_definition["columns"])
    scadenze_summary = summary_rows_for_table(scadenze_rows, scadenze_definition["columns"])
    scadenze_share_actions = associato_scadenze_share_actions(
        associato,
        scadenze_rows,
        current_work_year(scoped_query),
    )

    content = f"""
    <section class="hero">
      <div>
        <span class="eyebrow">Dettaglio associato</span>
        <h2>{esc(associato['associato'])}</h2>
        <p>Codice {esc(associato['codice_associato'])} | Stato {esc(associato['stato_associato'])} | Email {esc(associato['email'] or '-')} | Cellulare {esc(associato['telefono'] or '-')}</p>
      </div>
      <div class="hero-actions">
        <a class="button ghost" href="{esc(with_query('/report/incassi', scoped_query))}">Apri incassi filtrati</a>
        <a class="button ghost" href="{esc(with_query('/report/scadenze', scoped_query))}">Apri scadenze filtrate</a>
      </div>
    </section>
    {associato_detail_toolbar(associato_id, scoped_query)}
    {table_card(
        "Iscrizioni associato",
        "Riepilogo di tesseramento e iscrizioni dell'associato nell'anno di lavoro selezionato.",
        iscrizioni_rows,
        iscrizioni_columns,
        table_class="report-table",
        summary_rows=iscrizioni_summary,
    )}
    {table_card(
        "Scadenze dell'associato",
        "Quote ancora aperte o parzialmente saldate.",
        scadenze_rows,
        scadenze_definition["columns"],
        table_class="report-table",
        summary_rows=scadenze_summary,
        head_actions_html=scadenze_share_actions,
    )}
    {table_card(
        "Incassi dell'associato",
        "Movimenti registrati per l'associato selezionato.",
        incassi_rows,
        incassi_definition["columns"],
        table_class="report-table",
        summary_rows=incassi_summary,
    )}
    """
    return page("Dettaglio associato", "/report/associati", content, query_params, current_user)


def export_rows(rows: list[sqlite3.Row], columns: list[tuple]) -> list[list[object]]:
    exported_rows = [[report_display_value(row, column) for column in columns] for row in rows]
    exported_rows.extend(summary_rows_for_table(rows, columns))
    return exported_rows


def generate_report_xlsx(definition: dict, rows: list[sqlite3.Row]) -> tuple[str, bytes]:
    if not NODE_BIN.exists():
        raise RuntimeError(f"Runtime Node non trovato: {NODE_BIN}")
    if not ARTIFACT_TOOL_MODULE.exists():
        raise RuntimeError(f"Modulo export Excel non trovato: {ARTIFACT_TOOL_MODULE}")
    if not EXPORT_SCRIPT.exists():
        raise RuntimeError(f"Script export report non trovato: {EXPORT_SCRIPT}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temp_dir = OUTPUT_DIR / f"report-export-{uuid.uuid4().hex[:8]}"
    temp_dir.mkdir(parents=True, exist_ok=False)
    try:
        payload = {
            "title": definition["title"],
            "subtitle": definition["subtitle"],
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sheetName": definition["sheet_name"],
            "emptyMessage": "Nessun dato disponibile.",
            "filters": definition.get("filters", []),
            "columns": [{"key": column[0], "label": column[1]} for column in definition["columns"]],
            "rows": export_rows(rows, definition["columns"]),
        }

        payload_path = temp_dir / "report.json"
        output_path = temp_dir / definition["export_name"]
        payload_path.write_text(
            json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )

        environment = os.environ.copy()
        environment["ASSOCIAZIONE_ARTIFACT_TOOL_MODULE"] = str(ARTIFACT_TOOL_MODULE)

        completed = subprocess.run(
            [str(NODE_BIN), str(EXPORT_SCRIPT), str(payload_path), str(output_path)],
            capture_output=True,
            text=True,
            timeout=120,
            cwd=str(BASE_DIR),
            env=environment,
            check=False,
        )
        if completed.returncode != 0 or not output_path.exists():
            raise RuntimeError(
                completed.stderr.strip()
                or completed.stdout.strip()
                or "Export Excel non riuscito."
            )

        return output_path.name, output_path.read_bytes()
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def export_report_excel(start_response, report_key: str, query_params: dict[str, str]):
    try:
        definition = definition_with_search(get_report_definition(report_key, query_params), query_params)
    except KeyError:
        return not_found(start_response)

    try:
        rows = filter_report_rows(
            fetch_all(definition["query"], definition.get("params", ())),
            definition["columns"],
            report_search_term(query_params),
        )
        filename, content = generate_report_xlsx(definition, rows)
        start_response(
            "200 OK",
            [
                (
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                ("Content-Disposition", f'attachment; filename="{filename}"'),
                ("Content-Length", str(len(content))),
            ],
        )
        return [content]
    except Exception as error:
        start_response(
            "500 Internal Server Error",
            [("Content-Type", "text/plain; charset=utf-8")],
        )
        return [f"Errore durante l'export Excel: {error}".encode("utf-8")]


def export_associato_detail_excel(start_response, associato_id: int, query_params: dict[str, str]):
    associato = fetch_one(
        f"""
        SELECT id, codice_associato, {associato_display_sql('')} AS associato, email, telefono, stato_associato
        FROM associati
        WHERE id = ?
        """,
        (associato_id,),
    )
    if associato is None:
        return not_found(start_response)

    scoped_query = dict(work_year_query(query_params))
    scoped_query["associato_id"] = str(associato_id)
    iscrizioni_rows = associato_iscrizioni_rows(associato_id, current_work_year(scoped_query))
    incassi_definition = build_incassi_report_definition(scoped_query)
    scadenze_definition = build_scadenze_report_definition(scoped_query)
    incassi_rows = fetch_all(incassi_definition["query"], incassi_definition["params"])
    scadenze_rows = fetch_all(scadenze_definition["query"], scadenze_definition["params"])
    definition, rows = build_associato_detail_export_payload(associato, scoped_query, iscrizioni_rows, scadenze_rows, incassi_rows)
    try:
        filename, content = generate_report_xlsx(definition, rows)
        start_response(
            "200 OK",
            [
                ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ("Content-Disposition", f'attachment; filename="{filename}"'),
                ("Content-Length", str(len(content))),
            ],
        )
        return [content]
    except Exception as error:
        start_response("500 Internal Server Error", [("Content-Type", "text/plain; charset=utf-8")])
        return [f"Errore durante l'export Excel: {error}".encode("utf-8")]


def generate_report_pdf(definition: dict, rows: list[sqlite3.Row]) -> tuple[str, bytes]:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=20,
        textColor=colors.HexColor("#cb5f07"),
        spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#5f4a3c"),
        spaceAfter=4,
    )
    meta_style = ParagraphStyle(
        "ReportMeta",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#5f4a3c"),
        spaceAfter=3,
    )
    cell_style = ParagraphStyle(
        "ReportCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=7.7,
        leading=9.2,
        textColor=colors.HexColor("#3d2d22"),
        wordWrap="CJK",
    )
    header_cell_style = ParagraphStyle(
        "ReportHeaderCell",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        fontSize=8.1,
        leading=9.6,
    )
    summary_cell_style = ParagraphStyle(
        "ReportSummaryCell",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#6e3d14"),
    )

    story = []
    logo_path = STATIC_DIR / "logo-ca.jpg"
    if logo_path.exists():
        header_logo = Image(str(logo_path), width=24 * mm, height=24 * mm)
    else:
        header_logo = ""

    filter_lines = [f"<b>Generato il:</b> {esc(datetime.now().strftime('%d/%m/%Y %H:%M'))}"]
    for filter_row in definition.get("filters", []):
        filter_lines.append(f"<b>{esc(filter_row['label'])}:</b> {esc(filter_row['value'])}")

    header_table = Table(
        [
            [
                header_logo,
                [
                    Paragraph(APP_NAME, title_style),
                    Paragraph(definition["title"], title_style),
                    Paragraph(definition["subtitle"], subtitle_style),
                    Paragraph(" | ".join(filter_lines), meta_style),
                ],
            ]
        ],
        colWidths=[30 * mm, 240 * mm],
    )
    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(header_table)
    story.append(Spacer(1, 5 * mm))

    summary_rows = summary_rows_for_table(rows, definition["columns"])
    body_rows = [[report_display_value(row, column) for column in definition["columns"]] for row in rows]
    if rows:
        raw_table_rows = body_rows + summary_rows
    else:
        raw_table_rows = [[definition.get("empty_message", "Nessun dato disponibile.")] + [""] * (len(definition["columns"]) - 1)]

    def pdf_text(value: object) -> str:
        return esc(plain_text(value)).replace("\n", "<br/>")

    def distribute_column_widths(weights: list[float], total_width: float) -> list[float]:
        column_count = max(len(weights), 1)
        min_width = min(18 * mm, total_width / column_count)
        max_width = max(34 * mm, total_width * 0.32)
        widths = [min_width] * column_count
        remaining = max(total_width - sum(widths), 0)
        adjustable = set(range(column_count))
        normalized_weights = [max(float(weight), 1.0) for weight in weights]

        while adjustable and remaining > 0.1:
            total_weight = sum(normalized_weights[index] for index in adjustable) or float(len(adjustable))
            capped_any = False
            for index in list(adjustable):
                share = remaining * (normalized_weights[index] / total_weight)
                target_width = widths[index] + share
                if target_width >= max_width:
                    remaining -= max_width - widths[index]
                    widths[index] = max_width
                    adjustable.remove(index)
                    capped_any = True
            if not capped_any:
                total_weight = sum(normalized_weights[index] for index in adjustable) or float(len(adjustable))
                for index in adjustable:
                    widths[index] += remaining * (normalized_weights[index] / total_weight)
                remaining = 0

        total_current = sum(widths)
        if total_current > total_width and total_current > 0:
            scale = total_width / total_current
            widths = [width * scale for width in widths]
        return widths

    def column_weight(column_index: int, header_label: str) -> float:
        samples = [plain_text(header_label)]
        for raw_row in raw_table_rows:
            if column_index < len(raw_row):
                samples.append(plain_text(raw_row[column_index]))
        longest = max((max(len(part) for part in sample.splitlines()) if sample else 0) for sample in samples)
        return min(max(longest, 8), 42)

    column_count = max(len(definition["columns"]), 1)
    available_width = 277 * mm
    col_widths = distribute_column_widths(
        [column_weight(index, column[1]) for index, column in enumerate(definition["columns"])],
        available_width,
    )

    table_data: list[list[object]] = [[Paragraph(pdf_text(column[1]), header_cell_style) for column in definition["columns"]]]
    if rows:
        for raw_row in body_rows:
            table_data.append([Paragraph(pdf_text(value), cell_style) for value in raw_row])
        for raw_summary in summary_rows:
            table_data.append([Paragraph(pdf_text(value), summary_cell_style) for value in raw_summary])
    else:
        table_data.append(
            [Paragraph(pdf_text(value), cell_style) for value in raw_table_rows[0]]
        )

    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ef7f1a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#f0d2b8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fff7f0")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if rows and summary_rows:
        first_summary_row = len(table_data) - len(summary_rows)
        style_commands.extend(
            [
                ("BACKGROUND", (0, first_summary_row), (-1, -1), colors.HexColor("#fff0de")),
                ("FONTNAME", (0, first_summary_row), (-1, -1), "Helvetica-Bold"),
            ]
        )
    if not rows:
        style_commands.extend(
            [
                ("SPAN", (0, 1), (-1, 1)),
                ("ALIGN", (0, 1), (-1, 1), "CENTER"),
            ]
        )
    table.setStyle(TableStyle(style_commands))
    story.append(table)

    document.build(story)
    filename = definition.get("pdf_name") or definition["export_name"].replace(".xlsx", ".pdf")
    return filename, buffer.getvalue()


def export_report_pdf(start_response, report_key: str, query_params: dict[str, str]):
    try:
        definition = definition_with_search(get_report_definition(report_key, query_params), query_params)
    except KeyError:
        return not_found(start_response)

    try:
        rows = filter_report_rows(
            fetch_all(definition["query"], definition.get("params", ())),
            definition["columns"],
            report_search_term(query_params),
        )
        filename, content = generate_report_pdf(definition, rows)
        start_response(
            "200 OK",
            [
                ("Content-Type", "application/pdf"),
                ("Content-Disposition", f'attachment; filename="{filename}"'),
                ("Content-Length", str(len(content))),
            ],
        )
        return [content]
    except Exception as error:
        start_response(
            "500 Internal Server Error",
            [("Content-Type", "text/plain; charset=utf-8")],
        )
        return [f"Errore durante l'export PDF: {error}".encode("utf-8")]


def export_associato_detail_pdf(start_response, associato_id: int, query_params: dict[str, str]):
    associato = fetch_one(
        f"""
        SELECT id, codice_associato, {associato_display_sql('')} AS associato, email, telefono, stato_associato
        FROM associati
        WHERE id = ?
        """,
        (associato_id,),
    )
    if associato is None:
        return not_found(start_response)

    scoped_query = dict(work_year_query(query_params))
    scoped_query["associato_id"] = str(associato_id)
    iscrizioni_rows = associato_iscrizioni_rows(associato_id, current_work_year(scoped_query))
    incassi_definition = build_incassi_report_definition(scoped_query)
    scadenze_definition = build_scadenze_report_definition(scoped_query)
    incassi_rows = fetch_all(incassi_definition["query"], incassi_definition["params"])
    scadenze_rows = fetch_all(scadenze_definition["query"], scadenze_definition["params"])
    definition, rows = build_associato_detail_export_payload(associato, scoped_query, iscrizioni_rows, scadenze_rows, incassi_rows)
    try:
        filename, content = generate_report_pdf(definition, rows)
        start_response(
            "200 OK",
            [
                ("Content-Type", "application/pdf"),
                ("Content-Disposition", f'attachment; filename="{filename}"'),
                ("Content-Length", str(len(content))),
            ],
        )
        return [content]
    except Exception as error:
        start_response("500 Internal Server Error", [("Content-Type", "text/plain; charset=utf-8")])
        return [f"Errore durante l'export PDF: {error}".encode("utf-8")]


def receipt_link(payment_type: str, payment_id: object, query_params: dict[str, str]) -> str:
    return with_query(f"/ricevute/{payment_type}/{payment_id}", work_year_query(query_params))


def grouped_receipt_link(group_code: str, query_params: dict[str, str]) -> str:
    return with_query(f"/ricevute/corsi-rate-gruppo/{group_code}", work_year_query(query_params))


def multi_area_receipt_link(group_code: str, query_params: dict[str, str]) -> str:
    return with_query(f"/ricevute/multi-area-gruppo/{group_code}", work_year_query(query_params))


def popup_payment_requested(form_data: dict[str, str]) -> bool:
    return normalized(form_data, "procedi_pagamento", "") == "1"


def popup_payment_payload(form_data: dict[str, str], default_date: str) -> tuple[str, Decimal, str]:
    metodo_pagamento_id = required(form_data, "pagamento_metodo_id", "Metodo pagamento")
    importo = decimal_amount(required(form_data, "pagamento_importo", "Importo pagato"), minimum="0.01")
    data_pagamento = default_date or date.today().isoformat()
    return metodo_pagamento_id, importo, data_pagamento


def build_receipt_context(payment_type: str, payment_id: int) -> dict | None:
    payment_id = int(payment_id)
    if payment_type == "tesseramenti":
        row = fetch_one(
            f"""
            SELECT
                pt.id,
                pt.data_pagamento,
                pt.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pt.riferimento, '') AS riferimento,
                COALESCE(pt.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Tesseramento annuale' AS area,
                'Quota associativa anno ' || t.anno_sociale AS causale,
                t.anno_sociale AS work_year
            FROM pagamenti_tesseramenti pt
            JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
            JOIN associati a ON a.id = t.associato_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id
            WHERE pt.id = ?
            """,
            (payment_id,),
        )
        prefix = "TES"
    elif payment_type == "corsi-iscrizione":
        row = fetch_one(
            f"""
            SELECT
                pic.id,
                pic.data_pagamento,
                pic.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pic.riferimento, '') AS riferimento,
                COALESCE(pic.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Corso - iscrizione' AS area,
                'Quota iscrizione corso ' || c.nome AS causale,
                CAST(substr(COALESCE(ic.data_iscrizione, ic.data_inizio, ''), 1, 4) AS INTEGER) AS work_year
            FROM pagamenti_iscrizioni_corsi pic
            JOIN iscrizioni_corsi ic ON ic.id = pic.iscrizione_corso_id
            JOIN associati a ON a.id = ic.associato_id
            JOIN corsi c ON c.id = ic.corso_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pic.metodo_pagamento_id
            WHERE pic.id = ?
            """,
            (payment_id,),
        )
        prefix = "CIS"
    elif payment_type == "corsi-rata":
        row = fetch_one(
            f"""
            SELECT
                prc.id,
                prc.data_pagamento,
                prc.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(prc.riferimento, '') AS riferimento,
                COALESCE(prc.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Corso - quota mensile' AS area,
                'Quota mensile corso ' || c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS causale,
                r.anno AS work_year
            FROM pagamenti_rate_corsi prc
            JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
            JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
            JOIN associati a ON a.id = ic.associato_id
            JOIN corsi c ON c.id = ic.corso_id
            LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id
            WHERE prc.id = ?
            """,
            (payment_id,),
        )
        prefix = "CRM"
    elif payment_type == "campi-estivi":
        row = fetch_one(
            f"""
            SELECT
                pce.id,
                pce.data_pagamento,
                pce.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pce.riferimento, '') AS riferimento,
                COALESCE(pce.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Campo estivo' AS area,
                'Quota partecipazione Campo estivo ' || ce.anno AS causale,
                ce.anno AS work_year
            FROM pagamenti_campi_estivi pce
            JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
            JOIN associati a ON a.id = ice.associato_id
            JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id
            WHERE pce.id = ?
            """,
            (payment_id,),
        )
        prefix = "CES"
    elif payment_type == "eventi":
        row = fetch_one(
            f"""
            SELECT
                pe.id,
                pe.data_pagamento,
                pe.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pe.riferimento, '') AS riferimento,
                COALESCE(pe.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Evento' AS area,
                'Quota partecipazione ' || e.nome AS causale,
                CAST(substr(COALESCE(e.data_evento, ''), 1, 4) AS INTEGER) AS work_year
            FROM pagamenti_eventi pe
            JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
            JOIN associati a ON a.id = ie.associato_id
            JOIN eventi e ON e.id = ie.evento_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id
            WHERE pe.id = ?
            """,
            (payment_id,),
        )
        prefix = "EVT"
    else:
        return None

    if row is None:
        return None

    context = dict(row)
    context["payment_type"] = payment_type
    context["receipt_number"] = f"RCP-{prefix}-{payment_id:06d}"
    context["whatsapp_phone"] = clean_phone_number(context.get("telefono"))
    return context


def build_grouped_rate_receipt_context(group_code: str) -> dict | None:
    rows = fetch_all(
        f"""
        SELECT
            prc.id,
            prc.data_pagamento,
            prc.importo,
            COALESCE(mp.nome, '') AS metodo_pagamento,
            COALESCE(prc.riferimento, '') AS riferimento,
            COALESCE(prc.note, '') AS note,
            a.id AS associato_id,
            a.codice_associato,
            {associato_display_sql('a')} AS associato,
            COALESCE(a.email, '') AS email,
            COALESCE(a.telefono, '') AS telefono,
            c.nome AS corso,
            printf('%04d-%02d', r.anno, r.mese) AS competenza,
            r.anno AS work_year
        FROM pagamenti_rate_corsi prc
        JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
        JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
        JOIN associati a ON a.id = ic.associato_id
        JOIN corsi c ON c.id = ic.corso_id
        LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id
        WHERE prc.gruppo_ricevuta = ?
        ORDER BY r.anno, r.mese, c.nome
        """,
        (group_code,),
    )
    if not rows:
        return None

    first = dict(rows[0])
    items = [
        {
            "corso": row["corso"],
            "competenza": row["competenza"],
            "importo": float(row["importo"] or 0),
        }
        for row in rows
    ]
    total_amount = sum(item["importo"] for item in items)
    return {
        "id": group_code,
        "payment_type": "corsi-rate-gruppo",
        "receipt_number": f"RCP-CRMSET-{group_code}",
        "data_pagamento": first["data_pagamento"],
        "importo": total_amount,
        "metodo_pagamento": first["metodo_pagamento"],
        "riferimento": first["riferimento"],
        "note": first["note"],
        "codice_associato": first["codice_associato"],
        "associato": first["associato"],
        "email": first["email"],
        "telefono": first["telefono"],
        "area": "Corso - quote mensili",
        "causale": "SALDO QUOTE MENSILI CORSI",
        "items": items,
        "items_mode": "mensilita",
        "associato_id": int(first["associato_id"]),
        "work_year": int(first["work_year"] or date.today().year),
        "whatsapp_phone": clean_phone_number(first["telefono"]),
    }


def build_multi_area_receipt_context(group_code: str) -> dict | None:
    rows = fetch_all(
        f"""
        SELECT *
        FROM (
            SELECT
                pt.id,
                pt.data_pagamento,
                pt.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pt.riferimento, '') AS riferimento,
                COALESCE(pt.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Tesseramento annuale' AS area,
                'Anno ' || t.anno_sociale AS scadenza_riferimento,
                COALESCE(t.data_scadenza, t.data_tesseramento) AS scadenza,
                t.anno_sociale AS work_year
            FROM pagamenti_tesseramenti pt
            JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
            JOIN associati a ON a.id = t.associato_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id
            WHERE pt.gruppo_ricevuta = ?

            UNION ALL

            SELECT
                prc.id,
                prc.data_pagamento,
                prc.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(prc.riferimento, '') AS riferimento,
                COALESCE(prc.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Corso - quota mensile' AS area,
                c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS scadenza_riferimento,
                COALESCE(r.data_scadenza, printf('%04d-%02d-01', r.anno, r.mese)) AS scadenza,
                r.anno AS work_year
            FROM pagamenti_rate_corsi prc
            JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
            JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
            JOIN associati a ON a.id = ic.associato_id
            JOIN corsi c ON c.id = ic.corso_id
            LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id
            WHERE prc.gruppo_ricevuta = ?

            UNION ALL

            SELECT
                pce.id,
                pce.data_pagamento,
                pce.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pce.riferimento, '') AS riferimento,
                COALESCE(pce.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Campo estivo' AS area,
                ce.nome AS scadenza_riferimento,
                COALESCE(ce.data_inizio, ice.data_iscrizione) AS scadenza,
                ce.anno AS work_year
            FROM pagamenti_campi_estivi pce
            JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
            JOIN associati a ON a.id = ice.associato_id
            JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id
            WHERE pce.gruppo_ricevuta = ?

            UNION ALL

            SELECT
                pe.id,
                pe.data_pagamento,
                pe.importo,
                COALESCE(mp.nome, '') AS metodo_pagamento,
                COALESCE(pe.riferimento, '') AS riferimento,
                COALESCE(pe.note, '') AS note,
                a.id AS associato_id,
                a.codice_associato,
                {associato_display_sql('a')} AS associato,
                COALESCE(a.email, '') AS email,
                COALESCE(a.telefono, '') AS telefono,
                'Evento' AS area,
                e.nome AS scadenza_riferimento,
                e.data_evento AS scadenza,
                CAST(substr(COALESCE(e.data_evento, ''), 1, 4) AS INTEGER) AS work_year
            FROM pagamenti_eventi pe
            JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
            JOIN associati a ON a.id = ie.associato_id
            JOIN eventi e ON e.id = ie.evento_id
            LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id
            WHERE pe.gruppo_ricevuta = ?
        ) pagamenti
        ORDER BY scadenza, area, scadenza_riferimento
        """,
        (group_code, group_code, group_code, group_code),
    )
    if not rows:
        return None

    first = dict(rows[0])
    items = [
        {
            "area": row["area"],
            "riferimento": row["scadenza_riferimento"],
            "scadenza": row["scadenza"],
            "importo": float(row["importo"] or 0),
        }
        for row in rows
    ]
    total_amount = sum(item["importo"] for item in items)
    return {
        "id": group_code,
        "payment_type": "multi-area-gruppo",
        "receipt_number": f"RCP-MULTI-{group_code}",
        "data_pagamento": first["data_pagamento"],
        "importo": total_amount,
        "metodo_pagamento": first["metodo_pagamento"],
        "riferimento": first["riferimento"],
        "note": first["note"],
        "codice_associato": first["codice_associato"],
        "associato": first["associato"],
        "email": first["email"],
        "telefono": first["telefono"],
        "area": "Pagamento multi-area",
        "causale": "Saldo contemporaneo di scadenze provenienti da aree diverse",
        "items": items,
        "items_mode": "scadenze",
        "associato_id": int(first["associato_id"]),
        "work_year": int(first["work_year"] or date.today().year),
        "whatsapp_phone": clean_phone_number(first["telefono"]),
    }


def receipt_pending_dues_message(context: dict) -> str:
    associato_id = context.get("associato_id")
    work_year = context.get("work_year")
    if not associato_id or not work_year:
        return ""

    scoped_query = {
        "anno_lavoro": str(work_year),
        "associato_id": str(associato_id),
    }
    definition = build_scadenze_report_definition(scoped_query)
    rows = fetch_all(definition["query"], definition.get("params", ()))
    if not rows:
        return "\n\nAlla data odierna non risultano ulteriori scadenze da saldare."

    lines = ["\n\nSi riporta il riepilogo delle scadenze ancora da saldare:"]
    for row in rows[:8]:
        lines.append(
            f"- {row['area']} | {row['riferimento']} | residuo {money(row['saldo_residuo'])}"
        )
    if len(rows) > 8:
        lines.append(f"- ...e altre {len(rows) - 8} scadenze ancora aperte.")
    return "\n".join(lines)


def associato_scadenze_share_message(
    associato: sqlite3.Row,
    rows: list[sqlite3.Row],
    work_year: int,
) -> str:
    total_residuo = sum((decimal_amount(row["saldo_residuo"]) for row in rows), Decimal("0.00")).quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP,
    )
    lines = [
        f"Buongiorno {plain_text(associato['associato'])},",
        f"si riporta di seguito il riepilogo delle scadenze attualmente aperte per l'anno di lavoro {work_year}.",
    ]
    if rows:
        lines.extend(
            [
                f"Totale residuo: {money(total_residuo)}",
                "Dettaglio scadenze:",
            ]
        )
        for row in rows[:12]:
            lines.append(
                f"- {row['area']} | {row['riferimento']} | scadenza {row['scadenza']} | residuo {money(row['saldo_residuo'])}"
            )
        if len(rows) > 12:
            lines.append(f"- Ulteriori scadenze aperte non riportate in questo messaggio: {len(rows) - 12}.")
    else:
        lines.extend(
            [
                "Alla data odierna non risultano scadenze aperte.",
            ]
        )
    lines.extend(
        [
            "Restiamo a disposizione per eventuali chiarimenti.",
            "Cordiali saluti,",
            APP_NAME,
        ]
    )
    return "\n".join(lines)


def associato_scadenze_share_actions(
    associato: sqlite3.Row,
    rows: list[sqlite3.Row],
    work_year: int,
) -> str:
    message = associato_scadenze_share_message(associato, rows, work_year)
    subject = quote(f"Riepilogo scadenze aperte - {APP_NAME}")
    body = quote(message)
    phone = clean_phone_number(associato["telefono"])

    email_button = (
        f'<a class="button action" href="mailto:{esc(associato["email"])}?subject={subject}&body={body}">Prepara email</a>'
        if associato["email"]
        else '<span class="button action disabled">Email non presente</span>'
    )
    whatsapp_button = (
        f'<a class="button action" href="https://wa.me/{esc(phone)}?text={body}" target="_blank" rel="noopener">Invia WhatsApp</a>'
        if phone
        else '<span class="button action disabled">Cellulare non presente</span>'
    )
    return f'<div class="mini-actions">{email_button}{whatsapp_button}</div>'


def receipt_message(context: dict) -> str:
    closing = "\n\nCordiali saluti,\nOratorio Carlo Acutis"
    base_message = (
        f"Ricevuta {context['receipt_number']} - {APP_NAME}\n"
        f"Associato: {context['associato']}\n"
        f"Causale: {context['causale']}\n"
        f"Data pagamento: {context['data_pagamento']}\n"
        f"Importo: {money(context['importo'])}\n"
        f"Metodo: {context['metodo_pagamento'] or 'Non indicato'}"
    )
    if context.get("items"):
        if context.get("items_mode") == "scadenze":
            details = "\n".join(
                f"- {item['area']} | {item['riferimento']} | {item['scadenza']} ({money(item['importo'])})"
                for item in context["items"]
            )
            return base_message + "\nScadenze saldate:\n" + details + receipt_pending_dues_message(context) + closing

        details = "\n".join(
            f"- {item['corso']} {item['competenza']} ({money(item['importo'])})"
            for item in context["items"]
        )
        return base_message + "\nMensilita saldate:\n" + details + receipt_pending_dues_message(context) + closing
    return base_message + receipt_pending_dues_message(context) + closing


def receipt_toolbar(context: dict) -> str:
    pdf_url = f"/export/pdf/ricevuta/{context['payment_type']}/{context['id']}"
    email_url = ""
    whatsapp_url = ""
    if context.get("email"):
        subject = quote(f"Ricevuta pagamento {context['receipt_number']} - {APP_NAME}")
        body = quote(receipt_message(context))
        email_url = f"mailto:{context['email']}?subject={subject}&body={body}"
    if context.get("whatsapp_phone"):
        whatsapp_url = f"https://wa.me/{context['whatsapp_phone']}?text={quote(receipt_message(context))}"
    email_button = (
        f'<button type="button" class="button action" '
        f'onclick="return shareReceipt(this)" '
        f'data-channel="email" data-mailto="{esc(email_url)}">Prepara email</button>'
        if email_url
        else '<span class="button action disabled">Email non presente</span>'
    )
    whatsapp_button = (
        f'<button type="button" class="button action" '
        f'onclick="return shareReceipt(this)" '
        f'data-channel="whatsapp" data-whatsapp="{esc(whatsapp_url)}">Invia WhatsApp</button>'
        if whatsapp_url
        else '<span class="button action disabled">Cellulare non presente</span>'
    )
    return f"""
    <section class="report-toolbar screen-only">
      <div class="report-toolbar-copy">
        <span class="eyebrow">Azioni ricevuta</span>
        <p>Stampa la ricevuta, esportala in PDF oppure prepara l'invio via email o WhatsApp senza allegato automatico.</p>
      </div>
      <div class="report-toolbar-actions">
        <a class="button action" href="{esc(pdf_url)}">Esporta PDF</a>
        <button type="button" class="button action" onclick="window.print()">Stampa ricevuta</button>
        {email_button}
        {whatsapp_button}
      </div>
    </section>
    """


def receipt_page(
    payment_type: str,
    payment_id: object,
    query_params: dict[str, str],
    current_user: dict[str, object] | None = None,
) -> bytes:
    if payment_type == "corsi-rate-gruppo":
        context = build_grouped_rate_receipt_context(str(payment_id))
    elif payment_type == "multi-area-gruppo":
        context = build_multi_area_receipt_context(str(payment_id))
    else:
        context = build_receipt_context(payment_type, payment_id)
    if context is None:
        raise KeyError(payment_type)

    items_block = ""
    if context.get("items"):
        if context.get("items_mode") == "scadenze":
            items_rows = "".join(
                f"<tr><td>{esc(item['area'])}</td><td>{esc(item['riferimento'])}</td><td>{esc(item['scadenza'])}</td><td>{esc(money(item['importo']))}</td></tr>"
                for item in context["items"]
            )
            items_block = f"""
            <div class="receipt-note">
              <h3>Scadenze saldate</h3>
              <div class="table-wrap">
                <table class="data-table">
                  <thead><tr><th>Area</th><th>Riferimento</th><th>Scadenza</th><th>Importo</th></tr></thead>
                  <tbody>{items_rows}</tbody>
                </table>
              </div>
            </div>
            """
        else:
            items_rows = "".join(
                f"<tr><td>{esc(item['corso'])}</td><td>{esc(item['competenza'])}</td><td>{esc(money(item['importo']))}</td></tr>"
                for item in context["items"]
            )
            items_block = f"""
            <div class="receipt-note">
              <h3>Mensilita saldate</h3>
              <div class="table-wrap">
                <table class="data-table">
                  <thead><tr><th>Corso</th><th>Competenza</th><th>Importo</th></tr></thead>
                  <tbody>{items_rows}</tbody>
                </table>
              </div>
            </div>
            """

    content = f"""
    {receipt_toolbar(context)}
    <section class="card receipt-card">
      <div class="receipt-head">
        <div>
          <span class="eyebrow">Ricevuta di pagamento</span>
          <h2>{esc(APP_NAME)}</h2>
          <p>{esc(context['area'])}</p>
        </div>
        <div class="receipt-number-box">
          <strong>{esc(context['receipt_number'])}</strong>
          <span>Data {esc(context['data_pagamento'])}</span>
        </div>
      </div>
      <div class="receipt-grid">
        <div class="receipt-block">
          <h3>Associato</h3>
          <p><strong>{esc(context['associato'])}</strong></p>
          <p>Codice: {esc(context['codice_associato'])}</p>
          <p>Email: {esc(context['email'] or '-')}</p>
          <p>Cellulare: {esc(context['telefono'] or '-')}</p>
        </div>
        <div class="receipt-block">
          <h3>Dettagli pagamento</h3>
          <p>Causale: {esc(context['causale'])}</p>
          <p>Importo: <strong>{esc(money(context['importo']))}</strong></p>
          <p>Metodo: {esc(context['metodo_pagamento'] or 'Non indicato')}</p>
          <p>Riferimento: {esc(context['riferimento'] or '-')}</p>
        </div>
      </div>
      <div class="receipt-note">
        <h3>Note</h3>
        <p>{esc(context['note'] or 'Nessuna nota inserita.')}</p>
      </div>
      {items_block}
      <div class="receipt-footer">
        <p>Documento generato dal gestionale {esc(APP_NAME)}.</p>
      </div>
    </section>
    """
    return page("Ricevuta pagamento", "/report/incassi", content, query_params, current_user)


def generate_receipt_pdf(context: dict) -> tuple[str, bytes]:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReceiptTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#cb5f07"),
        spaceAfter=4,
    )
    normal_style = ParagraphStyle(
        "ReceiptBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#2c231d"),
    )

    story = []
    logo_path = STATIC_DIR / "logo-ca.jpg"
    header_logo = Image(str(logo_path), width=28 * mm, height=28 * mm) if logo_path.exists() else ""
    header = Table(
        [
            [
                header_logo,
                [
                    Paragraph(APP_NAME, title_style),
                    Paragraph("Ricevuta di pagamento", title_style),
                    Paragraph(f"Numero {context['receipt_number']} - Data {context['data_pagamento']}", normal_style),
                ],
            ]
        ],
        colWidths=[34 * mm, 130 * mm],
    )
    header.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(header)
    story.append(Spacer(1, 8 * mm))

    details = [
        ["Associato", context["associato"]],
        ["Codice associato", context["codice_associato"]],
        ["Area", context["area"]],
        ["Causale", context["causale"]],
        ["Importo", money(context["importo"])],
        ["Metodo pagamento", context["metodo_pagamento"] or "Non indicato"],
        ["Riferimento", context["riferimento"] or "-"],
        ["Email", context["email"] or "-"],
        ["Cellulare", context["telefono"] or "-"],
        ["Note", context["note"] or "Nessuna nota inserita."],
    ]
    detail_table = Table(details, colWidths=[45 * mm, 115 * mm])
    detail_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#fff3e7")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#cb5f07")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#f0d2b8")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(detail_table)
    if context.get("items"):
        story.append(Spacer(1, 6 * mm))
        if context.get("items_mode") == "scadenze":
            story.append(Paragraph("Scadenze saldate", title_style))
            item_rows = [["Area", "Riferimento", "Scadenza", "Importo"]] + [
                [item["area"], item["riferimento"], item["scadenza"], money(item["importo"])] for item in context["items"]
            ]
            item_table = Table(item_rows, colWidths=[40 * mm, 60 * mm, 30 * mm, 30 * mm])
        else:
            story.append(Paragraph("Mensilita saldate", title_style))
            item_rows = [["Corso", "Competenza", "Importo"]] + [
                [item["corso"], item["competenza"], money(item["importo"])] for item in context["items"]
            ]
            item_table = Table(item_rows, colWidths=[75 * mm, 35 * mm, 50 * mm])
        item_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ef7f1a")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#f0d2b8")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fff7f0")]),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.append(item_table)
    story.append(Spacer(1, 8 * mm))
    story.append(Paragraph("Documento generato dal gestionale dell'Oratorio Carlo Acutis.", normal_style))

    document.build(story)
    return f"ricevuta_{context['receipt_number'].lower()}.pdf", buffer.getvalue()


def export_receipt_pdf(start_response, payment_type: str, payment_id: str):
    if payment_type == "corsi-rate-gruppo":
        context = build_grouped_rate_receipt_context(payment_id)
    elif payment_type == "multi-area-gruppo":
        context = build_multi_area_receipt_context(payment_id)
    else:
        context = build_receipt_context(payment_type, int(payment_id))
    if context is None:
        return not_found(start_response)
    filename, content = generate_receipt_pdf(context)
    start_response(
        "200 OK",
        [
            ("Content-Type", "application/pdf"),
            ("Content-Disposition", f'attachment; filename="{filename}"'),
            ("Content-Length", str(len(content))),
        ],
    )
    return [content]


def handle_post(
    path: str,
    form_data: dict[str, str],
    start_response,
    current_user: dict[str, object] | None = None,
    request_cookies: dict[str, str] | None = None,
):
    context_query = work_year_query_from_form(form_data)
    if path in {"/azioni/accesso/primo-admin", "/azioni/accesso/login", "/azioni/accesso/recupera-password"}:
        next_target = login_destination(normalized(form_data, "next", "/"))
        if next_target and next_target != "/":
            context_query["next"] = next_target
    try:
        if path == "/azioni/accesso/primo-admin":
            if not bootstrap_admin_required():
                raise ValueError("L'amministratore iniziale e gia stato configurato.")
            username = validate_username(required(form_data, "username", "Username"))
            password = validate_password(
                required(form_data, "password", "Password"),
                required(form_data, "password_conferma", "Conferma password"),
            )
            with get_connection() as connection:
                user_id = create_access_user(connection, username=username, password=password, is_admin=True)
                session_token = create_user_session(connection, user_id)
                connection.commit()
            return redirect(
                start_response,
                "/",
                ok="Amministratore creato e accesso effettuato.",
                extra_headers=[session_cookie_header(session_token)],
            )

        if path == "/azioni/accesso/login":
            if bootstrap_admin_required():
                raise ValueError("Completa prima la creazione dell'amministratore.")
            username = validate_username(required(form_data, "username", "Username"))
            password = required(form_data, "password", "Password")
            next_target = login_destination(normalized(form_data, "next", "/"))
            with get_connection() as connection:
                cleanup_expired_sessions(connection)
                user_row = connection.execute(
                    """
                    SELECT *
                    FROM utenti_accesso
                    WHERE username = ?
                    """,
                    (username,),
                ).fetchone()
                if user_row is None or not verify_password(password, user_row):
                    raise ValueError("Username o password non validi.")
                if not user_row["attivo"]:
                    raise ValueError("Questo utente non e attivo.")
                session_token = create_user_session(connection, int(user_row["id"]))
                connection.commit()
            return redirect(
                start_response,
                next_target,
                ok="Accesso effettuato.",
                extra_headers=[session_cookie_header(session_token)],
            )

        if path == "/azioni/accesso/logout":
            session_token = (request_cookies or {}).get(SESSION_COOKIE_NAME, "")
            if session_token:
                with get_connection() as connection:
                    connection.execute("DELETE FROM sessioni_accesso WHERE session_token = ?", (session_token,))
                    connection.commit()
            return redirect(
                start_response,
                "/login",
                ok="Sessione chiusa.",
                extra_headers=[clear_session_cookie_header()],
            )

        if path == "/azioni/accesso/cambia-password":
            if current_user is None:
                return redirect(start_response, "/login", err="Accedi per modificare la password.")
            current_password = required(form_data, "password_attuale", "Password attuale")
            new_password = validate_password(
                required(form_data, "password", "Nuova password"),
                required(form_data, "password_conferma", "Conferma nuova password"),
            )
            with get_connection() as connection:
                user_row = connection.execute(
                    "SELECT * FROM utenti_accesso WHERE id = ?",
                    (int(current_user["id"]),),
                ).fetchone()
                if user_row is None or not verify_password(current_password, user_row):
                    raise ValueError("La password attuale non e corretta.")
                set_access_user_password(connection, int(current_user["id"]), new_password)
                connection.commit()
            return redirect(start_response, "/maschere/accesso", ok="Password aggiornata.", extra_query=context_query)

        if path == "/azioni/accesso/recupera-password":
            if bootstrap_admin_required():
                raise ValueError("Configura prima l'utente amministratore.")
            username = validate_username(required(form_data, "username", "Username"))
            recovery_email = validate_recovery_email(required(form_data, "email_recupero", "Email recupero"))
            new_password = validate_password(
                required(form_data, "password", "Nuova password"),
                required(form_data, "password_conferma", "Conferma nuova password"),
            )
            with get_connection() as connection:
                user_row = connection.execute(
                    """
                    SELECT *
                    FROM utenti_accesso
                    WHERE username = ?
                      AND is_admin = 0
                      AND attivo = 1
                    """,
                    (username,),
                ).fetchone()
                if user_row is None:
                    raise ValueError("Utente standard non trovato.")
                stored_email = (user_row["email_recupero"] or "").strip().lower()
                if stored_email != recovery_email.strip().lower():
                    raise ValueError("Email di recupero non corrispondente.")
                set_access_user_password(connection, int(user_row["id"]), new_password)
                connection.commit()
            return redirect(start_response, "/login", ok="Password aggiornata. Ora puoi accedere.", extra_query=context_query)

        if path == "/azioni/utenti/crea":
            if not current_user or not current_user.get("is_admin"):
                return redirect(
                    start_response,
                    "/",
                    err="Solo l'amministratore puo creare nuovi utenti.",
                    extra_query=context_query,
                )
            username = validate_username(required(form_data, "username", "Username"))
            password = validate_password(
                required(form_data, "password", "Password"),
                required(form_data, "password_conferma", "Conferma password"),
            )
            email_recupero = validate_recovery_email(required(form_data, "email_recupero", "Email recupero"))
            with get_connection() as connection:
                create_access_user(
                    connection,
                    username=username,
                    password=password,
                    is_admin=False,
                    email_recupero=email_recupero,
                )
                connection.commit()
            return redirect(start_response, "/maschere/utenti", ok="Utente creato.", extra_query=context_query)

        if path.startswith("/azioni/utenti/aggiorna/"):
            if not current_user or not current_user.get("is_admin"):
                return redirect(start_response, "/", err="Solo l'amministratore puo modificare gli utenti.", extra_query=context_query)
            user_id = int(path.removeprefix("/azioni/utenti/aggiorna/"))
            user_row = access_user_row(user_id)
            if user_row is None or user_row["is_admin"]:
                raise ValueError("Puoi gestire solo utenti standard.")
            with get_connection() as connection:
                connection.execute(
                    """
                    UPDATE utenti_accesso
                    SET username = ?, email_recupero = ?, aggiornato_il = ?
                    WHERE id = ?
                    """,
                    (
                        validate_username(required(form_data, "username", "Username")),
                        validate_recovery_email(required(form_data, "email_recupero", "Email recupero")),
                        current_timestamp(),
                        user_id,
                    ),
                )
                connection.commit()
            return redirect(start_response, f"/maschere/utenti/gestione/{user_id}", ok="Utente aggiornato.", extra_query=context_query)

        if path.startswith("/azioni/utenti/password/"):
            if not current_user or not current_user.get("is_admin"):
                return redirect(start_response, "/", err="Solo l'amministratore puo reimpostare le password.", extra_query=context_query)
            user_id = int(path.removeprefix("/azioni/utenti/password/"))
            user_row = access_user_row(user_id)
            if user_row is None or user_row["is_admin"]:
                raise ValueError("Puoi reimpostare solo password di utenti standard.")
            new_password = validate_password(
                required(form_data, "password", "Nuova password"),
                required(form_data, "password_conferma", "Conferma nuova password"),
            )
            with get_connection() as connection:
                set_access_user_password(connection, user_id, new_password)
                connection.commit()
            return redirect(start_response, f"/maschere/utenti/gestione/{user_id}", ok="Password utente aggiornata.", extra_query=context_query)

        if path.startswith("/azioni/utenti/stato/"):
            if not current_user or not current_user.get("is_admin"):
                return redirect(start_response, "/", err="Solo l'amministratore puo disattivare o riattivare gli utenti.", extra_query=context_query)
            user_id = int(path.removeprefix("/azioni/utenti/stato/"))
            user_row = access_user_row(user_id)
            if user_row is None or user_row["is_admin"]:
                raise ValueError("Puoi modificare lo stato solo degli utenti standard.")
            desired_action = normalized(form_data, "azione_stato", "")
            if desired_action not in {"disattiva", "riattiva"}:
                raise ValueError("Azione stato non valida.")
            attivo_value = 0 if desired_action == "disattiva" else 1
            with get_connection() as connection:
                connection.execute(
                    """
                    UPDATE utenti_accesso
                    SET attivo = ?, aggiornato_il = ?
                    WHERE id = ?
                    """,
                    (attivo_value, current_timestamp(), user_id),
                )
                if attivo_value == 0:
                    connection.execute("DELETE FROM sessioni_accesso WHERE utente_id = ?", (user_id,))
                connection.commit()
            message = "Utente standard disattivato." if attivo_value == 0 else "Utente standard riattivato."
            return redirect(start_response, f"/maschere/utenti/gestione/{user_id}", ok=message, extra_query=context_query)

        if path.startswith("/azioni/crud/aggiorna/"):
            entity_key, record_id = path.removeprefix("/azioni/crud/aggiorna/").split("/", 1)
            return handle_crud_update(entity_key, int(record_id), form_data, start_response, current_user)

        if path.startswith("/azioni/crud/elimina/"):
            entity_key, record_id = path.removeprefix("/azioni/crud/elimina/").split("/", 1)
            return handle_crud_delete(entity_key, int(record_id), form_data, start_response)

        if path == "/azioni/associati/crea":
            with get_connection() as connection:
                progressive_number = reserve_progressive_number(connection, "associati")
                connection.execute(
                    """
                    INSERT INTO associati (
                        numero_progressivo, codice_associato, nome, cognome, codice_fiscale, data_nascita,
                        sesso, comune_nascita, provincia_nascita, carica, email, telefono, indirizzo, cap, citta, provincia,
                        data_prima_iscrizione, stato_associato, note
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        progressive_number,
                        format_progressive_code("associati", progressive_number),
                        required(form_data, "nome", "Nome"),
                        required(form_data, "cognome", "Cognome"),
                        optional(form_data, "codice_fiscale"),
                        optional(form_data, "data_nascita"),
                        normalized(form_data, "sesso", "M") or "M",
                        optional(form_data, "comune_nascita"),
                        optional(form_data, "provincia_nascita"),
                        resolved_carica_value(form_data, current_user, existing_value="Associato"),
                        optional(form_data, "email"),
                        optional(form_data, "telefono"),
                        optional(form_data, "indirizzo"),
                        optional(form_data, "cap"),
                        optional(form_data, "citta"),
                        optional(form_data, "provincia"),
                        required(form_data, "data_prima_iscrizione", "Data prima iscrizione"),
                        normalized(form_data, "stato_associato", "Attivo") or "Attivo",
                        optional(form_data, "note"),
                    ),
                )
                connection.commit()
            return redirect(start_response, "/maschere/associati", ok="Associato salvato.", extra_query=context_query)

        if path == "/azioni/quote/crea":
            area = required(form_data, "area", "Area")
            if area not in {"tesseramenti", "campi-estivi"}:
                raise ValueError("Area quota non valida.")
            execute(
                """
                INSERT INTO quote_predefinite (
                    area, descrizione, importo, note
                ) VALUES (?, ?, ?, ?)
                """,
                (
                    area,
                    required(form_data, "descrizione", "Descrizione"),
                    required(form_data, "importo", "Importo"),
                    optional(form_data, "note"),
                ),
            )
            redirect_path = "/maschere/tesseramenti" if area == "tesseramenti" else "/maschere/campi-estivi"
            success_message = "Quota tesseramento salvata." if area == "tesseramenti" else "Quota Campo estivo salvata."
            return redirect(start_response, redirect_path, ok=success_message, extra_query=context_query)

        if path == "/azioni/tesseramenti/crea":
            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO tesseramenti_annuali (
                        associato_id, anno_sociale, data_tesseramento, importo_dovuto, data_scadenza, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "associato_id", "Associato"),
                        required(form_data, "anno_sociale", "Anno sociale"),
                        required(form_data, "data_tesseramento", "Data tesseramento"),
                        required(form_data, "importo_dovuto", "Importo dovuto"),
                        optional(form_data, "data_scadenza"),
                        optional(form_data, "note"),
                    ),
                )
                tesseramento_id = int(cursor.lastrowid)

                if popup_payment_requested(form_data):
                    metodo_pagamento_id, importo_pagato, data_pagamento = popup_payment_payload(
                        form_data,
                        optional(form_data, "data_tesseramento") or date.today().isoformat(),
                    )
                    importo_dovuto = decimal_amount(required(form_data, "importo_dovuto", "Importo dovuto"), minimum="0")
                    if importo_pagato > importo_dovuto:
                        raise ValueError("L'importo pagato non puo superare l'importo dovuto del tesseramento.")
                    payment_cursor = connection.execute(
                        """
                        INSERT INTO pagamenti_tesseramenti (
                            tesseramento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            tesseramento_id,
                            data_pagamento,
                            format(importo_pagato, ".2f"),
                            metodo_pagamento_id,
                            "",
                            optional(form_data, "note"),
                        ),
                    )
                    payment_id = int(payment_cursor.lastrowid)
                    connection.commit()
                    return redirect(
                        start_response,
                        f"/ricevute/tesseramenti/{payment_id}",
                        ok="Tesseramento e pagamento registrati.",
                        extra_query=context_query,
                    )

                connection.commit()
            return redirect(start_response, "/maschere/tesseramenti", ok="Tesseramento salvato.", extra_query=context_query)

        if path == "/azioni/tesseramenti/pagamento":
            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO pagamenti_tesseramenti (
                        tesseramento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "tesseramento_id", "Tesseramento"),
                        required(form_data, "data_pagamento", "Data pagamento"),
                        required(form_data, "importo", "Importo"),
                        required(form_data, "metodo_pagamento_id", "Metodo"),
                        optional(form_data, "riferimento"),
                        optional(form_data, "note"),
                    ),
                )
                payment_id = cursor.lastrowid
                connection.commit()
            return redirect(
                start_response,
                f"/ricevute/tesseramenti/{payment_id}",
                ok="Pagamento tesseramento registrato.",
                extra_query=context_query,
            )

        if path == "/azioni/corsi/tipologia":
            raise ValueError("Le tipologie corsi sono fisse e non e possibile inserirne di nuove.")

        if path == "/azioni/corsi/crea":
            with get_connection() as connection:
                progressive_number = reserve_progressive_number(connection, "corsi")
                connection.execute(
                    """
                    INSERT INTO corsi (
                        numero_progressivo, codice_corso, nome, tipologia_corso_id, descrizione, quota_iscrizione_standard,
                        quota_mensile_standard, sede, giorno_settimana, orario
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        progressive_number,
                        format_progressive_code("corsi", progressive_number),
                        required(form_data, "nome", "Nome corso"),
                        None,
                        optional(form_data, "descrizione"),
                        "0",
                        normalized(form_data, "quota_mensile_standard", "0"),
                        optional(form_data, "sede"),
                        optional(form_data, "giorno_settimana"),
                        optional(form_data, "orario"),
                    ),
                )
                connection.commit()
            return redirect(start_response, "/maschere/corsi", ok="Corso salvato.", extra_query=context_query)

        if path == "/azioni/corsi/iscrizione":
            data_iscrizione_value = required(form_data, "data_iscrizione", "Data iscrizione")
            try:
                data_iscrizione_date = date.fromisoformat(data_iscrizione_value)
            except ValueError:
                raise ValueError("La data di iscrizione del corso non e valida.")

            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO iscrizioni_corsi (
                        associato_id, corso_id, data_iscrizione, data_inizio, quota_iscrizione,
                        quota_mensile, stato_iscrizione, note
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "associato_id", "Associato"),
                        required(form_data, "corso_id", "Corso"),
                        data_iscrizione_value,
                        optional(form_data, "data_inizio"),
                        "0",
                        required(form_data, "quota_mensile", "Quota mensile"),
                        normalized(form_data, "stato_iscrizione", "Attiva") or "Attiva",
                        optional(form_data, "note"),
                    ),
                )
                iscrizione_id = int(cursor.lastrowid)
                quote_note = "Quota generata automaticamente al salvataggio dell'iscrizione"
                payment_scope = normalized(form_data, "pagamento_scope", "mese-iscrizione") or "mese-iscrizione"
                rate_ids = [ensure_course_rate_for_enrollment(
                    connection,
                    iscrizione_id,
                    data_iscrizione_date.year,
                    data_iscrizione_date.month,
                    note=quote_note,
                )]

                if popup_payment_requested(form_data):
                    metodo_pagamento_id, importo_pagato, data_pagamento = popup_payment_payload(
                        form_data,
                        data_iscrizione_value or date.today().isoformat(),
                    )
                    if payment_scope == "mensilita-future":
                        fine_competenza = required(form_data, "pagamento_competenza_fine", "Ultimo mese da pagare")
                        end_year, end_month = parse_year_month_value(fine_competenza, "Ultimo mese da pagare")
                        rate_ids = ensure_course_rates_for_enrollment_range(
                            connection,
                            iscrizione_id,
                            data_iscrizione_date.year,
                            data_iscrizione_date.month,
                            end_year,
                            end_month,
                            note=quote_note,
                        )

                    placeholders = ",".join("?" for _ in rate_ids)
                    rate_rows = connection.execute(
                        f"""
                        SELECT
                            r.id,
                            r.anno,
                            r.mese,
                            r.importo_dovuto,
                            COALESCE((
                                SELECT SUM(prc.importo)
                                FROM pagamenti_rate_corsi prc
                                WHERE prc.rata_corso_id = r.id
                            ), 0) AS importo_pagato
                        FROM rate_corsi_mensili r
                        WHERE r.id IN ({placeholders})
                        ORDER BY r.anno, r.mese, r.id
                        """,
                        tuple(rate_ids),
                    ).fetchall()
                    if len(rate_rows) != len(rate_ids):
                        raise ValueError("Una o piu quote mensili da pagare non sono piu disponibili.")

                    residui: list[tuple[int, Decimal]] = []
                    totale_residuo = Decimal("0.00")
                    for row in rate_rows:
                        residuo_rata = (
                            decimal_amount(row["importo_dovuto"]) - decimal_amount(row["importo_pagato"])
                        ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                        residui.append((int(row["id"]), residuo_rata))
                        totale_residuo = (totale_residuo + residuo_rata).quantize(
                            Decimal("0.01"),
                            rounding=ROUND_HALF_UP,
                        )

                    if importo_pagato > totale_residuo:
                        if payment_scope == "mensilita-future":
                            raise ValueError("L'importo pagato non puo superare il totale delle mensilita selezionate.")
                        raise ValueError("L'importo pagato non puo superare la quota mensile del mese di iscrizione.")

                    if payment_scope == "mensilita-future":
                        group_code = generate_receipt_group_code()
                        remaining = importo_pagato
                        for rate_id, residuo_rata in residui:
                            if remaining <= Decimal("0.00"):
                                break
                            importo_rata = min(remaining, residuo_rata).quantize(
                                Decimal("0.01"),
                                rounding=ROUND_HALF_UP,
                            )
                            if importo_rata <= Decimal("0.00"):
                                continue
                            connection.execute(
                                """
                                INSERT INTO pagamenti_rate_corsi (
                                    rata_corso_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    rate_id,
                                    data_pagamento,
                                    format(importo_rata, ".2f"),
                                    metodo_pagamento_id,
                                    "",
                                    optional(form_data, "note"),
                                    group_code,
                                ),
                            )
                            remaining = (remaining - importo_rata).quantize(
                                Decimal("0.01"),
                                rounding=ROUND_HALF_UP,
                            )
                        connection.commit()
                        return redirect(
                            start_response,
                            f"/ricevute/corsi-rate-gruppo/{group_code}",
                            ok="Iscrizione corso, quote mensili e pagamento registrati.",
                            extra_query=context_query,
                        )

                    payment_cursor = connection.execute(
                        """
                        INSERT INTO pagamenti_rate_corsi (
                            rata_corso_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            rate_ids[0],
                            data_pagamento,
                            format(importo_pagato, ".2f"),
                            metodo_pagamento_id,
                            "",
                            optional(form_data, "note"),
                        ),
                    )
                    payment_id = int(payment_cursor.lastrowid)
                    connection.commit()
                    return redirect(
                        start_response,
                        f"/ricevute/corsi-rata/{payment_id}",
                        ok="Iscrizione corso, quota mensile e pagamento registrati.",
                        extra_query=context_query,
                    )

                connection.commit()
            return redirect(
                start_response,
                "/maschere/corsi",
                ok="Iscrizione corso salvata. La quota mensile del mese di iscrizione e stata generata automaticamente.",
                extra_query=context_query,
            )

        if path == "/azioni/corsi/pagamento-iscrizione":
            raise ValueError("Il pagamento di iscrizione al corso e stato disattivato.")

        if path == "/azioni/corsi/rata":
            raise ValueError("L'inserimento della singola quota mensile e disattivato. Le quote corsi vengono generate automaticamente dal gestionale.")

        if path in ("/azioni/corsi/pagamento-rata", "/azioni/corsi/pagamento-rate"):
            rate_ids = multi_values(form_data, "rata_corso_id")
            if not rate_ids:
                raise ValueError("Seleziona almeno una mensilita da saldare.")

            associato_id = required(form_data, "associato_id", "Associato")
            importo_totale = decimal_amount(required(form_data, "importo", "Importo"), minimum="0.01")
            placeholders = ",".join("?" for _ in rate_ids)
            rows = fetch_all(
                f"""
                SELECT id, associato_id, corso, competenza, saldo_residuo
                FROM v_rate_corsi_saldo
                WHERE id IN ({placeholders})
                ORDER BY anno, mese, corso
                """,
                tuple(int(rate_id) for rate_id in rate_ids),
            )
            if len(rows) != len(rate_ids):
                raise ValueError("Una o piu mensilita selezionate non sono piu disponibili.")

            selected_associati = {str(row["associato_id"]) for row in rows}
            if selected_associati != {associato_id}:
                raise ValueError("Le mensilita selezionate devono appartenere tutte allo stesso associato.")

            if any(float(row["saldo_residuo"] or 0) <= 0 for row in rows):
                raise ValueError("Una o piu mensilita selezionate risultano gia saldate.")

            totale_residuo = sum(decimal_amount(row["saldo_residuo"]) for row in rows)
            if importo_totale > totale_residuo:
                raise ValueError("L'importo inserito supera il residuo totale delle mensilita selezionate.")

            group_code = generate_receipt_group_code()
            remaining = importo_totale
            with get_connection() as connection:
                for row in rows:
                    residuo_rata = decimal_amount(row["saldo_residuo"])
                    if remaining <= Decimal("0.00"):
                        break
                    importo_rata = min(remaining, residuo_rata).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    if importo_rata <= Decimal("0.00"):
                        continue
                    connection.execute(
                        """
                        INSERT INTO pagamenti_rate_corsi (
                            rata_corso_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            row["id"],
                            required(form_data, "data_pagamento", "Data pagamento"),
                            format(importo_rata, ".2f"),
                            required(form_data, "metodo_pagamento_id", "Metodo"),
                            optional(form_data, "riferimento"),
                            optional(form_data, "note"),
                            group_code,
                        ),
                    )
                    remaining = (remaining - importo_rata).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                connection.commit()

            return redirect(
                start_response,
                f"/ricevute/corsi-rate-gruppo/{group_code}",
                ok="Pagamento quote mensili registrato.",
                extra_query=context_query,
            )

        if path == "/azioni/corsi/genera-rate":
            anno = int(required(form_data, "anno", "Anno"))
            mese = int(required(form_data, "mese", "Mese"))
            data_scadenza = optional(form_data, "data_scadenza")
            note = optional(form_data, "note")
            with get_connection() as connection:
                inserted, skipped = generate_course_rates_for_month(
                    connection,
                    anno,
                    mese,
                    data_scadenza=data_scadenza,
                    note=note,
                )
                connection.commit()

            message = (
                f"Sono state generate {inserted} quote mensili per {month_label(mese)} {anno}."
                if inserted
                else f"Nessuna nuova quota da generare per {month_label(mese)} {anno}."
            )
            if skipped:
                message += f" {skipped} gia presenti."
            return redirect(start_response, "/maschere/corsi", ok=message, extra_query=context_query)

        if path == "/azioni/corsi/genera-rate-mancanti":
            work_year = int(required(form_data, "anno", "Anno"))
            reminder = pending_course_monthly_generations(work_year)
            if reminder is None:
                return redirect(
                    start_response,
                    "/",
                    ok="Non risultano quote mensili mancanti da generare.",
                    extra_query=context_query or {"anno_lavoro": str(work_year)},
                )

            inserted_total = 0
            skipped_total = 0
            generated_labels: list[str] = []
            with get_connection() as connection:
                for anno, mese in reminder["missing_months"]:
                    inserted, skipped = generate_course_rates_for_month(
                        connection,
                        int(anno),
                        int(mese),
                        data_scadenza=default_mass_rate_due_date(int(anno), int(mese)),
                        note="Generazione automatica mesi mancanti",
                    )
                    inserted_total += inserted
                    skipped_total += skipped
                    generated_labels.append(f"{month_label(int(mese))} {int(anno)}")
                connection.commit()

            context = context_query or {"anno_lavoro": str(work_year)}
            message = (
                f"Generazione automatica completata per: {', '.join(generated_labels)}. "
                f"Nuove quote create: {inserted_total}."
            )
            if skipped_total:
                message += f" Quote gia presenti: {skipped_total}."
            return redirect(start_response, "/", ok=message, extra_query=context)

        if path == "/azioni/campi-estivi/crea":
            year = int(required(form_data, "anno", "Anno"))
            existing_camp = fetch_scalar("SELECT id FROM campi_estivi WHERE anno = ?", (year,))
            if existing_camp:
                raise ValueError("Per ogni anno di lavoro e previsto un solo Campo estivo.")
            with get_connection() as connection:
                progressive_number = reserve_progressive_number(connection, "campi_estivi")
                connection.execute(
                    """
                    INSERT INTO campi_estivi (
                        numero_progressivo, codice_campo, nome, anno, data_inizio, data_fine, sede,
                        quota_partecipazione_standard, posti_massimi, descrizione
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        progressive_number,
                        format_progressive_code("campi_estivi", progressive_number),
                        required(form_data, "nome", "Nome"),
                        year,
                        required(form_data, "data_inizio", "Data inizio"),
                        required(form_data, "data_fine", "Data fine"),
                        optional(form_data, "sede"),
                        normalized(form_data, "quota_partecipazione_standard", "0"),
                        optional(form_data, "posti_massimi"),
                        optional(form_data, "descrizione"),
                    ),
                )
                connection.commit()
            return redirect(start_response, "/maschere/campi-estivi", ok="Campo estivo salvato.", extra_query=context_query)

        if path == "/azioni/campi-estivi/quota-standard":
            work_year = current_work_year(context_query)
            standard_fee = required(form_data, "quota_partecipazione_standard", "Quota di partecipazione standard")
            with get_connection() as connection:
                ensure_estate_record(connection, work_year, standard_fee)
                connection.commit()
            return redirect(
                start_response,
                "/maschere/campi-estivi",
                ok="Quota standard Campo estivo aggiornata.",
                extra_query=context_query,
            )

        if path == "/azioni/campi-estivi/iscrizione":
            work_year = current_work_year(context_query)
            with get_connection() as connection:
                estate_id = connection.execute(
                    "SELECT id FROM campi_estivi WHERE anno = ? ORDER BY id LIMIT 1",
                    (work_year,),
                ).fetchone()
                campo_estivo_id = (
                    int(estate_id["id"])
                    if estate_id is not None
                    else ensure_estate_record(connection, work_year)
                )
                cursor = connection.execute(
                    """
                    INSERT INTO iscrizioni_campi_estivi (
                        associato_id, campo_estivo_id, data_iscrizione, quota_partecipazione, stato_iscrizione, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "associato_id", "Associato"),
                        campo_estivo_id,
                        required(form_data, "data_iscrizione", "Data iscrizione"),
                        required(form_data, "quota_partecipazione", "Quota partecipazione"),
                        normalized(form_data, "stato_iscrizione", "Iscritto") or "Iscritto",
                        optional(form_data, "note"),
                    ),
                )
                iscrizione_id = int(cursor.lastrowid)

                if popup_payment_requested(form_data):
                    metodo_pagamento_id, importo_pagato, data_pagamento = popup_payment_payload(
                        form_data,
                        required(form_data, "data_iscrizione", "Data iscrizione"),
                    )
                    quota_dovuta = decimal_amount(required(form_data, "quota_partecipazione", "Importo"), minimum="0")
                    if importo_pagato > quota_dovuta:
                        raise ValueError("L'importo pagato non puo superare l'importo dovuto del Campo estivo.")
                    payment_cursor = connection.execute(
                        """
                        INSERT INTO pagamenti_campi_estivi (
                            iscrizione_campo_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            iscrizione_id,
                            data_pagamento,
                            format(importo_pagato, ".2f"),
                            metodo_pagamento_id,
                            "",
                            optional(form_data, "note"),
                        ),
                    )
                    payment_id = int(payment_cursor.lastrowid)
                    connection.commit()
                    return redirect(
                        start_response,
                        f"/ricevute/campi-estivi/{payment_id}",
                        ok="Iscrizione e pagamento Campo estivo registrati.",
                        extra_query=context_query,
                    )
                connection.commit()
            return redirect(start_response, "/maschere/campi-estivi", ok="Iscrizione Campo estivo salvata.", extra_query=context_query)

        if path == "/azioni/campi-estivi/pagamento":
            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO pagamenti_campi_estivi (
                        iscrizione_campo_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "iscrizione_campo_id", "Iscrizione Campo estivo"),
                        required(form_data, "data_pagamento", "Data pagamento"),
                        required(form_data, "importo", "Importo"),
                        required(form_data, "metodo_pagamento_id", "Metodo"),
                        optional(form_data, "riferimento"),
                        optional(form_data, "note"),
                    ),
                )
                payment_id = cursor.lastrowid
                connection.commit()
            return redirect(
                start_response,
                f"/ricevute/campi-estivi/{payment_id}",
                ok="Pagamento Campo estivo registrato.",
                extra_query=context_query,
            )

        if path == "/azioni/eventi/crea":
            with get_connection() as connection:
                progressive_number = reserve_progressive_number(connection, "eventi")
                connection.execute(
                    """
                    INSERT INTO eventi (
                        numero_progressivo, codice_evento, nome, tipologia, data_evento, luogo,
                        quota_partecipazione_standard, posti_massimi, descrizione
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        progressive_number,
                        format_progressive_code("eventi", progressive_number),
                        required(form_data, "nome", "Nome evento"),
                        optional(form_data, "tipologia"),
                        required(form_data, "data_evento", "Data evento"),
                        optional(form_data, "luogo"),
                        normalized(form_data, "quota_partecipazione_standard", "0"),
                        optional(form_data, "posti_massimi"),
                        optional(form_data, "descrizione"),
                    ),
                )
                connection.commit()
            return redirect(start_response, "/maschere/eventi", ok="Evento salvato.", extra_query=context_query)

        if path == "/azioni/eventi/iscrizione":
            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO iscrizioni_eventi (
                        associato_id, evento_id, data_iscrizione, quota_partecipazione, stato_iscrizione, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "associato_id", "Associato"),
                        required(form_data, "evento_id", "Evento"),
                        required(form_data, "data_iscrizione", "Data iscrizione"),
                        required(form_data, "quota_partecipazione", "Quota partecipazione"),
                        normalized(form_data, "stato_iscrizione", "Iscritto") or "Iscritto",
                        optional(form_data, "note"),
                    ),
                )
                iscrizione_id = int(cursor.lastrowid)

                if popup_payment_requested(form_data):
                    metodo_pagamento_id, importo_pagato, data_pagamento = popup_payment_payload(
                        form_data,
                        required(form_data, "data_iscrizione", "Data iscrizione"),
                    )
                    quota_dovuta = decimal_amount(required(form_data, "quota_partecipazione", "Quota dovuta"), minimum="0")
                    if importo_pagato > quota_dovuta:
                        raise ValueError("L'importo pagato non puo superare l'importo dovuto dell'evento.")
                    payment_cursor = connection.execute(
                        """
                        INSERT INTO pagamenti_eventi (
                            iscrizione_evento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            iscrizione_id,
                            data_pagamento,
                            format(importo_pagato, ".2f"),
                            metodo_pagamento_id,
                            "",
                            optional(form_data, "note"),
                        ),
                    )
                    payment_id = int(payment_cursor.lastrowid)
                    connection.commit()
                    return redirect(
                        start_response,
                        f"/ricevute/eventi/{payment_id}",
                        ok="Iscrizione e pagamento evento registrati.",
                        extra_query=context_query,
                    )
                connection.commit()
            return redirect(start_response, "/maschere/eventi", ok="Iscrizione evento salvata.", extra_query=context_query)

        if path == "/azioni/eventi/pagamento":
            with get_connection() as connection:
                cursor = connection.execute(
                    """
                    INSERT INTO pagamenti_eventi (
                        iscrizione_evento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        required(form_data, "iscrizione_evento_id", "Iscrizione evento"),
                        required(form_data, "data_pagamento", "Data pagamento"),
                        required(form_data, "importo", "Importo"),
                        required(form_data, "metodo_pagamento_id", "Metodo"),
                        optional(form_data, "riferimento"),
                        optional(form_data, "note"),
                    ),
                )
                payment_id = cursor.lastrowid
                connection.commit()
            return redirect(
                start_response,
                f"/ricevute/eventi/{payment_id}",
                ok="Pagamento evento registrato.",
                extra_query=context_query,
            )

        if path == "/azioni/pagamenti-multi-area/crea":
            scadenza_tokens = multi_values(form_data, "scadenza_id")
            if not scadenza_tokens:
                raise ValueError("Seleziona almeno una scadenza da saldare.")

            associato_id = required(form_data, "associato_id", "Associato")
            importo_totale = decimal_amount(required(form_data, "importo", "Importo"), minimum="0.01")
            rows = load_multi_area_scadenze(scadenza_tokens)

            selected_associati = {str(row["associato_id"]) for row in rows}
            if selected_associati != {associato_id}:
                raise ValueError("Le scadenze selezionate devono appartenere tutte allo stesso associato.")

            if any(decimal_amount(row["saldo_residuo"]) <= Decimal("0.00") for row in rows):
                raise ValueError("Una o piu scadenze selezionate risultano gia saldate.")

            one_time_kinds = {"campi-estivi", "eventi"}
            for row in rows:
                if row["kind"] in one_time_kinds and decimal_amount(row["importo_pagato"]) > Decimal("0.00"):
                    raise ValueError(
                        f"La scadenza {row['area']} - {row['riferimento']} ha gia un pagamento parziale e non puo essere inclusa nel pagamento multi-area."
                    )

            totale_residuo = sum(decimal_amount(row["saldo_residuo"]) for row in rows)
            if importo_totale > totale_residuo:
                raise ValueError("L'importo inserito supera il residuo totale delle scadenze selezionate.")

            ordered_rows = []
            for index, row in enumerate(rows):
                priority = 1 if row["kind"] in one_time_kinds else 0
                ordered_rows.append((priority, index, row))
            ordered_rows.sort(key=lambda item: (item[0], item[1]))

            group_code = generate_multi_area_group_code()
            remaining = importo_totale
            inserted = 0
            with get_connection() as connection:
                for _, _, row in ordered_rows:
                    residuo = decimal_amount(row["saldo_residuo"])
                    if remaining <= Decimal("0.00"):
                        break

                    if row["kind"] in one_time_kinds and remaining < residuo:
                        raise ValueError(
                            f"La scadenza {row['area']} - {row['riferimento']} puo essere inclusa solo a saldo completo nel pagamento multi-area."
                        )

                    importo_riga = min(remaining, residuo).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    if importo_riga <= Decimal("0.00"):
                        continue

                    common_params = (
                        required(form_data, "data_pagamento", "Data pagamento"),
                        format(importo_riga, ".2f"),
                        required(form_data, "metodo_pagamento_id", "Metodo"),
                        optional(form_data, "riferimento"),
                        optional(form_data, "note"),
                        group_code,
                    )

                    if row["kind"] == "tesseramenti":
                        connection.execute(
                            """
                            INSERT INTO pagamenti_tesseramenti (
                                tesseramento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            (row["source_id"], *common_params),
                        )
                    elif row["kind"] == "corsi-rate":
                        connection.execute(
                            """
                            INSERT INTO pagamenti_rate_corsi (
                                rata_corso_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            (row["source_id"], *common_params),
                        )
                    elif row["kind"] == "campi-estivi":
                        connection.execute(
                            """
                            INSERT INTO pagamenti_campi_estivi (
                                iscrizione_campo_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            (row["source_id"], *common_params),
                        )
                    else:
                        connection.execute(
                            """
                            INSERT INTO pagamenti_eventi (
                                iscrizione_evento_id, data_pagamento, importo, metodo_pagamento_id, riferimento, note, gruppo_ricevuta
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            (row["source_id"], *common_params),
                        )

                    inserted += 1
                    remaining = (remaining - importo_riga).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                if inserted == 0:
                    raise ValueError("Nessun pagamento multi-area e stato registrato.")
                connection.commit()

            return redirect(
                start_response,
                f"/ricevute/multi-area-gruppo/{group_code}",
                ok="Pagamento multi-area registrato.",
                extra_query=context_query,
            )

        if path.startswith("/azioni/pagamenti-multi-area/elimina/"):
            group_code = path.removeprefix("/azioni/pagamenti-multi-area/elimina/").strip()
            if not group_code:
                raise ValueError("Pagamento multi-area non valido.")
            delete_multi_area_group(group_code)
            return redirect(
                start_response,
                "/maschere/pagamenti-multi-area",
                ok="Pagamento multi-area eliminato e scadenze ripristinate.",
                extra_query={**context_query, "vista": "dati"},
            )

        return redirect(start_response, "/", err="Azione non riconosciuta.", extra_query=context_query)

    except ValueError as error:
        return redirect(start_response, fallback_path(path), err=str(error), extra_query=context_query)
    except sqlite3.IntegrityError as error:
        return redirect(start_response, fallback_path(path), err=friendly_db_error(error), extra_query=context_query)
    except sqlite3.DatabaseError as error:
        return redirect(start_response, fallback_path(path), err=friendly_db_error(error), extra_query=context_query)


def fallback_path(path: str) -> str:
    if path.startswith("/azioni/crud/aggiorna/") or path.startswith("/azioni/crud/elimina/"):
        record_path = path.split("/")
        entity_key = record_path[4] if len(record_path) > 4 else ""
        config = CRUD_CONFIG.get(entity_key)
        if config:
            return config["return_path"]

    mapping = {
        "/azioni/associati/crea": "/maschere/associati",
        "/azioni/accesso/primo-admin": "/login",
        "/azioni/accesso/login": "/login",
        "/azioni/accesso/logout": "/login",
        "/azioni/accesso/cambia-password": "/maschere/accesso",
        "/azioni/accesso/recupera-password": "/recupera-password",
        "/azioni/utenti/crea": "/maschere/utenti",
        "/azioni/tesseramenti/crea": "/maschere/tesseramenti",
        "/azioni/tesseramenti/pagamento": "/maschere/tesseramenti",
        "/azioni/corsi/tipologia": "/maschere/corsi",
        "/azioni/corsi/crea": "/maschere/corsi",
        "/azioni/corsi/iscrizione": "/maschere/corsi",
        "/azioni/corsi/pagamento-iscrizione": "/maschere/corsi",
        "/azioni/corsi/rata": "/maschere/corsi",
        "/azioni/corsi/pagamento-rata": "/maschere/corsi",
        "/azioni/corsi/pagamento-rate": "/maschere/corsi",
        "/azioni/corsi/genera-rate": "/maschere/corsi",
        "/azioni/corsi/genera-rate-mancanti": "/",
        "/azioni/campi-estivi/crea": "/maschere/campi-estivi",
        "/azioni/campi-estivi/iscrizione": "/maschere/campi-estivi",
        "/azioni/campi-estivi/pagamento": "/maschere/campi-estivi",
        "/azioni/eventi/crea": "/maschere/eventi",
        "/azioni/eventi/iscrizione": "/maschere/eventi",
        "/azioni/eventi/pagamento": "/maschere/eventi",
        "/azioni/pagamenti-multi-area/crea": "/maschere/pagamenti-multi-area",
    }
    if path.startswith("/azioni/pagamenti-multi-area/elimina/"):
        return "/maschere/pagamenti-multi-area"
    if path.startswith("/azioni/utenti/aggiorna/") or path.startswith("/azioni/utenti/password/") or path.startswith("/azioni/utenti/stato/"):
        return f"/maschere/utenti/gestione/{path.rsplit('/', 1)[-1]}"
    return mapping.get(path, "/")


def response_message_from_headers(headers: list[tuple[str, str]]) -> tuple[str, str]:
    location = next((value for name, value in headers if name.lower() == "location"), "")
    if not location:
        return "", ""
    parsed = urlsplit(location)
    query = parse_qs(parsed.query, keep_blank_values=True)
    ok_message = " ".join(query.get("ok", [])).strip()
    err_message = " ".join(query.get("err", [])).strip()
    return ok_message, err_message


def request_client_ip(environ: dict) -> str:
    forwarded = (environ.get("HTTP_X_FORWARDED_FOR") or "").strip()
    if forwarded:
        return forwarded.split(",", 1)[0].strip()
    return (environ.get("REMOTE_ADDR") or "").strip()


def request_user_agent(environ: dict) -> str:
    return (environ.get("HTTP_USER_AGENT") or "").strip()


def request_work_year_for_log(query_params: dict[str, str], form_data: dict[str, str]) -> int | None:
    raw_value = (form_data.get("anno_lavoro") or query_params.get("anno_lavoro") or "").strip()
    if raw_value.isdigit():
        return int(raw_value)
    return None


def activity_username_for_log(
    current_user: dict[str, object] | None,
    path: str,
    form_data: dict[str, str],
) -> str:
    if current_user and current_user.get("username"):
        return str(current_user["username"])
    if path in {"/azioni/accesso/login", "/azioni/accesso/primo-admin", "/azioni/accesso/recupera-password"}:
        return normalized(form_data, "username", "") or "anonimo"
    return "anonimo"


def associato_log_context_by_id(associato_id: int) -> tuple[int | None, str, str]:
    row = fetch_one(
        f"""
        SELECT id, codice_associato, {associato_display_sql('')} AS associato
        FROM associati
        WHERE id = ?
        """,
        (associato_id,),
    )
    if row is None:
        return None, "", ""
    return int(row["id"]), str(row["codice_associato"] or ""), str(row["associato"] or "")


def resolve_activity_associato_context(
    path: str,
    query_params: dict[str, str],
    form_data: dict[str, str],
) -> tuple[int | None, str, str]:
    direct_value = normalized(form_data, "associato_id", "") or normalized(query_params, "associato_id", "")
    if direct_value.isdigit():
        return associato_log_context_by_id(int(direct_value))

    if path.startswith("/report/associato/"):
        associato_id = path.removeprefix("/report/associato/")
        if associato_id.isdigit():
            return associato_log_context_by_id(int(associato_id))

    if path.startswith("/export/excel/associato/"):
        associato_id = path.removeprefix("/export/excel/associato/")
        if associato_id.isdigit():
            return associato_log_context_by_id(int(associato_id))

    if path.startswith("/export/pdf/associato/"):
        associato_id = path.removeprefix("/export/pdf/associato/")
        if associato_id.isdigit():
            return associato_log_context_by_id(int(associato_id))

    if path.startswith("/modifica/associati/"):
        associato_id = path.removeprefix("/modifica/associati/")
        if associato_id.isdigit():
            return associato_log_context_by_id(int(associato_id))

    if path.startswith("/azioni/crud/aggiorna/associati/") or path.startswith("/azioni/crud/elimina/associati/"):
        associato_id = path.rsplit("/", 1)[-1]
        if associato_id.isdigit():
            return associato_log_context_by_id(int(associato_id))

    tesseramento_id = normalized(form_data, "tesseramento_id", "")
    if tesseramento_id.isdigit():
        row = fetch_one("SELECT associato_id FROM tesseramenti_annuali WHERE id = ?", (int(tesseramento_id),))
        if row is not None:
            return associato_log_context_by_id(int(row["associato_id"]))

    iscrizione_campo_id = normalized(form_data, "iscrizione_campo_id", "")
    if iscrizione_campo_id.isdigit():
        row = fetch_one("SELECT associato_id FROM iscrizioni_campi_estivi WHERE id = ?", (int(iscrizione_campo_id),))
        if row is not None:
            return associato_log_context_by_id(int(row["associato_id"]))

    iscrizione_evento_id = normalized(form_data, "iscrizione_evento_id", "")
    if iscrizione_evento_id.isdigit():
        row = fetch_one("SELECT associato_id FROM iscrizioni_eventi WHERE id = ?", (int(iscrizione_evento_id),))
        if row is not None:
            return associato_log_context_by_id(int(row["associato_id"]))

    rata_ids = multi_values(form_data, "rata_corso_id")
    if rata_ids:
        placeholders = ",".join("?" for _ in rata_ids)
        rows = fetch_all(
            f"""
            SELECT DISTINCT ic.associato_id
            FROM rate_corsi_mensili r
            JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
            WHERE r.id IN ({placeholders})
            """,
            tuple(int(rata_id) for rata_id in rata_ids if rata_id.isdigit()),
        )
        associato_ids = {int(row["associato_id"]) for row in rows if row["associato_id"] is not None}
        if len(associato_ids) == 1:
            return associato_log_context_by_id(next(iter(associato_ids)))

    if path.startswith("/azioni/pagamenti-multi-area/elimina/"):
        group_code = path.removeprefix("/azioni/pagamenti-multi-area/elimina/").strip()
        if group_code:
            row = fetch_one(
                """
                SELECT associato_id
                FROM (
                    SELECT t.associato_id AS associato_id
                    FROM pagamenti_tesseramenti pt
                    JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
                    WHERE pt.gruppo_ricevuta = ?
                    UNION ALL
                    SELECT ic.associato_id AS associato_id
                    FROM pagamenti_rate_corsi prc
                    JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
                    JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
                    WHERE prc.gruppo_ricevuta = ?
                    UNION ALL
                    SELECT ice.associato_id AS associato_id
                    FROM pagamenti_campi_estivi pce
                    JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
                    WHERE pce.gruppo_ricevuta = ?
                    UNION ALL
                    SELECT ie.associato_id AS associato_id
                    FROM pagamenti_eventi pe
                    JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
                    WHERE pe.gruppo_ricevuta = ?
                ) groups
                LIMIT 1
                """,
                (group_code, group_code, group_code, group_code),
            )
            if row is not None and row["associato_id"] is not None:
                return associato_log_context_by_id(int(row["associato_id"]))

    return None, "", ""


def report_title_for_key(report_key: str, query_params: dict[str, str]) -> str:
    try:
        return get_report_definition(report_key, query_params)["title"]
    except Exception:
        return report_key or "Report"


def describe_logged_activity(
    method: str,
    path: str,
    query_params: dict[str, str],
    form_data: dict[str, str],
) -> tuple[str, str, str]:
    work_year = request_work_year_for_log(query_params, form_data)
    work_year_detail = f"Anno di lavoro: {work_year}" if work_year else ""

    if method == "POST":
        if path == "/azioni/accesso/primo-admin":
            return "Accesso", "Configurazione utente amministratore iniziale", ""
        if path == "/azioni/accesso/login":
            return "Accesso", "Accesso al software", ""
        if path == "/azioni/accesso/logout":
            return "Accesso", "Logout dal software", ""
        if path == "/azioni/accesso/cambia-password":
            return "Accesso", "Cambio password del profilo corrente", ""
        if path == "/azioni/accesso/recupera-password":
            return "Accesso", "Recupero password utente standard", ""
        if path == "/azioni/utenti/crea":
            return "Utenti", "Creazione nuovo utente standard", ""
        if path.startswith("/azioni/utenti/aggiorna/"):
            return "Utenti", "Aggiornamento dati utente standard", ""
        if path.startswith("/azioni/utenti/password/"):
            return "Utenti", "Reimpostazione password utente standard", ""
        if path.startswith("/azioni/utenti/stato/"):
            action = normalized(form_data, "azione_stato", "")
            detail = "Azione: disattivazione" if action == "disattiva" else "Azione: riattivazione"
            return "Utenti", "Aggiornamento stato utente standard", detail
        if path == "/azioni/associati/crea":
            return "Associati", "Creazione nuovo associato", plain_text(
                f"{normalized(form_data, 'cognome', '')} {normalized(form_data, 'nome', '')}"
            ).strip()
        if path.startswith("/azioni/crud/aggiorna/associati/"):
            return "Associati", "Aggiornamento anagrafica associato", ""
        if path.startswith("/azioni/crud/elimina/associati/"):
            return "Associati", "Eliminazione associato", ""
        if path == "/azioni/tesseramenti/crea":
            detail = work_year_detail or f"Anno sociale: {normalized(form_data, 'anno_sociale', '')}"
            return "Tesseramenti", "Registrazione nuovo tesseramento", detail
        if path == "/azioni/tesseramenti/pagamento":
            return "Tesseramenti", "Registrazione pagamento tesseramento", ""
        if path == "/azioni/corsi/crea":
            return "Corsi", "Creazione nuovo corso", normalized(form_data, "nome", "")
        if path == "/azioni/corsi/iscrizione":
            return "Corsi", "Registrazione iscrizione corso", work_year_detail
        if path in {"/azioni/corsi/pagamento-rata", "/azioni/corsi/pagamento-rate"}:
            return "Corsi", "Registrazione pagamento quote mensili corso", work_year_detail
        if path == "/azioni/corsi/genera-rate":
            month_value = normalized(form_data, "mese", "")
            month_text = month_label(int(month_value)) if month_value.isdigit() else month_value
            return "Corsi", "Generazione massiva quote mensili corsi", (
                f"Competenza: {month_text} {normalized(form_data, 'anno', '')}".strip()
            )
        if path == "/azioni/corsi/genera-rate-mancanti":
            return "Corsi", "Generazione automatica quote mensili mancanti", work_year_detail
        if path == "/azioni/campi-estivi/crea":
            return ESTATE_LABEL, "Creazione scheda annuale Campo estivo", work_year_detail
        if path == "/azioni/campi-estivi/iscrizione":
            return ESTATE_LABEL, "Registrazione iscrizione Campo estivo", work_year_detail
        if path == "/azioni/campi-estivi/pagamento":
            return ESTATE_LABEL, "Registrazione pagamento Campo estivo", work_year_detail
        if path == "/azioni/eventi/crea":
            return "Eventi", "Creazione nuovo evento", normalized(form_data, "nome", "")
        if path == "/azioni/eventi/iscrizione":
            return "Eventi", "Registrazione iscrizione evento", work_year_detail
        if path == "/azioni/eventi/pagamento":
            return "Eventi", "Registrazione pagamento evento", work_year_detail
        if path == "/azioni/pagamenti-multi-area/crea":
            return "Pagamenti", "Registrazione pagamento multi-area", work_year_detail
        if path.startswith("/azioni/pagamenti-multi-area/elimina/"):
            return "Pagamenti", "Annullamento pagamento multi-area", work_year_detail
        if path.startswith("/azioni/crud/aggiorna/"):
            entity_key = path.removeprefix("/azioni/crud/aggiorna/").split("/", 1)[0]
            return "Dati", f"Aggiornamento record {entity_key}", work_year_detail
        if path.startswith("/azioni/crud/elimina/"):
            entity_key = path.removeprefix("/azioni/crud/elimina/").split("/", 1)[0]
            return "Dati", f"Eliminazione record {entity_key}", work_year_detail
        return "Operazioni", f"Azione {path}", work_year_detail

    if path == "/api/report-share":
        channel = normalized(query_params, "channel", "")
        report_title = report_title_for_key(normalized(query_params, "report_key", ""), query_params)
        detail = f"Canale: {channel}" if channel else ""
        if channel == "whatsapp-group":
            detail = "Canale: gruppo WhatsApp"
        return "Report", f"Preparazione invio report {report_title}", detail

    if path.startswith("/export/excel/"):
        if path.startswith("/export/excel/associato/"):
            return "Report", "Export Excel dettaglio associato", work_year_detail
        report_key = path.removeprefix("/export/excel/")
        return "Report", f"Export Excel report {report_title_for_key(report_key, query_params)}", work_year_detail

    if path.startswith("/export/pdf/ricevuta/"):
        return "Ricevute", "Export PDF ricevuta di pagamento", ""

    if path.startswith("/export/pdf/"):
        if path.startswith("/export/pdf/associato/"):
            return "Report", "Export PDF dettaglio associato", work_year_detail
        report_key = path.removeprefix("/export/pdf/")
        return "Report", f"Export PDF report {report_title_for_key(report_key, query_params)}", work_year_detail

    if path == "/":
        return "Navigazione", "Apertura dashboard", work_year_detail

    page_titles = {
        "/login": "Apertura pagina di accesso",
        "/recupera-password": "Apertura pagina recupero password",
        "/maschere/accesso": "Apertura profilo accesso",
        "/maschere/utenti": "Apertura maschera utenti",
        "/maschere/consiglio-direttivo": "Apertura maschera Consiglio Direttivo",
        "/maschere/associati": "Apertura maschera associati",
        "/maschere/tesseramenti": "Apertura maschera tesseramenti",
        "/maschere/corsi": "Apertura maschera corsi",
        "/maschere/campi-estivi": "Apertura maschera Campo estivo",
        "/maschere/eventi": "Apertura maschera eventi",
        "/maschere/pagamenti-multi-area": "Apertura maschera pagamenti",
        "/report/associati": "Apertura report Posizione associati",
        "/report/tesseramenti": "Apertura report Situazione tesseramenti",
        "/report/scadenze": "Apertura report Scadenze da incassare",
        "/report/corsi": "Apertura report Situazione corsi",
        "/report/campi-estivi": "Apertura report Situazione campo estivo",
        "/report/eventi": "Apertura report Situazione eventi",
        "/report/partecipanti": "Apertura report Partecipanti attività",
        "/report/incassi": "Apertura report Incassi totali",
        "/report/registro-attivita": "Apertura report Registro attivita",
    }
    if path in page_titles:
        category = "Report" if path.startswith("/report/") else "Navigazione"
        return category, page_titles[path], work_year_detail
    if path.startswith("/report/associato/"):
        return "Report", "Apertura dettaglio associato", work_year_detail
    if path.startswith("/ricevute/"):
        return "Ricevute", "Apertura ricevuta di pagamento", work_year_detail
    if path.startswith("/modifica/"):
        return "Navigazione", "Apertura maschera modifica record", work_year_detail
    return "Navigazione", f"Accesso a {path}", work_year_detail


def should_log_activity(method: str, path: str) -> bool:
    if path.startswith("/static/"):
        return False
    if path in {"/api/codice-fiscale", "/api/codice-fiscale/calcola", "/api/comuni", "/api/cap"}:
        return False
    if path == "/favicon.ico":
        return False
    return (
        path.startswith("/azioni/")
        or path.startswith("/report/")
        or path.startswith("/export/")
        or path.startswith("/ricevute/")
        or path.startswith("/modifica/")
        or path.startswith("/maschere/")
        or path in {"/", "/login", "/recupera-password", "/api/report-share"}
    )


def log_request_activity(
    environ: dict,
    method: str,
    path: str,
    query_params: dict[str, str],
    form_data: dict[str, str],
    current_user: dict[str, object] | None,
    response_status: str,
    response_headers: list[tuple[str, str]],
) -> None:
    if not should_log_activity(method, path):
        return
    category, description, detail = describe_logged_activity(method, path, query_params, form_data)
    associato_id, associato_codice, associato_nominativo = resolve_activity_associato_context(path, query_params, form_data)
    ok_message, err_message = response_message_from_headers(response_headers)
    if err_message:
        outcome = f"Errore: {err_message}"
    elif ok_message:
        outcome = f"OK: {ok_message}"
    elif response_status.startswith("2"):
        outcome = "Completata"
    else:
        outcome = response_status
    record_activity(
        username=activity_username_for_log(current_user, path, form_data),
        path=path,
        method=method,
        description=description,
        category=category,
        outcome=outcome,
        work_year=request_work_year_for_log(query_params, form_data),
        detail=detail,
        associato_id=associato_id,
        associato_codice=associato_codice,
        associato_nominativo=associato_nominativo,
        ip_client=request_client_ip(environ),
        user_agent=request_user_agent(environ),
    )


def serve_static(start_response, path: str):
    relative_path = path.removeprefix("/static/")
    target = (STATIC_DIR / relative_path).resolve()

    if not target.is_file() or STATIC_DIR.resolve() not in target.parents:
        return not_found(start_response)

    mime_type, _ = mimetypes.guess_type(str(target))
    if mime_type == "text/css":
        content_type = "text/css; charset=utf-8"
    else:
        content_type = mime_type or "application/octet-stream"

    start_response("200 OK", [("Content-Type", content_type)])
    return [target.read_bytes()]


def not_found(start_response) -> list[bytes]:
    start_response("404 Not Found", [("Content-Type", "text/plain; charset=utf-8")])
    return [b"Pagina non trovata."]


def dispatch_request(
    environ,
    start_response,
    *,
    parsed_request: tuple[str, str, dict[str, str], dict[str, str], dict[str, str]] | None = None,
    current_user: dict[str, object] | None = None,
):
    if parsed_request is None:
        path, method, query_params, form_data, request_cookies = parse_request(environ)
    else:
        path, method, query_params, form_data, request_cookies = parsed_request
    if current_user is None:
        current_user = current_user_from_cookies(request_cookies)

    if path.startswith("/static/"):
        return serve_static(start_response, path)

    if method == "GET" and path == "/login":
        if current_user is not None:
            return redirect(start_response, "/", extra_query=work_year_query(query_params))
        start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
        return [login_page(query_params)]

    if method == "GET" and path == "/recupera-password":
        if current_user is not None:
            return redirect(start_response, "/", extra_query=work_year_query(query_params))
        start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
        return [recover_password_page(query_params)]

    if method == "POST" and path in {
        "/azioni/accesso/primo-admin",
        "/azioni/accesso/login",
        "/azioni/accesso/logout",
        "/azioni/accesso/recupera-password",
    }:
        return handle_post(path, form_data, start_response, current_user, request_cookies)

    if current_user is None:
        if path.startswith("/api/"):
            return json_response(
                start_response,
                {"ok": False, "error": "Sessione non valida. Accedi di nuovo."},
                status="401 Unauthorized",
            )
        request_query = {key: value for key, value in query_params.items() if key not in {"ok", "err", "next"}}
        requested_target = login_destination(urlunsplit(("", "", path, urlencode(request_query), "")))
        login_query = {"next": requested_target}
        message = "Primo accesso: crea l'utente amministratore." if bootstrap_admin_required() else None
        return redirect(start_response, "/login", ok=message, extra_query=login_query)

    if method == "GET" and path == "/api/codice-fiscale":
        payload = decode_codice_fiscale_birth_data(query_params.get("value", ""))
        return json_response(start_response, {"found": bool(payload), **(payload or {})})

    if method == "GET" and path == "/api/codice-fiscale/calcola":
        payload = calculate_codice_fiscale(
            nome=query_params.get("nome", ""),
            cognome=query_params.get("cognome", ""),
            data_nascita=query_params.get("data_nascita", ""),
            sesso=query_params.get("sesso", ""),
            comune_nascita=query_params.get("comune_nascita", ""),
        )
        return json_response(start_response, {"found": bool(payload), **(payload or {})})

    if method == "GET" and path == "/api/comuni":
        payload = lookup_comune_details(query_params.get("nome", ""))
        return json_response(start_response, {"found": bool(payload), **(payload or {})})

    if method == "GET" and path == "/api/cap":
        payload = lookup_cap_details(query_params.get("value", ""))
        return json_response(start_response, {"found": bool(payload), **(payload or {})})

    if method == "GET" and path == "/api/report-share":
        try:
            report_key = query_params.get("report_key", "")
            if report_requires_admin(report_key) and not current_user.get("is_admin"):
                return json_response(
                    start_response,
                    {"ok": False, "error": "Solo l'amministratore puo usare questo report."},
                    status="403 Forbidden",
                )
            recipient_id = query_params.get("recipient_id", "")
            channel = query_params.get("channel", "")
            share_query = {
                key: value
                for key, value in query_params.items()
                if key not in {"report_key", "recipient_id", "channel"}
            }
            payload = build_report_share_payload(report_key, share_query, recipient_id, channel)
            if payload is None:
                return json_response(start_response, {"ok": False, "error": "Report non disponibile."}, status="400 Bad Request")
            return json_response(start_response, {"ok": True, **payload})
        except KeyError:
            return json_response(start_response, {"ok": False, "error": "Report non disponibile."}, status="404 Not Found")
        except ValueError as error:
            return json_response(start_response, {"ok": False, "error": str(error)}, status="400 Bad Request")

    if method == "GET" and path == "/api/pagamenti-multi-area/quote-corso-future":
        try:
            work_year = current_work_year(query_params)
            associato_id = int(required(query_params, "associato_id", "Associato"))
            iscrizione_corso_id = int(required(query_params, "iscrizione_corso_id", "Corso"))
            fine_competenza = required(query_params, "fine_competenza", "Competenza finale")
            payload = generate_multi_area_future_course_payload(
                work_year,
                associato_id,
                iscrizione_corso_id,
                fine_competenza,
            )
            return json_response(start_response, payload)
        except ValueError as error:
            return json_response(start_response, {"ok": False, "error": str(error)}, status="400 Bad Request")

    if method == "GET" and path.startswith("/export/pdf/ricevuta/"):
        receipt_path = path.removeprefix("/export/pdf/ricevuta/")
        if "/" not in receipt_path:
            return not_found(start_response)
        payment_type, payment_id = receipt_path.split("/", 1)
        return export_receipt_pdf(start_response, payment_type, payment_id)

    if method == "GET" and path.startswith("/export/excel/"):
        if path.startswith("/export/excel/associato/"):
            associato_id = path.removeprefix("/export/excel/associato/")
            if not associato_id.isdigit():
                return not_found(start_response)
            return export_associato_detail_excel(start_response, int(associato_id), query_params)
        report_key = path.removeprefix("/export/excel/")
        if report_requires_admin(report_key) and not current_user.get("is_admin"):
            return redirect(
                start_response,
                "/",
                err="Solo l'amministratore puo esportare questo report.",
                extra_query=work_year_query(query_params),
            )
        return export_report_excel(start_response, report_key, query_params)

    if method == "GET" and path.startswith("/export/pdf/"):
        if path.startswith("/export/pdf/associato/"):
            associato_id = path.removeprefix("/export/pdf/associato/")
            if not associato_id.isdigit():
                return not_found(start_response)
            return export_associato_detail_pdf(start_response, int(associato_id), query_params)
        report_key = path.removeprefix("/export/pdf/")
        if report_requires_admin(report_key) and not current_user.get("is_admin"):
            return redirect(
                start_response,
                "/",
                err="Solo l'amministratore puo esportare questo report.",
                extra_query=work_year_query(query_params),
            )
        return export_report_pdf(start_response, report_key, query_params)

    if method == "POST":
        return handle_post(path, form_data, start_response, current_user, request_cookies)

    if method == "GET" and path.startswith("/modifica/"):
        edit_path_value = path.removeprefix("/modifica/")
        if "/" not in edit_path_value:
            return not_found(start_response)
        entity_key, record_id = edit_path_value.split("/", 1)
        try:
            start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
            return [render_crud_edit_page(entity_key, int(record_id), query_params, current_user)]
        except (KeyError, ValueError):
            return not_found(start_response)

    if method == "GET" and path.startswith("/ricevute/"):
        receipt_path = path.removeprefix("/ricevute/")
        if "/" not in receipt_path:
            return not_found(start_response)
        payment_type, payment_id = receipt_path.split("/", 1)
        try:
            start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
            if payment_type in {"corsi-rate-gruppo", "multi-area-gruppo"}:
                return [receipt_page(payment_type, payment_id, query_params, current_user)]
            return [receipt_page(payment_type, int(payment_id), query_params, current_user)]
        except (KeyError, ValueError):
            return not_found(start_response)

    if method == "GET" and path.startswith("/report/associato/"):
        associato_id = path.removeprefix("/report/associato/")
        try:
            start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
            return [associato_report_page(int(associato_id), query_params, current_user)]
        except (KeyError, ValueError):
            return not_found(start_response)

    if method == "GET" and path.startswith("/maschere/utenti/gestione/"):
        if not current_user.get("is_admin"):
            return redirect(
                start_response,
                "/",
                err="Solo l'amministratore puo accedere alla gestione utenti.",
                extra_query=work_year_query(query_params),
            )
        user_id = path.removeprefix("/maschere/utenti/gestione/")
        try:
            start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
            return [gestione_utente_page(int(user_id), query_params, current_user)]
        except (KeyError, ValueError):
            return not_found(start_response)

    if path == "/maschere/utenti" and not current_user.get("is_admin"):
        return redirect(
            start_response,
            "/",
            err="Solo l'amministratore puo accedere alla gestione utenti.",
            extra_query=work_year_query(query_params),
        )

    if path == "/report/registro-attivita" and not current_user.get("is_admin"):
        return redirect(
            start_response,
            "/",
            err="Solo l'amministratore puo accedere al registro attivita.",
            extra_query=work_year_query(query_params),
        )

    routes = {
        "/": dashboard_page,
        "/maschere/accesso": access_profile_page,
        "/maschere/utenti": utenti_page,
        "/maschere/consiglio-direttivo": consiglio_direttivo_page,
        "/maschere/associati": associati_page,
        "/maschere/tesseramenti": tesseramenti_page,
        "/maschere/corsi": corsi_page,
        "/maschere/campi-estivi": campi_estivi_page,
        "/maschere/eventi": eventi_page,
        "/maschere/pagamenti-multi-area": pagamenti_multi_area_page,
        "/report/associati": report_associati,
        "/report/tesseramenti": report_tesseramenti,
        "/report/scadenze": report_scadenze,
        "/report/corsi": report_corsi,
        "/report/campi-estivi": report_campi_estivi,
        "/report/eventi": report_eventi,
        "/report/partecipanti": report_partecipanti,
        "/report/incassi": report_incassi,
        "/report/registro-attivita": report_registro_attivita,
    }

    handler = routes.get(path)
    if handler is None:
        return not_found(start_response)

    start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
    return [handler(query_params, current_user=current_user)]


def app(environ, start_response):
    parsed_request = parse_request(environ)
    path, method, query_params, form_data, request_cookies = parsed_request
    current_user = current_user_from_cookies(request_cookies)
    captured_response: dict[str, object] = {
        "status": "",
        "headers": [],
    }

    def logging_start_response(status, headers, exc_info=None):
        captured_response["status"] = status
        captured_response["headers"] = list(headers)
        if exc_info is None:
            return start_response(status, headers)
        return start_response(status, headers, exc_info)

    response = dispatch_request(
        environ,
        logging_start_response,
        parsed_request=parsed_request,
        current_user=current_user,
    )
    try:
        log_request_activity(
            environ,
            method,
            path,
            query_params,
            form_data,
            current_user,
            str(captured_response.get("status", "")),
            list(captured_response.get("headers", [])),
        )
    except Exception:
        pass
    return response


ensure_schema()


class ThreadingWSGIServer(ThreadingMixIn, WSGIServer):
    daemon_threads = True


def main() -> None:
    try:
        auto_generate_course_rates_on_startup()
    except Exception:
        pass
    host = os.environ.get("ASSOCIAZIONE_HOST", "127.0.0.1")
    port = int(os.environ.get("ASSOCIAZIONE_PORT", "8000"))

    print(f"Database: {DB_PATH}")
    print(f"Web app disponibile su http://{host}:{port}")
    with make_server(host, port, app, server_class=ThreadingWSGIServer) as server:
        server.serve_forever()


if __name__ == "__main__":
    main()
