from __future__ import annotations

import json
import os
import shutil
import sqlite3
import base64
import gzip
from contextlib import contextmanager
from datetime import datetime
from io import StringIO
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


TABLE_NAMES = (
    "reports_clean",
    "products_clean",
    "cases_mvp",
    "cases_enriched",
    "ai_cases_input",
    "invalid_reports",
    "summary",
    "decision_summary",
)


def _supabase_settings() -> tuple[str, str] | None:
    url = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
    key = (
        os.environ.get("SUPABASE_SECRET_KEY", "").strip()
        or os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    )
    if not url and not key:
        return None
    if not url or not key:
        raise RuntimeError("SUPABASE_URL a SUPABASE_SECRET_KEY musia byť nastavené spolu.")
    return url, key


def uses_persistent_cloud_storage() -> bool:
    return _supabase_settings() is not None


def storage_location_label() -> str:
    if uses_persistent_cloud_storage():
        return "Supabase – trvalé cloudové úložisko"
    return f"Lokálna SQLite – {database_path().name} (na Streamlit Cloud je dočasná)"


def database_path() -> Path:
    configured_path = os.environ.get("INSTRUCTION_VALIDATOR_DB")
    if configured_path:
        return Path(configured_path)
    return Path(__file__).resolve().parent / ".data" / "instruction_validator.sqlite3"


def backup_directory() -> Path:
    configured_path = os.environ.get("INSTRUCTION_VALIDATOR_BACKUP_DIR")
    if configured_path:
        return Path(configured_path)
    return database_path().parent / "backups"


@contextmanager
def _connect():
    path = database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _supabase_request(
    method: str,
    table: str,
    *,
    query: dict[str, str] | None = None,
    payload=None,
    prefer: str | None = None,
):
    settings = _supabase_settings()
    if settings is None:
        raise RuntimeError("Supabase nie je nakonfigurovaný.")
    base_url, key = settings
    target = f"{base_url}/rest/v1/{table}"
    if query:
        target = f"{target}?{urlencode(query)}"
    headers = {
        "apikey": key,
        "Content-Type": "application/json",
    }
    if not key.startswith("sb_secret_"):
        headers["Authorization"] = f"Bearer {key}"
    if prefer:
        headers["Prefer"] = prefer
    body = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(target, data=body, headers=headers, method=method)
    try:
        with urlopen(request, timeout=30) as response:
            response_body = response.read()
    except HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Supabase požiadavka zlyhala ({error.code}): {detail}") from error
    if not response_body:
        return None
    return json.loads(response_body.decode("utf-8"))


def initialize_storage() -> None:
    if uses_persistent_cloud_storage():
        _supabase_request("GET", "batches", query={"select": "batch_id", "limit": "1"})
        return
    with _connect() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS batches (
                batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                input_signature TEXT NOT NULL,
                config_json TEXT NOT NULL,
                warnings_json TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS batch_tables (
                batch_id INTEGER NOT NULL,
                table_name TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                PRIMARY KEY (batch_id, table_name),
                FOREIGN KEY (batch_id) REFERENCES batches(batch_id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS reviews (
                batch_id INTEGER NOT NULL,
                case_id TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (batch_id, case_id),
                FOREIGN KEY (batch_id) REFERENCES batches(batch_id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_batches_created_at ON batches(created_at);
            """
        )


def _frame_to_json(frame: pd.DataFrame) -> str:
    raw = frame.to_json(orient="split", date_format="iso", default_handler=str).encode("utf-8")
    return "gzip:" + base64.b64encode(gzip.compress(raw)).decode("ascii")


def _frame_from_json(payload: str) -> pd.DataFrame:
    if payload.startswith("gzip:"):
        payload = gzip.decompress(base64.b64decode(payload[5:])).decode("utf-8")
    return pd.read_json(StringIO(payload), orient="split", convert_dates=False)


def save_pipeline_result(result, input_signature: str, config: dict) -> int:
    initialize_storage()
    created_at = datetime.now().astimezone().replace(microsecond=0).isoformat()
    if uses_persistent_cloud_storage():
        inserted = _supabase_request(
            "POST",
            "batches",
            query={"select": "batch_id"},
            payload={
                "created_at": created_at,
                "input_signature": input_signature,
                "config_json": json.dumps(config, ensure_ascii=False),
                "warnings_json": json.dumps(result.warnings, ensure_ascii=False),
            },
            prefer="return=representation",
        )
        batch_id = int(inserted[0]["batch_id"])
        for table_name in TABLE_NAMES:
            _supabase_request(
                "POST",
                "batch_tables",
                payload={
                    "batch_id": batch_id,
                    "table_name": table_name,
                    "payload_json": _frame_to_json(getattr(result, table_name)),
                },
                prefer="return=minimal",
            )
        return batch_id
    with _connect() as connection:
        cursor = connection.execute(
            "INSERT INTO batches (created_at, input_signature, config_json, warnings_json) VALUES (?, ?, ?, ?)",
            (created_at, input_signature, json.dumps(config, ensure_ascii=False), json.dumps(result.warnings, ensure_ascii=False)),
        )
        batch_id = int(cursor.lastrowid)
        for table_name in TABLE_NAMES:
            connection.execute(
                "INSERT INTO batch_tables (batch_id, table_name, payload_json) VALUES (?, ?, ?)",
                (batch_id, table_name, _frame_to_json(getattr(result, table_name))),
            )
    return batch_id


def _result_from_rows(batch_row, table_rows):
    from processing import PipelineResult

    tables = {row["table_name"]: _frame_from_json(row["payload_json"]) for row in table_rows}
    if set(TABLE_NAMES) - set(tables):
        return None
    return PipelineResult(
        **{table_name: tables[table_name] for table_name in TABLE_NAMES},
        warnings=json.loads(batch_row["warnings_json"]),
    )


def _load_batch(connection: sqlite3.Connection, batch_row: sqlite3.Row):
    table_rows = connection.execute(
        "SELECT table_name, payload_json FROM batch_tables WHERE batch_id = ?",
        (batch_row["batch_id"],),
    ).fetchall()
    return _result_from_rows(batch_row, table_rows)


def load_latest_pipeline_result():
    initialize_storage()
    if uses_persistent_cloud_storage():
        batches = _supabase_request(
            "GET",
            "batches",
            query={"select": "batch_id,warnings_json", "order": "batch_id.desc", "limit": "1"},
        )
        if not batches:
            return None, None
        batch_row = batches[0]
        table_rows = _supabase_request(
            "GET",
            "batch_tables",
            query={"select": "table_name,payload_json", "batch_id": f"eq.{batch_row['batch_id']}"},
        )
        result = _result_from_rows(batch_row, table_rows)
        return (result, int(batch_row["batch_id"])) if result is not None else (None, None)
    with _connect() as connection:
        batch_row = connection.execute("SELECT * FROM batches ORDER BY batch_id DESC LIMIT 1").fetchone()
        if batch_row is None:
            return None, None
        result = _load_batch(connection, batch_row)
        return (result, int(batch_row["batch_id"])) if result is not None else (None, None)


def latest_batch_key() -> tuple[int, str] | None:
    """Return a small cache key that changes only when a new batch is uploaded."""
    initialize_storage()
    if uses_persistent_cloud_storage():
        rows = _supabase_request(
            "GET",
            "batches",
            query={"select": "batch_id,created_at", "order": "batch_id.desc", "limit": "1"},
        )
        if not rows:
            return None
        return int(rows[0]["batch_id"]), str(rows[0]["created_at"])
    with _connect() as connection:
        row = connection.execute(
            "SELECT batch_id, created_at FROM batches ORDER BY batch_id DESC LIMIT 1"
        ).fetchone()
    if row is None:
        return None
    return int(row["batch_id"]), str(row["created_at"])


def load_reviews(batch_id: int | None) -> dict[str, dict]:
    if batch_id is None:
        return {}
    initialize_storage()
    if uses_persistent_cloud_storage():
        rows = _supabase_request(
            "GET",
            "reviews",
            query={"select": "case_id,payload_json", "batch_id": f"eq.{batch_id}"},
        )
    else:
        with _connect() as connection:
            rows = connection.execute(
                "SELECT case_id, payload_json FROM reviews WHERE batch_id = ?",
                (batch_id,),
            ).fetchall()
    return {str(row["case_id"]): json.loads(row["payload_json"]) for row in rows}


def save_review(batch_id: int | None, case_id: str, review: dict) -> None:
    if batch_id is None:
        return
    initialize_storage()
    updated_at = datetime.now().astimezone().replace(microsecond=0).isoformat()
    values = {
        "batch_id": batch_id,
        "case_id": str(case_id),
        "payload_json": json.dumps(review, ensure_ascii=False),
        "updated_at": updated_at,
    }
    if uses_persistent_cloud_storage():
        _supabase_request(
            "POST",
            "reviews",
            query={"on_conflict": "batch_id,case_id"},
            payload=values,
            prefer="resolution=merge-duplicates,return=minimal",
        )
        return
    with _connect() as connection:
        connection.execute(
            """
            INSERT INTO reviews (batch_id, case_id, payload_json, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(batch_id, case_id) DO UPDATE SET
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (batch_id, str(case_id), values["payload_json"], updated_at),
        )


def backup_database(retention: int = 20) -> Path | None:
    """Create a downloadable backup of the active storage."""
    target_dir = backup_directory()
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    if uses_persistent_cloud_storage():
        result, batch_id = load_latest_pipeline_result()
        if result is None or batch_id is None:
            return None
        target_path = target_dir / f"instruction-validator-{timestamp}.json"
        payload = {
            "exported_at": datetime.now().astimezone().replace(microsecond=0).isoformat(),
            "batch_id": batch_id,
            "warnings": result.warnings,
            "tables": {name: _frame_to_json(getattr(result, name)) for name in TABLE_NAMES},
            "reviews": load_reviews(batch_id),
        }
        target_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        pattern = "instruction-validator-*.json"
    else:
        source_path = database_path()
        if not source_path.exists():
            return None
        target_path = target_dir / f"instruction-validator-{timestamp}.sqlite3"
        shutil.copy2(source_path, target_path)
        pattern = "instruction-validator-*.sqlite3"
    backups = sorted(target_dir.glob(pattern), key=lambda path: path.stat().st_mtime, reverse=True)
    for old_backup in backups[max(retention, 1):]:
        old_backup.unlink()
    return target_path
