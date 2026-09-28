from __future__ import annotations

import json
import os
import shutil
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from io import StringIO
from pathlib import Path

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


def initialize_storage() -> None:
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
    return frame.to_json(orient="split", date_format="iso", default_handler=str)


def _frame_from_json(payload: str) -> pd.DataFrame:
    return pd.read_json(StringIO(payload), orient="split", convert_dates=False)


def save_pipeline_result(result, input_signature: str, config: dict) -> int:
    initialize_storage()
    created_at = datetime.now().astimezone().replace(microsecond=0).isoformat()
    with _connect() as connection:
        cursor = connection.execute(
            "INSERT INTO batches (created_at, input_signature, config_json, warnings_json) VALUES (?, ?, ?, ?)",
            (created_at, input_signature, json.dumps(config, ensure_ascii=False), json.dumps(result.warnings, ensure_ascii=False)),
        )
        batch_id = int(cursor.lastrowid)
        for table_name in TABLE_NAMES:
            frame = getattr(result, table_name)
            connection.execute(
                "INSERT INTO batch_tables (batch_id, table_name, payload_json) VALUES (?, ?, ?)",
                (batch_id, table_name, _frame_to_json(frame)),
            )
    return batch_id


def _load_batch(connection: sqlite3.Connection, batch_row: sqlite3.Row):
    from processing import PipelineResult

    table_rows = connection.execute(
        "SELECT table_name, payload_json FROM batch_tables WHERE batch_id = ?",
        (batch_row["batch_id"],),
    ).fetchall()
    tables = {row["table_name"]: _frame_from_json(row["payload_json"]) for row in table_rows}
    if set(TABLE_NAMES) - set(tables):
        return None
    return PipelineResult(
        **{table_name: tables[table_name] for table_name in TABLE_NAMES},
        warnings=json.loads(batch_row["warnings_json"]),
    )


def load_latest_pipeline_result():
    initialize_storage()
    with _connect() as connection:
        batch_row = connection.execute(
            "SELECT * FROM batches ORDER BY batch_id DESC LIMIT 1"
        ).fetchone()
        if batch_row is None:
            return None, None
        result = _load_batch(connection, batch_row)
        if result is None:
            return None, None
        return result, int(batch_row["batch_id"])


def latest_batch_key() -> tuple[int, str] | None:
    """Return a small cache key that changes only when a new batch is uploaded."""
    initialize_storage()
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
    with _connect() as connection:
        connection.execute(
            """
            INSERT INTO reviews (batch_id, case_id, payload_json, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(batch_id, case_id) DO UPDATE SET
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (batch_id, str(case_id), json.dumps(review, ensure_ascii=False), updated_at),
        )


def backup_database(retention: int = 20) -> Path | None:
    """Create a consistent SQLite copy and retain only the newest backups."""
    source_path = database_path()
    if not source_path.exists():
        return None
    target_dir = backup_directory()
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    target_path = target_dir / f"instruction-validator-{timestamp}.sqlite3"
    shutil.copy2(source_path, target_path)
    backups = sorted(
        target_dir.glob("instruction-validator-*.sqlite3"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old_backup in backups[max(retention, 1):]:
        old_backup.unlink()
    return target_path
