"""Persistencia local de intentos en SQLite."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from virmetrix.domain.ports import EvaluationResult


@dataclass
class AttemptRecord:
    attempt_id: int
    exercise_id: str
    version: str
    evaluator: str
    started_at: str
    finished_at: str
    duration_seconds: float | None
    score_earned: int
    score_possible: int
    rules_version: str
    criteria: list[dict[str, str]]


class AttemptStore:
    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_schema()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _init_schema(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS attempts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exercise_id TEXT NOT NULL,
                    version TEXT NOT NULL DEFAULT '',
                    evaluator TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    finished_at TEXT NOT NULL,
                    duration_seconds REAL,
                    score_earned INTEGER NOT NULL,
                    score_possible INTEGER NOT NULL,
                    rules_version TEXT NOT NULL DEFAULT ''
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS criterion_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    attempt_id INTEGER NOT NULL,
                    criterion_id TEXT NOT NULL,
                    status TEXT NOT NULL,
                    evidence TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY (attempt_id) REFERENCES attempts(id) ON DELETE CASCADE
                )
                """
            )
            conn.commit()

    def save(
        self,
        result: EvaluationResult,
        *,
        version: str = "",
        started_at: str | None = None,
        finished_at: str | None = None,
        duration_seconds: float | None = None,
    ) -> int:
        now = datetime.now(timezone.utc).isoformat()
        started = started_at or now
        finished = finished_at or now
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO attempts
                (exercise_id, version, evaluator, started_at, finished_at, duration_seconds,
                 score_earned, score_possible, rules_version)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    result.exercise_id,
                    version,
                    result.evaluator,
                    started,
                    finished,
                    duration_seconds,
                    result.score_earned,
                    result.score_possible,
                    result.rules_version,
                ),
            )
            attempt_id = int(cur.lastrowid)
            for c in result.criteria:
                conn.execute(
                    """
                    INSERT INTO criterion_results (attempt_id, criterion_id, status, evidence)
                    VALUES (?, ?, ?, ?)
                    """,
                    (attempt_id, c.criterion_id, c.status.value, c.evidence),
                )
            conn.commit()
        return attempt_id

    def list_attempts(self, exercise_id: str | None = None, limit: int = 50) -> list[AttemptRecord]:
        query = "SELECT * FROM attempts"
        params: list[Any] = []
        if exercise_id:
            query += " WHERE exercise_id = ?"
            params.append(exercise_id)
        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)
        records: list[AttemptRecord] = []
        with self._connect() as conn:
            for row in conn.execute(query, params):
                crit_rows = conn.execute(
                    "SELECT criterion_id, status, evidence FROM criterion_results WHERE attempt_id = ?",
                    (row["id"],),
                ).fetchall()
                criteria = [
                    {"criterion_id": cr["criterion_id"], "status": cr["status"], "evidence": cr["evidence"]}
                    for cr in crit_rows
                ]
                records.append(
                    AttemptRecord(
                        attempt_id=int(row["id"]),
                        exercise_id=row["exercise_id"],
                        version=row["version"],
                        evaluator=row["evaluator"],
                        started_at=row["started_at"],
                        finished_at=row["finished_at"],
                        duration_seconds=row["duration_seconds"],
                        score_earned=int(row["score_earned"]),
                        score_possible=int(row["score_possible"]),
                        rules_version=row["rules_version"],
                        criteria=criteria,
                    )
                )
        return records

    def count(self) -> int:
        with self._connect() as conn:
            return int(conn.execute("SELECT COUNT(*) FROM attempts").fetchone()[0])