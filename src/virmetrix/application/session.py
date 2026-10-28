"""Sesión de práctica/examen con temporizador."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Callable


class ExamMode(str, Enum):
    PRACTICE = "practice"
    EXAM = "exam"


class SessionStatus(str, Enum):
    IDLE = "idle"
    RUNNING = "running"
    PAUSED = "paused"
    FINISHED = "finished"


@dataclass
class ExamSession:
    mode: ExamMode
    exercise_ids: list[str]
    duration_seconds: int | None = None  # None => sin límite (práctica)
    current_index: int = 0
    status: SessionStatus = SessionStatus.IDLE
    remaining_seconds: int | None = None
    started_at: str | None = None
    finished_at: str | None = None
    _clock: Callable[[], float] = field(default=lambda: 0.0, repr=False)

    def start(self) -> None:
        if self.status in (SessionStatus.RUNNING, SessionStatus.FINISHED):
            return
        self.status = SessionStatus.RUNNING
        self.started_at = datetime.now(timezone.utc).isoformat()
        if self.mode == ExamMode.EXAM and self.duration_seconds is not None:
            self.remaining_seconds = self.duration_seconds
        else:
            self.remaining_seconds = None

    def pause(self) -> None:
        if self.status == SessionStatus.RUNNING:
            self.status = SessionStatus.PAUSED

    def resume(self) -> None:
        if self.status == SessionStatus.PAUSED:
            self.status = SessionStatus.RUNNING

    def tick(self) -> SessionStatus:
        if self.status != SessionStatus.RUNNING:
            return self.status
        if self.remaining_seconds is not None:
            self.remaining_seconds = max(0, self.remaining_seconds - 1)
            if self.remaining_seconds == 0:
                self.finish()
        return self.status

    def finish(self) -> None:
        if self.status == SessionStatus.FINISHED:
            return
        self.status = SessionStatus.FINISHED
        self.finished_at = datetime.now(timezone.utc).isoformat()

    def next_exercise(self) -> bool:
        if self.current_index < len(self.exercise_ids) - 1:
            self.current_index += 1
            return True
        return False

    def prev_exercise(self) -> bool:
        if self.current_index > 0:
            self.current_index -= 1
            return True
        return False

    @property
    def current_exercise_id(self) -> str | None:
        if 0 <= self.current_index < len(self.exercise_ids):
            return self.exercise_ids[self.current_index]
        return None

    @property
    def is_timed(self) -> bool:
        return self.mode == ExamMode.EXAM and self.duration_seconds is not None

    def format_remaining(self) -> str:
        if self.remaining_seconds is None:
            return "∞"
        m, s = divmod(self.remaining_seconds, 60)
        return f"{m:02d}:{s:02d}"