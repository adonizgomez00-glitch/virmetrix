"""Contratos de dominio (puertos) para VirMetrix.

Definidos en Fase 2 (Architect) de ITER-001. Implementaciones concretas viven en
adaptadores (infraestructura), no en el núcleo.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Protocol, runtime_checkable


class CriterionStatus(str, Enum):
    PASS = "pass"
    FAIL = "fail"
    NOT_VERIFIABLE = "not_verifiable"
    ERROR = "error"


@dataclass(frozen=True)
class CriterionResult:
    criterion_id: str
    status: CriterionStatus
    evidence: str


@dataclass(frozen=True)
class EvaluationResult:
    exercise_id: str
    evaluator: str  # "docx_portable" | "word_com"
    status: str  # "completed" | "error"
    criteria: list[CriterionResult] = field(default_factory=list)
    score_earned: int = 0
    score_possible: int = 0
    evaluated_at: str = ""  # ISO 8601 local
    rules_version: str = ""
    error_message: str = ""


@runtime_checkable
class EvaluationAdapter(Protocol):
    """Puerto común para evaluadores portable y Word COM."""

    def is_available(self) -> tuple[bool, str]:
        """Devuelve (disponible, motivo)."""
        ...

    def evaluate(self, work_file_path: str, exercise: "ExerciseMetadata") -> EvaluationResult:
        ...


@runtime_checkable
class AttemptStore(Protocol):
    """Puerto de persistencia local de intentos (implementación: SQLite)."""

    def save(self, result: "EvaluationResult", **kwargs: object) -> int:
        ...

    def list_attempts(self, exercise_id: str | None = None, limit: int = 50) -> list[object]:
        ...


@dataclass(frozen=True)
class ExerciseMetadata:
    exercise_id: str
    version: str
    language: str
    competency: str
    instructions: str
    criteria: list[str]  # IDs C1..Cn
    rules_version: str
    target_phrase: str | None = None
    expected_text: str | None = None
    criteria_specs: list[dict] = field(default_factory=list)
