"""Helper para ejecutar evaluacion Word COM en un proceso separado.

Word COM (STA) se bloquea si se invoca desde un hilo secundario sin bomba de
mensajes. Ejecutar en un proceso propio evita bloquear la UI.
"""

from __future__ import annotations

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.domain.ports import EvaluationResult, ExerciseMetadata


def evaluate_com(work_file_path: str, exercise: ExerciseMetadata) -> EvaluationResult:
    orch = EvaluationOrchestrator()
    return orch.evaluate("word_com", work_file_path, exercise)
