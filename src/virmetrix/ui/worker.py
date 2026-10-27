"""Worker para ejecutar evaluación fuera del hilo principal de UI."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.domain.ports import EvaluationResult, ExerciseMetadata


class EvaluationWorker(QObject):
    finished = Signal(object)  # EvaluationResult
    failed = Signal(str)

    def __init__(self, orchestrator: EvaluationOrchestrator, mode: str, work_file_path: str, exercise: ExerciseMetadata):
        super().__init__()
        self.orchestrator = orchestrator
        self.mode = mode
        self.work_file_path = work_file_path
        self.exercise = exercise

    def run(self) -> None:
        try:
            result = self.orchestrator.evaluate(self.mode, self.work_file_path, self.exercise)
            self.finished.emit(result)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"{type(e).__name__}: {e}")