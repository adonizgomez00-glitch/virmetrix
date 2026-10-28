"""Worker para ejecutar evaluación fuera del hilo principal de UI."""

from __future__ import annotations

import multiprocessing
import time
from pathlib import Path

from PySide6.QtCore import QObject, Signal

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.domain.ports import EvaluationResult, ExerciseMetadata
from virmetrix.infrastructure.evaluators.com_process import evaluate_com


class EvaluationWorker(QObject):
    finished = Signal(object)  # EvaluationResult
    failed = Signal(str)
    progress = Signal(int)  # 0..100 (best-effort)

    def __init__(self, orchestrator: EvaluationOrchestrator, mode: str, work_file_path: str, exercise: ExerciseMetadata):
        super().__init__()
        self.orchestrator = orchestrator
        self.mode = mode
        self.work_file_path = work_file_path
        self.exercise = exercise
        self._process: multiprocessing.Process | None = None
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True
        if self._process is not None and self._process.is_alive():
            self._process.terminate()

    def run(self) -> None:
        try:
            if self.mode == "word_com":
                result = self._run_com()
            else:
                result = self.orchestrator.evaluate(self.mode, self.work_file_path, self.exercise)
            if self._cancelled:
                return
            self.finished.emit(result)
        except Exception as e:  # noqa: BLE001
            if not self._cancelled:
                self.failed.emit(f"{type(e).__name__}: {e}")

    def _run_com(self) -> EvaluationResult:
        ctx = multiprocessing.get_context("spawn")
        q: multiprocessing.Queue = ctx.Queue()
        self._process = ctx.Process(
            target=_com_target, args=(q, self.work_file_path, self.exercise)
        )
        self._process.start()
        # Wait with polling to allow cancellation.
        while self._process.is_alive():
            if self._cancelled:
                self._process.terminate()
                self._process.join(timeout=2)
                raise RuntimeError("Cancelado")
            self._process.join(timeout=0.1)
        if self._cancelled:
            raise RuntimeError("Cancelado")
        if q.empty():
            raise RuntimeError("El proceso COM no devolvió resultado")
        return q.get()


def _com_target(q, work_file_path, exercise) -> None:
    try:
        result = evaluate_com(work_file_path, exercise)
        q.put(result)
    except Exception as e:  # noqa: BLE001
        q.put(e)
    finally:
        q.close()
