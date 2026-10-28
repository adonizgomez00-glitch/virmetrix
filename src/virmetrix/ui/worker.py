"""Worker para ejecutar evaluacion fuera del hilo principal de UI."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import QObject, Signal

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.domain.ports import CriterionResult, CriterionStatus, EvaluationResult, ExerciseMetadata


class EvaluationWorker(QObject):
    finished = Signal(object)  # EvaluationResult
    failed = Signal(str)

    def __init__(self, orchestrator: EvaluationOrchestrator, mode: str, work_file_path: str, exercise: ExerciseMetadata):
        super().__init__()
        self.orchestrator = orchestrator
        self.mode = mode
        self.work_file_path = str(Path(work_file_path).resolve())
        self.exercise = exercise
        self._proc: subprocess.Popen | None = None
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True
        if self._proc is not None and self._proc.poll() is None:
            try:
                self._proc.terminate()
            except Exception:
                pass

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
        helper = Path(__file__).resolve().parents[1] / "infrastructure" / "evaluators" / "com_eval_cli.py"
        cmd = [sys.executable, str(helper), self.work_file_path, self.exercise.exercise_id]
        self._proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        # Poll to allow cancellation.
        while self._proc.poll() is None:
            if self._cancelled:
                self._proc.terminate()
                self._proc.wait(timeout=2)
                raise RuntimeError("Cancelado")
            import time
            time.sleep(0.1)
        if self._cancelled:
            raise RuntimeError("Cancelado")
        stdout = self._proc.stdout.read() if self._proc.stdout else ""
        stderr = self._proc.stderr.read() if self._proc.stderr else ""
        if self._proc.returncode != 0:
            raise RuntimeError(stderr.strip() or "Error en proceso COM")
        data = json.loads(stdout)
        if data.get("status") == "error" and data.get("error_message"):
            # Build error result
            pass
        criteria = [
            CriterionResult(
                criterion_id=c["criterion_id"],
                status=CriterionStatus(c["status"]),
                evidence=c["evidence"],
            )
            for c in data["criteria"]
        ]
        return EvaluationResult(
            exercise_id=data["exercise_id"],
            evaluator=data["evaluator"],
            status=data["status"],
            criteria=criteria,
            score_earned=data["score_earned"],
            score_possible=data["score_possible"],
            rules_version=data["rules_version"],
            error_message=data.get("error_message", ""),
        )
