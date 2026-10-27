"""UI mínima de VirMetrix (PySide6) para el vertical slice word-format-001."""

from __future__ import annotations

import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QThread, Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.infrastructure.catalog.loader import list_exercises, load_exercise
from virmetrix.infrastructure.persistence.sqlite import AttemptStore
from virmetrix.ui.worker import EvaluationWorker

ROOT = Path(__file__).resolve().parents[3]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VirMetrix")
        self.orch = EvaluationOrchestrator()
        self.store = AttemptStore(Path.home() / ".virmetrix" / "history.sqlite")

        self._thread: QThread | None = None
        self._worker: EvaluationWorker | None = None
        self._started_at: str | None = None
        self._start_perf: float | None = None
        self.work_copy: str | None = None

        self.exercises = list_exercises()
        if not self.exercises:
            raise RuntimeError("Catálogo vacío.")
        self.current_exercise_id = self.exercises[0]["exercise_id"]
        self.exercise = load_exercise(self.current_exercise_id)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(QLabel("Ejercicio:"))
        self.exercise_combo = QComboBox()
        for ex in self.exercises:
            self.exercise_combo.addItem(ex["title"], ex["exercise_id"])
        self.exercise_combo.currentIndexChanged.connect(self._on_exercise_changed)
        layout.addWidget(self.exercise_combo)

        layout.addWidget(QLabel(self.exercise.instructions))

        self.instructions_label = layout.itemAt(layout.count() - 1).widget()

        self.mode_combo = QComboBox()
        for mode, (avail, reason) in self.orch.available_modes().items():
            label = f"{mode} — {'disponible' if avail else 'NO disponible: ' + reason}"
            self.mode_combo.addItem(label, mode)
        layout.addWidget(self.mode_combo)

        btn_copy = QPushButton("Crear copia de trabajo")
        btn_copy.clicked.connect(self.create_copy)
        layout.addWidget(btn_copy)

        self.btn_eval = QPushButton("Evaluar")
        self.btn_eval.clicked.connect(self.evaluate)
        layout.addWidget(self.btn_eval)

        self.result_view = QTextEdit()
        self.result_view.setReadOnly(True)
        layout.addWidget(self.result_view)

    def _reference_path(self) -> Path:
        return ROOT / "exercises" / self.current_exercise_id / "reference.docx"

    def _on_exercise_changed(self, _index):
        self.current_exercise_id = self.exercise_combo.currentData()
        self.exercise = load_exercise(self.current_exercise_id)
        self.work_copy = None
        self.instructions_label.setText(self.exercise.instructions)
        self.result_view.append(f"Ejercicio seleccionado: {self.current_exercise_id}")

    def create_copy(self):
        ref = self._reference_path()
        if not ref.is_file():
            QMessageBox.critical(self, "Error", "Falta el documento de referencia.")
            return
        work_dir = Path.home() / ".virmetrix" / "work"
        self.work_copy = self.orch.create_work_copy(str(ref), str(work_dir))
        self.result_view.append(f"Copia de trabajo creada: {self.work_copy}")

    def evaluate(self):
        if not self.work_copy:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        mode = self.mode_combo.currentData()
        avail, reason = self.orch.available_modes()[mode]
        if not avail:
            QMessageBox.warning(self, "Modo no disponible", reason)
            return

        self.btn_eval.setEnabled(False)
        self.result_view.append("Evaluando...")

        self._started_at = datetime.now(timezone.utc).isoformat()
        self._start_perf = time.perf_counter()

        self._thread = QThread()
        self._worker = EvaluationWorker(self.orch, mode, self.work_copy, self.exercise)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._worker.failed.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _on_finished(self, result):
        duration = None
        if self._start_perf is not None:
            duration = time.perf_counter() - self._start_perf
        try:
            self.store.save(
                result,
                version=self.exercise.version,
                started_at=self._started_at,
                finished_at=datetime.now(timezone.utc).isoformat(),
                duration_seconds=duration,
            )
        except Exception as e:  # noqa: BLE001
            self.result_view.append(f"Aviso: no se pudo guardar historial: {e}")

        lines = [f"Evaluador: {result.evaluator}", f"Puntuación: {result.score_earned}/{result.score_possible}"]
        if result.error_message:
            lines.append(f"Error: {result.error_message}")
        for c in result.criteria:
            lines.append(f"{c.criterion_id}: {c.status.value} — {c.evidence}")
        lines.append(f"Historial guardado. Intentos totales: {self.store.count()}")
        self.result_view.setPlainText("\n".join(lines))
        self.btn_eval.setEnabled(True)

    def _on_failed(self, msg):
        self.result_view.setPlainText(f"Error: {msg}")
        self.btn_eval.setEnabled(True)


def run():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
