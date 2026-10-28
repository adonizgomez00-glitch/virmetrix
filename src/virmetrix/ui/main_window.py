"""UI de VirMetrix (PySide6) — catálogo, sesión y temporizador."""

from __future__ import annotations

import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QThread, QTimer, Qt
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
from virmetrix.application.session import ExamMode, ExamSession, SessionStatus
from virmetrix.infrastructure.catalog.loader import list_exercises, load_exercise
from virmetrix.infrastructure.persistence.sqlite import AttemptStore
from virmetrix.ui.worker import EvaluationWorker

from virmetrix.paths import resource_root

ROOT = resource_root()
EXAM_DURATION_SECONDS = 300


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

        self.session = ExamSession(
            mode=ExamMode.PRACTICE,
            exercise_ids=[e["exercise_id"] for e in self.exercises],
        )

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Modo de uso
        layout.addWidget(QLabel("Modo:"))
        self.mode_use_combo = QComboBox()
        self.mode_use_combo.addItem("Práctica (sin tiempo)", ExamMode.PRACTICE)
        self.mode_use_combo.addItem("Examen (cronometrado)", ExamMode.EXAM)
        self.mode_use_combo.setAccessibleName("Modo de uso")
        self.mode_use_combo.setToolTip("Selecciona práctica (sin tiempo) o examen (cronometrado).")
        self.mode_use_combo.currentIndexChanged.connect(self._on_mode_changed)
        layout.addWidget(self.mode_use_combo)

        # Ejercicio
        layout.addWidget(QLabel("Ejercicio:"))
        self.exercise_combo = QComboBox()
        for ex in self.exercises:
            self.exercise_combo.addItem(ex["title"], ex["exercise_id"])
        self.exercise_combo.setAccessibleName("Selector de ejercicio")
        self.exercise_combo.setToolTip("Elige el ejercicio a practicar.")
        self.exercise_combo.currentIndexChanged.connect(self._on_exercise_changed)
        layout.addWidget(self.exercise_combo)

        self.instructions_label = QLabel("")
        self.instructions_label.setAccessibleName("Instrucciones del ejercicio")
        self.instructions_label.setWordWrap(True)
        layout.addWidget(self.instructions_label)

        # Temporizador
        self.timer_label = QLabel("Tiempo: ∞")
        self.timer_label.setAccessibleName("Temporizador")
        layout.addWidget(self.timer_label)

        # Modo de evaluación
        self.eval_mode_combo = QComboBox()
        for mode, (avail, reason) in self.orch.available_modes().items():
            label = f"{mode} — {'disponible' if avail else 'NO disponible: ' + reason}"
            self.eval_mode_combo.addItem(label, mode)
        self.eval_mode_combo.setAccessibleName("Modo de evaluación")
        self.eval_mode_combo.setToolTip("Selecciona el evaluador: portable o Word COM.")
        layout.addWidget(self.eval_mode_combo)
        self.mode_combo = self.eval_mode_combo  # alias for tests/compatibility

        # Botones
        btn_copy = QPushButton("Crear copia de trabajo")
        btn_copy.setAccessibleName("Crear copia de trabajo")
        btn_copy.setToolTip("Crea una copia de trabajo del documento de referencia.")
        btn_copy.clicked.connect(self.create_copy)
        layout.addWidget(btn_copy)

        btn_open = QPushButton("Abrir documento")
        btn_open.setAccessibleName("Abrir documento")
        btn_open.setToolTip("Abre la copia de trabajo en Microsoft Word (o el programa predeterminado).")
        btn_open.clicked.connect(self.open_document)
        layout.addWidget(btn_open)

        nav_layout = QVBoxLayout()
        btn_prev = QPushButton("Anterior")
        btn_prev.setAccessibleName("Ejercicio anterior")
        btn_prev.setToolTip("Ir al ejercicio anterior.")
        btn_prev.clicked.connect(self.prev_exercise)
        nav_layout.addWidget(btn_prev)
        btn_next = QPushButton("Siguiente")
        btn_next.setAccessibleName("Ejercicio siguiente")
        btn_next.setToolTip("Ir al ejercicio siguiente.")
        btn_next.clicked.connect(self.next_exercise)
        nav_layout.addWidget(btn_next)
        layout.addLayout(nav_layout)

        self.btn_start = QPushButton("Iniciar sesión")
        self.btn_start.setAccessibleName("Iniciar sesión")
        self.btn_start.setToolTip("Inicia la sesión de práctica o examen.")
        self.btn_start.clicked.connect(self.start_session)
        layout.addWidget(self.btn_start)

        self.btn_pause = QPushButton("Pausar")
        self.btn_pause.setAccessibleName("Pausar o reanudar")
        self.btn_pause.setToolTip("Pausa o reanuda la sesión.")
        self.btn_pause.clicked.connect(self.toggle_pause)
        self.btn_pause.setEnabled(False)
        layout.addWidget(self.btn_pause)

        self.btn_eval = QPushButton("Evaluar")
        self.btn_eval.setAccessibleName("Evaluar")
        self.btn_eval.setToolTip("Evalúa la copia de trabajo (Enter).")
        self.btn_eval.clicked.connect(self.evaluate)
        self.btn_eval.setDefault(True)
        layout.addWidget(self.btn_eval)

        self.btn_finish = QPushButton("Finalizar")
        self.btn_finish.setAccessibleName("Finalizar sesión")
        self.btn_finish.setToolTip("Finaliza la sesión actual.")
        self.btn_finish.clicked.connect(self.finish_session)
        self.btn_finish.setEnabled(False)
        layout.addWidget(self.btn_finish)

        btn_about = QPushButton("Acerca de")
        btn_about.setAccessibleName("Acerca de VirMetrix")
        btn_about.setToolTip("Información de la aplicación.")
        btn_about.clicked.connect(self._show_about)
        layout.addWidget(btn_about)

        self.result_view = QTextEdit()
        self.result_view.setReadOnly(True)
        self.result_view.setAccessibleName("Resultados de la evaluación")
        layout.addWidget(self.result_view)

        self._load_current_exercise()
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._on_tick)

    def _load_current_exercise(self):
        eid = self.session.current_exercise_id
        if eid is None:
            return
        # sync combo
        idx = self.exercise_combo.findData(eid)
        if idx >= 0:
            self.exercise_combo.blockSignals(True)
            self.exercise_combo.setCurrentIndex(idx)
            self.exercise_combo.blockSignals(False)
        self.exercise = load_exercise(eid)
        self.instructions_label.setText(self.exercise.instructions)
        self.work_copy = None

    def _reference_path(self) -> Path:
        return ROOT / "exercises" / self.session.current_exercise_id / "reference.docx"

    def _on_mode_changed(self, _index):
        self.session.mode = self.mode_use_combo.currentData()
        if self.session.mode == ExamMode.EXAM:
            self.session.duration_seconds = EXAM_DURATION_SECONDS
        else:
            self.session.duration_seconds = None
        self.timer_label.setText(f"Tiempo: {self.session.format_remaining()}")

    def _on_exercise_changed(self, _index):
        eid = self.exercise_combo.currentData()
        if eid and eid != self.session.current_exercise_id:
            # keep session index in sync if possible
            if eid in self.session.exercise_ids:
                self.session.current_index = self.session.exercise_ids.index(eid)
        self._load_current_exercise()
        self.work_copy = None
        self.result_view.append(f"Ejercicio seleccionado: {eid}")

    def create_copy(self):
        ref = self._reference_path()
        if not ref.is_file():
            QMessageBox.critical(self, "Error", "Falta el documento de referencia.")
            return
        work_dir = Path.home() / ".virmetrix" / "work"
        self.work_copy = self.orch.create_work_copy(str(ref), str(work_dir))
        self.result_view.append(f"Copia de trabajo creada: {self.work_copy}")
        try:
            self.orch.open_document(self.work_copy)
            self.result_view.append("Documento abierto en el programa predeterminado.")
        except Exception as e:  # noqa: BLE001
            self.result_view.append(f"No se pudo abrir automáticamente: {e}")

    def start_session(self):
        if self.session.status == SessionStatus.FINISHED:
            # reset
            self.session = ExamSession(
                mode=self.mode_use_combo.currentData(),
                exercise_ids=[e["exercise_id"] for e in self.exercises],
                duration_seconds=EXAM_DURATION_SECONDS if self.mode_use_combo.currentData() == ExamMode.EXAM else None,
            )
            self._load_current_exercise()
        self.session.start()
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_finish.setEnabled(True)
        if self.session.is_timed:
            self._timer.start()
        self.timer_label.setText(f"Tiempo: {self.session.format_remaining()}")
        self.result_view.append(f"Sesión iniciada ({self.session.mode.value}).")

    def toggle_pause(self):
        if self.session.status == SessionStatus.RUNNING:
            self.session.pause()
            self._timer.stop()
            self.btn_pause.setText("Reanudar")
            self.result_view.append("Sesión pausada.")
        elif self.session.status == SessionStatus.PAUSED:
            self.session.resume()
            if self.session.is_timed:
                self._timer.start()
            self.btn_pause.setText("Pausar")
            self.result_view.append("Sesión reanudada.")
        self.timer_label.setText(f"Tiempo: {self.session.format_remaining()}")

    def _on_tick(self):
        self.session.tick()
        self.timer_label.setText(f"Tiempo: {self.session.format_remaining()}")
        if self.session.status == SessionStatus.FINISHED:
            self._timer.stop()
            self._finish_ui()

    def next_exercise(self):
        if self.session.next_exercise():
            self._load_current_exercise()
            self.work_copy = None

    def prev_exercise(self):
        if self.session.prev_exercise():
            self._load_current_exercise()
            self.work_copy = None

    def open_document(self):
        if not self.work_copy:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        try:
            self.orch.open_document(self.work_copy)
            self.result_view.append(f"Abriendo: {self.work_copy}")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Error", f"No se pudo abrir el documento: {e}")

    def evaluate(self):
        if not self.work_copy:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        mode = self.eval_mode_combo.currentData()
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

    def finish_session(self):
        if self.session.status == SessionStatus.FINISHED:
            return
        self.session.finish()
        self._timer.stop()
        self._finish_ui()
        self.result_view.append("Sesión finalizada.")

    def _finish_ui(self):
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_finish.setEnabled(False)
        self.btn_pause.setText("Pausar")
        self.timer_label.setText(f"Tiempo: {self.session.format_remaining()}")


    def _show_about(self):
        QMessageBox.information(
            self,
            "Acerca de VirMetrix",
            "VirMetrix — simulador de práctica para MOS Word Associate (MO-100, Office 2019).\n"
            "Proyecto independiente, no afiliado a Microsoft.\n"
            "Versión: 0.1.0 (ITER-006)\n"
            "Licencia: por definir.",
        )


def run():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
