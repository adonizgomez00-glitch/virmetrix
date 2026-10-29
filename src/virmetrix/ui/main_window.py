"""UI de VirMetrix (PySide6) — panel de práctica moderno estilo gmetrix."""

from __future__ import annotations

import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QThread, QTimer, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.application.session import ExamMode, ExamSession, SessionStatus
from virmetrix.infrastructure.catalog.loader import list_exercises, load_exercise
from virmetrix.infrastructure.persistence.sqlite import AttemptStore
from virmetrix.ui.worker import EvaluationWorker

ROOT = Path(__file__).resolve().parents[3]
EXAM_DURATION_SECONDS = 300

APP_QSS = """
QWidget#root { background: #eef2f7; }

QFrame#header {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #4338ca, stop:0.55 #2563eb, stop:1 #0891b2);
}
QLabel#title { color: #ffffff; font-size: 22px; font-weight: 800; }
QLabel#subtitle { color: rgba(255, 255, 255, 0.85); font-size: 11px; font-weight: 600; }
QLabel#timer {
    color: #ffffff; font-size: 26px; font-weight: 700;
    font-family: 'Cascadia Mono', 'Consolas', 'DejaVu Sans Mono', monospace;
}
QLabel#timerCaption { color: rgba(255, 255, 255, 0.8); font-size: 9px; font-weight: 700; }

QScrollArea { border: none; background: transparent; }
QWidget#content { background: transparent; }

QFrame#card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
}
QFrame#cardAccent {
    background: #ffffff;
    border: 1px solid #bfdbfe;
    border-left: 4px solid #2563eb;
    border-radius: 12px;
}

QLabel#eyebrow { color: #64748b; font-size: 9px; font-weight: 800; }
QLabel#instructions { color: #0f172a; font-size: 13px; }
QLabel#counter { color: #64748b; font-size: 11px; }
QLabel#score { color: #0f172a; font-size: 30px; font-weight: 800; }
QLabel#scoreCaption { color: #64748b; font-size: 11px; font-weight: 600; }
QLabel#hint { color: #94a3b8; font-size: 10px; }

QLabel#chip {
    background: #f1f5f9;
    color: #94a3b8;
    border: 1px solid #e2e8f0;
    border-radius: 11px;
    padding: 4px 10px;
    font-size: 10px;
    font-weight: 700;
}
QLabel#chip[state="active"] {
    background: #2563eb;
    border-color: #2563eb;
    color: #ffffff;
}
QLabel#chip[state="done"] {
    background: #ecfdf5;
    border-color: #a7f3d0;
    color: #047857;
}

QComboBox {
    background: #f8fafc;
    border: 1px solid #dbe3ee;
    border-radius: 8px;
    padding: 7px 10px;
    color: #0f172a;
    font-size: 12px;
}
QComboBox:hover { border-color: #93c5fd; background: #ffffff; }
QComboBox:focus { border-color: #2563eb; background: #ffffff; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #dbe3ee;
    border-radius: 8px;
    padding: 4px;
    selection-background-color: #eff6ff;
    selection-color: #1d4ed8;
    outline: none;
}

QPushButton {
    background: #f8fafc;
    color: #334155;
    border: 1px solid #dbe3ee;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    font-weight: 600;
}
QPushButton:hover { background: #eff6ff; border-color: #93c5fd; color: #1d4ed8; }
QPushButton:pressed { background: #dbeafe; }
QPushButton:focus { border: 2px solid #2563eb; padding: 7px 11px; }

QPushButton#primary {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #2563eb, stop:1 #0891b2);
    color: #ffffff;
    border: none;
    border-radius: 9px;
    padding: 11px 14px;
    font-size: 13px;
    font-weight: 800;
}
QPushButton#primary:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #1d4ed8, stop:1 #0e7490);
}
QPushButton#primary:pressed { background: #1e40af; }

QPushButton#danger {
    background: #fef2f2;
    border: 1px solid #fecaca;
    color: #dc2626;
}
QPushButton#danger:hover { background: #fee2e2; border-color: #fca5a5; color: #b91c1c; }

QPushButton#ghost {
    background: transparent;
    border: 1px solid #e2e8f0;
    color: #475569;
    font-weight: 600;
}
QPushButton#ghost:hover { background: #f8fafc; border-color: #cbd5e1; color: #1e293b; }

QPushButton:disabled {
    background: #eef2f7;
    border-color: #eef2f7;
    color: #a8b3c2;
}
QPushButton#primary:disabled,
QPushButton#danger:disabled,
QPushButton#ghost:disabled {
    background: #eef2f7;
    border: 1px solid #eef2f7;
    color: #a8b3c2;
}

QTextEdit#results {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 6px;
    color: #334155;
    font-size: 11px;
}

QProgressBar {
    background: #e2e8f0;
    border: none;
    border-radius: 4px;
    max-height: 8px;
    min-height: 8px;
    text-align: center;
    color: transparent;
}
QProgressBar::chunk {
    border-radius: 4px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #2563eb, stop:1 #06b6d4);
}

QCheckBox { color: #64748b; font-size: 11px; spacing: 6px; }
QCheckBox::indicator {
    width: 14px; height: 14px;
    border: 1px solid #cbd5e1; border-radius: 4px;
    background: #ffffff;
}
QCheckBox::indicator:checked { background: #2563eb; border-color: #2563eb; }

QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #cbd5e1;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #94a3b8; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
"""


def _eyebrow(text: str) -> QLabel:
    label = QLabel(text.upper())
    label.setObjectName("eyebrow")
    return label


def _chip(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("chip")
    label.setProperty("state", "idle")
    label.setAlignment(Qt.AlignCenter)
    return label


def _add_shadow(widget: QWidget) -> None:
    effect = QGraphicsDropShadowEffect(widget)
    effect.setBlurRadius(14)
    effect.setOffset(0, 2)
    effect.setColor(QColor(15, 23, 42, 30))
    widget.setGraphicsEffect(effect)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VirMetrix")
        self.resize(460, 780)
        self.setMinimumSize(400, 560)
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setStyleSheet(APP_QSS)

        self.orch = EvaluationOrchestrator()
        self.store = AttemptStore(Path.home() / ".virmetrix" / "history.sqlite")

        self._thread: QThread | None = None
        self._worker: EvaluationWorker | None = None
        self._started_at: str | None = None
        self._start_perf: float | None = None
        self._current_step = 1
        self.work_copy: str | None = None

        self.exercises = list_exercises()
        if not self.exercises:
            raise RuntimeError("Catálogo vacío.")

        self.session = ExamSession(
            mode=ExamMode.PRACTICE,
            exercise_ids=[e["exercise_id"] for e in self.exercises],
        )

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_header())

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget()
        content.setObjectName("content")
        scroll.setWidget(content)
        root_layout.addWidget(scroll, 1)
        self.scroll = scroll

        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 12, 12, 14)
        layout.setSpacing(10)

        layout.addLayout(self._build_steps())
        layout.addWidget(self._build_exercise_card())
        layout.addWidget(self._build_task_card())
        layout.addWidget(self._build_actions_card())
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card)
        layout.addLayout(self._build_footer())
        layout.addStretch(1)

        for card in content.findChildren(QFrame):
            if card.objectName() in ("card", "cardAccent"):
                _add_shadow(card)

        self._load_current_exercise()
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._on_tick)
        self._set_step(1)

    # ------------------------------------------------------------------ UI
    def _build_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        row = QHBoxLayout(header)
        row.setContentsMargins(16, 14, 16, 14)
        row.setSpacing(10)

        brand = QVBoxLayout()
        brand.setSpacing(1)
        title = QLabel("VirMetrix")
        title.setObjectName("title")
        title.setAccessibleName("VirMetrix")
        subtitle = QLabel("MOS WORD · MO-100 · OFFICE 2019")
        subtitle.setObjectName("subtitle")
        brand.addWidget(title)
        brand.addWidget(subtitle)
        row.addLayout(brand)
        row.addStretch(1)

        timer_box = QVBoxLayout()
        timer_box.setSpacing(0)
        caption = QLabel("TIEMPO")
        caption.setObjectName("timerCaption")
        caption.setAlignment(Qt.AlignRight)
        self.timer_label = QLabel("∞")
        self.timer_label.setObjectName("timer")
        self.timer_label.setAccessibleName("Temporizador")
        self.timer_label.setAlignment(Qt.AlignRight)
        timer_box.addWidget(caption)
        timer_box.addWidget(self.timer_label)
        row.addLayout(timer_box)
        return header

    def _build_steps(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(6)
        self.chip_copy = _chip("1 · Copia")
        self.chip_edit = _chip("2 · Editar")
        self.chip_grade = _chip("3 · Evaluar")
        for chip in (self.chip_copy, self.chip_edit, self.chip_grade):
            row.addWidget(chip, 1)
        return row

    def _build_exercise_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(_eyebrow("Ejercicio"))
        head.addStretch(1)
        self.step_label = QLabel("")
        self.step_label.setObjectName("counter")
        self.step_label.setAccessibleName("Paso actual")
        head.addWidget(self.step_label)
        col.addLayout(head)

        self.exercise_combo = QComboBox()
        self.exercise_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        self.exercise_combo.setMinimumContentsLength(12)
        for ex in self.exercises:
            self.exercise_combo.addItem(ex["title"], ex["exercise_id"])
        self.exercise_combo.setAccessibleName("Selector de ejercicio")
        self.exercise_combo.setToolTip("Elige el ejercicio a practicar.")
        self.exercise_combo.currentIndexChanged.connect(self._on_exercise_changed)
        col.addWidget(self.exercise_combo)

        nav = QHBoxLayout()
        nav.setSpacing(8)
        btn_prev = QPushButton("◀  Anterior")
        btn_prev.setObjectName("ghost")
        btn_prev.setAccessibleName("Ejercicio anterior")
        btn_prev.clicked.connect(self.prev_exercise)
        btn_next = QPushButton("Siguiente  ▶")
        btn_next.setObjectName("ghost")
        btn_next.setAccessibleName("Ejercicio siguiente")
        btn_next.clicked.connect(self.next_exercise)
        nav.addWidget(btn_prev)
        nav.addWidget(btn_next)
        col.addLayout(nav)
        return card

    def _build_task_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("cardAccent")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 14)
        col.setSpacing(6)
        col.addWidget(_eyebrow("Tarea"))
        self.instructions_label = QLabel("")
        self.instructions_label.setObjectName("instructions")
        self.instructions_label.setAccessibleName("Instrucciones del ejercicio")
        self.instructions_label.setWordWrap(True)
        self.instructions_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        col.addWidget(self.instructions_label)
        return card

    def _build_actions_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)

        row_mode = QHBoxLayout()
        row_mode.setSpacing(8)
        lbl_mode = QLabel("Modo")
        lbl_mode.setObjectName("counter")
        row_mode.addWidget(lbl_mode)
        self.mode_use_combo = QComboBox()
        self.mode_use_combo.addItem("Práctica (sin tiempo)", ExamMode.PRACTICE)
        self.mode_use_combo.addItem("Examen (cronometrado)", ExamMode.EXAM)
        self.mode_use_combo.setAccessibleName("Modo de uso")
        self.mode_use_combo.currentIndexChanged.connect(self._on_mode_changed)
        row_mode.addWidget(self.mode_use_combo, 1)
        col.addLayout(row_mode)

        row_eval = QHBoxLayout()
        row_eval.setSpacing(8)
        lbl_eval = QLabel("Evaluador")
        lbl_eval.setObjectName("counter")
        row_eval.addWidget(lbl_eval)
        self.eval_mode_combo = QComboBox()
        self.eval_mode_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        self.eval_mode_combo.setMinimumContentsLength(10)
        pretty_modes = {"docx_portable": "Portable (sin Word)", "word_com": "Word COM (real)"}
        for mode, (avail, reason) in self.orch.available_modes().items():
            self.eval_mode_combo.addItem(pretty_modes.get(mode, mode), mode)
            if not avail:
                idx = self.eval_mode_combo.count() - 1
                self.eval_mode_combo.setItemData(idx, f"No disponible: {reason}", Qt.ToolTipRole)
        self.eval_mode_combo.setAccessibleName("Modo de evaluación")
        self.eval_mode_combo.setToolTip("Selecciona el evaluador: portable o Word COM.")
        row_eval.addWidget(self.eval_mode_combo, 1)
        col.addLayout(row_eval)
        self.mode_combo = self.eval_mode_combo

        self.btn_copy = QPushButton("1.  Crear copia de trabajo")
        self.btn_copy.setAccessibleName("Crear copia de trabajo")
        self.btn_copy.setToolTip("Crea una copia de trabajo del documento de referencia.")
        self.btn_copy.clicked.connect(self.create_copy)
        col.addWidget(self.btn_copy)

        self.btn_open = QPushButton("2.  Abrir en Word")
        self.btn_open.setAccessibleName("Abrir documento")
        self.btn_open.setToolTip("Abre la copia de trabajo en Microsoft Word.")
        self.btn_open.clicked.connect(self.open_document)
        col.addWidget(self.btn_open)

        self.btn_eval = QPushButton("3.  Evaluar ahora")
        self.btn_eval.setObjectName("primary")
        self.btn_eval.setAccessibleName("Evaluar")
        self.btn_eval.setToolTip("Evalúa la copia de trabajo.")
        self.btn_eval.setMinimumHeight(40)
        self.btn_eval.clicked.connect(self.evaluate)
        self.btn_eval.setDefault(True)
        col.addWidget(self.btn_eval)

        self.btn_cancel = QPushButton("Cancelar evaluación")
        self.btn_cancel.setObjectName("danger")
        self.btn_cancel.setAccessibleName("Cancelar evaluación")
        self.btn_cancel.clicked.connect(self.cancel_evaluation)
        self.btn_cancel.setVisible(False)
        col.addWidget(self.btn_cancel)

        row_session = QHBoxLayout()
        row_session.setSpacing(6)
        self.btn_start = QPushButton("Iniciar")
        self.btn_start.setObjectName("ghost")
        self.btn_start.setAccessibleName("Iniciar sesión")
        self.btn_start.clicked.connect(self.start_session)
        self.btn_pause = QPushButton("Pausar")
        self.btn_pause.setObjectName("ghost")
        self.btn_pause.setAccessibleName("Pausar o reanudar")
        self.btn_pause.clicked.connect(self.toggle_pause)
        self.btn_pause.setEnabled(False)
        self.btn_finish = QPushButton("Finalizar")
        self.btn_finish.setObjectName("ghost")
        self.btn_finish.setAccessibleName("Finalizar sesión")
        self.btn_finish.clicked.connect(self.finish_session)
        self.btn_finish.setEnabled(False)
        row_session.addWidget(self.btn_start)
        row_session.addWidget(self.btn_pause)
        row_session.addWidget(self.btn_finish)
        col.addLayout(row_session)
        return card

    def _build_result_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(_eyebrow("Resultado"))
        head.addStretch(1)
        score_box = QVBoxLayout()
        score_box.setSpacing(0)
        self.score_label = QLabel("—")
        self.score_label.setObjectName("score")
        self.score_label.setAccessibleName("Puntuación")
        self.score_label.setAlignment(Qt.AlignRight)
        self.score_label.setVisible(False)
        self.score_caption = QLabel("sin evaluar")
        self.score_caption.setObjectName("scoreCaption")
        self.score_caption.setAlignment(Qt.AlignRight)
        score_box.addWidget(self.score_label)
        score_box.addWidget(self.score_caption)
        head.addLayout(score_box)
        col.addLayout(head)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        self.progress.setAccessibleName("Progreso de evaluación")
        col.addWidget(self.progress)

        self.result_view = QTextEdit()
        self.result_view.setObjectName("results")
        self.result_view.setReadOnly(True)
        self.result_view.setAccessibleName("Resultados de la evaluación")
        self.result_view.setMinimumHeight(88)
        self.result_view.setMaximumHeight(150)
        self.result_view.document().setDocumentMargin(6)
        self.result_view.setPlaceholderText(
            "Los resultados de cada criterio aparecerán aquí tras evaluar."
        )
        col.addWidget(self.result_view)
        return card

    def _build_footer(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(8)
        self.stay_on_top = QCheckBox("Mantener encima")
        self.stay_on_top.setAccessibleName("Mantener ventana encima")
        self.stay_on_top.setChecked(True)
        self.stay_on_top.stateChanged.connect(self._toggle_stay_on_top)
        row.addWidget(self.stay_on_top)
        row.addStretch(1)
        btn_about = QPushButton("Acerca de")
        btn_about.setObjectName("ghost")
        btn_about.setAccessibleName("Acerca de VirMetrix")
        btn_about.clicked.connect(self._show_about)
        row.addWidget(btn_about)
        return row

    # ---------------------------------------------------------------- steps
    def _set_step(self, step: int) -> None:
        self._current_step = step
        states = {
            1: (("active", "idle", "idle")),
            2: (("done", "active", "idle")),
            3: (("done", "done", "active")),
        }.get(step, ("idle", "idle", "idle"))
        for chip, state in zip(
            (self.chip_copy, self.chip_edit, self.chip_grade), states
        ):
            chip.setProperty("state", state)
            chip.style().unpolish(chip)
            chip.style().polish(chip)

    # ------------------------------------------------------------- lifecycle
    def _toggle_stay_on_top(self, state):
        if state:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()

    def _load_current_exercise(self):
        eid = self.session.current_exercise_id
        if eid is None:
            return
        idx = self.exercise_combo.findData(eid)
        if idx >= 0:
            self.exercise_combo.blockSignals(True)
            self.exercise_combo.setCurrentIndex(idx)
            self.exercise_combo.blockSignals(False)
        self.exercise = load_exercise(eid)
        self.instructions_label.setText(self.exercise.instructions)
        self.step_label.setText(
            f"Ejercicio {self.session.current_index + 1} de {len(self.session.exercise_ids)}"
        )
        self.work_copy = None
        self._reset_score()

    def _reference_path(self) -> Path:
        return ROOT / "exercises" / self.session.current_exercise_id / "reference.docx"

    def _on_mode_changed(self, _index):
        self.session.mode = self.mode_use_combo.currentData()
        if self.session.mode == ExamMode.EXAM:
            self.session.duration_seconds = EXAM_DURATION_SECONDS
        else:
            self.session.duration_seconds = None
        self.timer_label.setText(f"{self.session.format_remaining()}")

    def _on_exercise_changed(self, _index):
        eid = self.exercise_combo.currentData()
        if eid and eid in self.session.exercise_ids:
            self.session.current_index = self.session.exercise_ids.index(eid)
        self._load_current_exercise()
        self.work_copy = None
        self._set_step(1)
        self.result_view.setPlainText(f"Ejercicio seleccionado: {eid}")

    # ---------------------------------------------------------------- actions
    def create_copy(self):
        ref = self._reference_path()
        if not ref.is_file():
            QMessageBox.critical(self, "Error", "Falta el documento de referencia.")
            return
        work_dir = Path.home() / ".virmetrix" / "work"
        self.work_copy = self.orch.create_work_copy(str(ref), str(work_dir))
        self.result_view.setPlainText(f"Copia de trabajo creada:\n{self.work_copy}")
        self._set_step(2)
        self._bring_to_front()

    def open_document(self):
        if not self.work_copy:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        try:
            self.orch.open_document(self.work_copy)
            self.result_view.setPlainText(f"Documento abierto:\n{self.work_copy}")
            self._set_step(3)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Error", f"No se pudo abrir el documento: {e}")
        self._bring_to_front()

    def _bring_to_front(self):
        self.raise_()
        self.activateWindow()
        try:
            if sys.platform == "win32":
                import ctypes

                hwnd = int(self.winId())
                ctypes.windll.user32.SetWindowPos(
                    hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040
                )
        except Exception:
            pass

    def _scroll_to_result(self) -> None:
        QTimer.singleShot(
            0, lambda: self.scroll.ensureWidgetVisible(self.result_card, 0, 8)
        )

    def start_session(self):
        if self.session.status == SessionStatus.FINISHED:
            self.session = ExamSession(
                mode=self.mode_use_combo.currentData(),
                exercise_ids=[e["exercise_id"] for e in self.exercises],
                duration_seconds=(
                    EXAM_DURATION_SECONDS
                    if self.mode_use_combo.currentData() == ExamMode.EXAM
                    else None
                ),
            )
            self._load_current_exercise()
        self.session.start()
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_finish.setEnabled(True)
        if self.session.is_timed:
            self._timer.start()
        self.timer_label.setText(f"{self.session.format_remaining()}")
        self.result_view.setPlainText(f"Sesión iniciada ({self.session.mode.value}).")

    def toggle_pause(self):
        if self.session.status == SessionStatus.RUNNING:
            self.session.pause()
            self._timer.stop()
            self.btn_pause.setText("Reanudar")
            self.result_view.setPlainText("Sesión pausada.")
        elif self.session.status == SessionStatus.PAUSED:
            self.session.resume()
            if self.session.is_timed:
                self._timer.start()
            self.btn_pause.setText("Pausar")
            self.result_view.setPlainText("Sesión reanudada.")
        self.timer_label.setText(f"{self.session.format_remaining()}")

    def _on_tick(self):
        self.session.tick()
        self.timer_label.setText(f"{self.session.format_remaining()}")
        if self.session.status == SessionStatus.FINISHED:
            self._timer.stop()
            self._finish_ui()

    def next_exercise(self):
        if self.session.next_exercise():
            self._load_current_exercise()
            self.work_copy = None
            self._set_step(1)

    def prev_exercise(self):
        if self.session.prev_exercise():
            self._load_current_exercise()
            self.work_copy = None
            self._set_step(1)

    # ------------------------------------------------------------- evaluation
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
        self.btn_cancel.setVisible(True)
        self.progress.setVisible(True)
        self.result_view.setPlainText("Evaluando… esto puede tardar unos segundos.")

        self._started_at = datetime.now(timezone.utc).isoformat()
        self._start_perf = time.perf_counter()
        self._scroll_to_result()

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

    def _reset_score(self) -> None:
        self.score_label.setText("—")
        self.score_label.setStyleSheet("")
        self.score_label.setVisible(False)
        self.score_caption.setText("sin evaluar")

    def _show_score(self, earned: int, possible: int) -> None:
        self.score_label.setVisible(True)
        self.score_label.setText(f"{earned}/{possible}")
        ratio = (earned / possible) if possible else 0.0
        if ratio >= 1.0:
            color, caption = "#059669", "completado"
        elif ratio >= 0.6:
            color, caption = "#d97706", "casi"
        else:
            color, caption = "#dc2626", "a mejorar"
        self.score_label.setStyleSheet(f"color: {color};")
        self.score_caption.setText(f"{caption} · {ratio * 100:.0f}%")

    def _render_result(self, result) -> str:
        def esc(text: str) -> str:
            return (
                str(text)
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )

        palette = {
            "pass": ("#059669", "✔"),
            "fail": ("#dc2626", "✖"),
            "not_verifiable": ("#64748b", "•"),
            "error": ("#b45309", "!"),
        }
        parts = [
            f'<div style="color:#334155;font-weight:600;">'
            f"Evaluador: {esc(result.evaluator)}</div>"
        ]
        if result.error_message:
            parts.append(
                f'<div style="color:#b45309;margin-top:4px;">'
                f"Error: {esc(result.error_message)}</div>"
            )
        for c in result.criteria:
            color, icon = palette.get(c.status.value, ("#64748b", "•"))
            parts.append(
                f'<div style="margin-top:4px;color:{color};">'
                f"{icon} <b>{esc(c.criterion_id)}</b> · {esc(c.status.value)}"
                f'<span style="color:#64748b;"> — {esc(c.evidence)}</span></div>'
            )
        parts.append(
            '<div style="margin-top:8px;color:#94a3b8;font-size:10px;">'
            f"Historial guardado · intentos totales: {self.store.count()}</div>"
        )
        return "".join(parts)

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

        self.progress.setVisible(False)
        self.btn_cancel.setVisible(False)
        self.result_view.setHtml(self._render_result(result))
        self._show_score(result.score_earned, result.score_possible)
        self._set_step(3)
        self.btn_eval.setEnabled(True)
        self._scroll_to_result()
        self._bring_to_front()

    def _on_failed(self, msg):
        self.progress.setVisible(False)
        self.btn_cancel.setVisible(False)
        self.result_view.setPlainText(f"Error: {msg}")
        self._reset_score()
        self.score_caption.setText("error")
        self.btn_eval.setEnabled(True)

    def cancel_evaluation(self):
        if self._worker is not None:
            try:
                self._worker.cancel()
            except Exception:
                pass
        if self._thread is not None and self._thread.isRunning():
            self._thread.wait(2000)
        self.progress.setVisible(False)
        self.btn_cancel.setVisible(False)
        self.btn_eval.setEnabled(True)
        self.result_view.setPlainText("Evaluación cancelada.")

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
        self.timer_label.setText(f"{self.session.format_remaining()}")

    def _show_about(self):
        QMessageBox.information(
            self,
            "Acerca de VirMetrix",
            "VirMetrix — simulador de práctica para MOS Word Associate "
            "(MO-100, Office 2019).\n"
            "Proyecto independiente, no afiliado a Microsoft.\n"
            "Versión: 0.1.0 (ITER-021)\n"
            "Licencia: gratuito para práctica y autoevaluación.",
        )


def run():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
