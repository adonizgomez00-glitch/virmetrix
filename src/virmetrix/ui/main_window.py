"""UI de VirMetrix (PySide6) — configuración (vertical) + práctica (horizontal)."""

from __future__ import annotations

import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QThread, QTimer, Qt
from PySide6.QtGui import QColor, QIcon
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
from virmetrix.paths import resource_root
from virmetrix.ui.worker import EvaluationWorker

# resource_root() funciona en desarrollo y en el ejecutable empaquetado:
# en modo congelado (PyInstaller) apunta a _internal/, donde se copian los
# ejercicios. Path(__file__).resolve().parents[3] en congelado apunta un
# nivel mas arriba (dist\VirMetrix) y el reference.docx no se encontraba.
ROOT = resource_root()
EXAM_DURATION_SECONDS = 300

APP_QSS = """
QWidget#root { background: #eef2f7; }
QFrame#header { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #4338ca, stop:0.55 #2563eb, stop:1 #0891b2); }
QLabel#title { color: #ffffff; font-size: 24px; font-weight: 900; }
QLabel#subtitle { color: rgba(255, 255, 255, 0.85); font-size: 11px; font-weight: 600; }
QLabel#timer { color: #ffffff; font-size: 28px; font-weight: 800; font-family: 'Cascadia Mono', 'Consolas', 'DejaVu Sans Mono', monospace; }
QLabel#timerCaption { color: rgba(255, 255, 255, 0.8); font-size: 9px; font-weight: 700; }
QScrollArea { border: none; background: transparent; }
QWidget#content { background: transparent; }
QFrame#card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; }
QFrame#cardAccent { background: #ffffff; border: 1px solid #bfdbfe; border-left: 4px solid #2563eb; border-radius: 12px; }
QLabel#eyebrow { color: #64748b; font-size: 9px; font-weight: 800; }
QLabel#instructions { color: #0f172a; font-size: 13px; }
QLabel#counter { color: #64748b; font-size: 11px; }
QLabel#score { color: #0f172a; font-size: 30px; font-weight: 800; }
QLabel#scoreCaption { color: #64748b; font-size: 11px; font-weight: 600; }
QLabel#hint { color: #94a3b8; font-size: 10px; }
QLabel#chip { background: #f1f5f9; color: #94a3b8; border: 1px solid #e2e8f0; border-radius: 11px; padding: 4px 10px; font-size: 10px; font-weight: 700; }
QLabel#chip[state="active"] { background: #2563eb; border-color: #2563eb; color: #ffffff; }
QLabel#chip[state="done"] { background: #ecfdf5; border-color: #a7f3d0; color: #047857; }
QComboBox { background: #f8fafc; border: 1px solid #dbe3ee; border-radius: 8px; padding: 7px 10px; color: #0f172a; font-size: 12px; }
QComboBox:hover { border-color: #93c5fd; background: #ffffff; }
QComboBox:focus { border-color: #2563eb; background: #ffffff; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #dbe3ee; border-radius: 8px; padding: 4px; selection-background-color: #eff6ff; selection-color: #1d4ed8; outline: none; }
QPushButton { background: #f8fafc; color: #334155; border: 1px solid #dbe3ee; border-radius: 8px; padding: 8px 12px; font-size: 12px; font-weight: 600; }
QPushButton:hover { background: #eff6ff; border-color: #93c5fd; color: #1d4ed8; }
QPushButton:pressed { background: #dbeafe; }
QPushButton:focus { border: 2px solid #2563eb; padding: 7px 11px; }
QPushButton#primary { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #0891b2); color: #ffffff; border: none; border-radius: 9px; padding: 11px 14px; font-size: 13px; font-weight: 800; }
QPushButton#primary:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4ed8, stop:1 #0e7490); }
QPushButton#primary:pressed { background: #1e40af; }
QPushButton#danger { background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; }
QPushButton#danger:hover { background: #fee2e2; border-color: #fca5a5; color: #b91c1c; }
QPushButton#ghost { background: transparent; border: 1px solid #e2e8f0; color: #475569; font-weight: 600; }
QPushButton#ghost:hover { background: #f8fafc; border-color: #cbd5e1; color: #1e293b; }
QPushButton:disabled { background: #eef2f7; border-color: #eef2f7; color: #a8b3c2; }
QPushButton#primary:disabled, QPushButton#danger:disabled, QPushButton#ghost:disabled { background: #eef2f7; border: 1px solid #eef2f7; color: #a8b3c2; }
QTextEdit#results { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 6px; color: #334155; font-size: 11px; }
QProgressBar { background: #e2e8f0; border: none; border-radius: 4px; max-height: 8px; min-height: 8px; text-align: center; color: transparent; }
QProgressBar::chunk { border-radius: 4px; background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #06b6d4); }
QCheckBox { color: #64748b; font-size: 11px; spacing: 6px; }
QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #cbd5e1; border-radius: 4px; background: #ffffff; }
QCheckBox::indicator:checked { background: #2563eb; border-color: #2563eb; }
QScrollBar:vertical { background: transparent; width: 8px; margin: 2px; }
QScrollBar::handle:vertical { background: #cbd5e1; border-radius: 4px; min-height: 30px; }
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

class ConfigWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VirMetrix — Configuración")
        self.resize(440, 620)
        self.setMinimumSize(380, 520)
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setStyleSheet(APP_QSS)

        self.orch = EvaluationOrchestrator()
        self.store = AttemptStore(Path.home() / ".virmetrix" / "history.sqlite")

        self._all_exercises = list_exercises()
        self.exercises = [e for e in self._all_exercises if e.get("module", "A") == "A"]
        if not self.exercises:
            raise RuntimeError("Catálogo vacío.")

        self.session = ExamSession(
            mode=ExamMode.PRACTICE,
            exercise_ids=[e["exercise_id"] for e in self.exercises],
        )

        self.work_copy: str | None = None
        self.practice_window = None
        self._doc_app = None
        self._doc_obj = None
        self._doc_proc = None
        self._started_at: str | None = None
        self._start_perf: float | None = None

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

        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 12, 12, 14)
        layout.setSpacing(10)

        layout.addWidget(self._build_mode_card())
        layout.addWidget(self._build_actions_card())
        layout.addWidget(self._build_session_card())
        layout.addLayout(self._build_footer())
        layout.addStretch(1)

        for card in content.findChildren(QFrame):
            if card.objectName() in ("card", "cardAccent"):
                _add_shadow(card)

        self._setup_tab_order()
        self._on_mode_changed(0)

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
        return header

    def _build_mode_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)
        col.addWidget(_eyebrow("Configuración"))

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

        row_mod = QHBoxLayout()
        row_mod.setSpacing(8)
        lbl_mod = QLabel("Evaluación")
        lbl_mod.setObjectName("counter")
        row_mod.addWidget(lbl_mod)
        self.module_combo = QComboBox()
        self.module_combo.addItem("Evaluación A (actual)", "A")
        self.module_combo.addItem("Evaluación B (pendiente)", "B")
        self.module_combo.addItem("Todo (A + B)", "all")
        self.module_combo.setAccessibleName("Módulo de evaluación")
        self.module_combo.currentIndexChanged.connect(self._on_module_changed)
        row_mod.addWidget(self.module_combo, 1)
        col.addLayout(row_mod)

        row_eval = QHBoxLayout()
        row_eval.setSpacing(8)
        lbl_eval = QLabel("Evaluador")
        lbl_eval.setObjectName("counter")
        row_eval.addWidget(lbl_eval)
        self.eval_mode_combo = QComboBox()
        pretty_modes = {"docx_portable": "Portable (sin Word)", "word_com": "Word COM (real)"}
        for mode, (avail, reason) in self.orch.available_modes().items():
            self.eval_mode_combo.addItem(pretty_modes.get(mode, mode), mode)
            if not avail:
                idx = self.eval_mode_combo.count() - 1
                self.eval_mode_combo.setItemData(idx, f"No disponible: {reason}", Qt.ToolTipRole)
        self.eval_mode_combo.setAccessibleName("Modo de evaluación")
        row_eval.addWidget(self.eval_mode_combo, 1)
        col.addLayout(row_eval)
        self.mode_combo = self.eval_mode_combo
        return card

    def _build_actions_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)
        col.addWidget(_eyebrow("Documento"))

        self.btn_copy = QPushButton("1.  Crear copia de trabajo")
        self.btn_copy.setAccessibleName("Crear copia de trabajo")
        self.btn_copy.clicked.connect(self.create_copy)
        col.addWidget(self.btn_copy)

        self.btn_open = QPushButton("2.  Abrir en Word")
        self.btn_open.setAccessibleName("Abrir documento")
        self.btn_open.clicked.connect(lambda: self.open_document())
        col.addWidget(self.btn_open)

        self.btn_start_practice = QPushButton("Iniciar práctica (abrir ventana)")
        self.btn_start_practice.setObjectName("primary")
        self.btn_start_practice.setAccessibleName("Iniciar práctica")
        self.btn_start_practice.setMinimumHeight(40)
        self.btn_start_practice.clicked.connect(self.open_practice)
        col.addWidget(self.btn_start_practice)
        return card

    def _build_session_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        col = QVBoxLayout(card)
        col.setContentsMargins(14, 12, 14, 12)
        col.setSpacing(8)
        col.addWidget(_eyebrow("Sesión"))

        row_session = QHBoxLayout()
        row_session.setSpacing(6)
        self.btn_start = QPushButton("Iniciar")
        self.btn_start.setObjectName("ghost")
        self.btn_start.clicked.connect(self.start_session)
        self.btn_pause = QPushButton("Pausar")
        self.btn_pause.setObjectName("ghost")
        self.btn_pause.clicked.connect(self.toggle_pause)
        self.btn_pause.setEnabled(False)
        self.btn_finish = QPushButton("Finalizar")
        self.btn_finish.setObjectName("ghost")
        self.btn_finish.clicked.connect(self.finish_session)
        self.btn_finish.setEnabled(False)
        row_session.addWidget(self.btn_start)
        row_session.addWidget(self.btn_pause)
        row_session.addWidget(self.btn_finish)
        col.addLayout(row_session)
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
        self.btn_about = QPushButton("Acerca de")
        self.btn_about.setObjectName("ghost")
        self.btn_about.setAccessibleName("Acerca de VirMetrix")
        self.btn_about.clicked.connect(self._show_about)
        row.addWidget(self.btn_about)
        return row

    def _setup_tab_order(self) -> None:
        order = [
            self.mode_use_combo,
            self.eval_mode_combo,
            self.btn_copy,
            self.btn_open,
            self.btn_start_practice,
            self.btn_start,
            self.btn_pause,
            self.btn_finish,
            self.stay_on_top,
            self.btn_about,
        ]
        for first, second in zip(order, order[1:]):
            self.setTabOrder(first, second)

    def _toggle_stay_on_top(self, state):
        if state:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()

    def _on_mode_changed(self, _index):
        self.session.mode = self.mode_use_combo.currentData()
        if self.session.mode == ExamMode.EXAM:
            self.session.duration_seconds = EXAM_DURATION_SECONDS
        else:
            self.session.duration_seconds = None

    def _filtered_exercises(self) -> list[dict]:
        mod = self.module_combo.currentData() if hasattr(self, "module_combo") else "A"
        if mod == "all":
            return [e for e in self._all_exercises]
        return [e for e in self._all_exercises if e.get("module", "A") == mod]

    def _on_module_changed(self, _index):
        self.exercises = self._filtered_exercises()
        if not self.exercises:
            QMessageBox.information(self, "Evaluación", "No hay ejercicios disponibles para este módulo.")
            return
        self.session = ExamSession(
            mode=self.mode_use_combo.currentData(),
            exercise_ids=[e["exercise_id"] for e in self.exercises],
            duration_seconds=(EXAM_DURATION_SECONDS if self.mode_use_combo.currentData() == ExamMode.EXAM else None),
        )

    def _reference_path(self) -> Path:
        return ROOT / "exercises" / self.session.current_exercise_id / "reference.docx"

    def create_copy(self):
        ref = self._reference_path()
        if not ref.is_file():
            QMessageBox.critical(self, "Error", "Falta el documento de referencia.")
            return
        work_dir = Path.home() / ".virmetrix" / "work"
        self.work_copy = self.orch.create_work_copy(str(ref), str(work_dir))
        if self.practice_window is not None:
            self.practice_window.result_view.setPlainText(f"Copia de trabajo creada:\n{self.work_copy}")
        QMessageBox.information(self, "Copia", f"Copia de trabajo creada:\n{self.work_copy}")

    def close_current_document(self):
        if self._doc_obj is not None:
            try:
                self._doc_obj.Close(SaveChanges=False)
            except Exception:
                pass
            self._doc_obj = None
        if self._doc_proc is not None:
            try:
                if self._doc_proc.poll() is None:
                    self._doc_proc.terminate()
                    self._doc_proc.wait(5)
            except Exception:
                try:
                    self._doc_proc.kill()
                except Exception:
                    pass
            self._doc_proc = None

    def open_document(self, path=None):
        if path is None:
            path = self.work_copy
        if not path:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        self.close_current_document()
        # Windows: usar Word COM para controlar la ventana
        if sys.platform == "win32":
            try:
                import pythoncom
                import win32com.client

                pythoncom.CoInitialize()
                if self._doc_app is None:
                    self._doc_app = win32com.client.Dispatch("Word.Application")
                    self._doc_app.Visible = True
                self._doc_obj = self._doc_app.Documents.Open(path)
                return
            except Exception:
                pass
        # Linux/macOS: lanzar instancia aislada de LibreOffice para poder cerrarla después
        if sys.platform != "win32":
            try:
                import subprocess
                import tempfile

                profile = tempfile.mkdtemp(prefix="lo_profile_", dir="/tmp/opencode")
                self._doc_proc = subprocess.Popen(
                    ["soffice", f"-env:UserInstallation=file://{profile}", path]
                )
                return
            except Exception:
                pass
        # Fallback genérico
        try:
            self.orch.open_document(path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir el documento: {e}")

    def open_practice(self):
        self.session.current_index = 0
        if self.practice_window is None:
            self.practice_window = PracticeWindow(self)
        self.practice_window.exercises = self.exercises
        self.practice_window._load_current_exercise()
        self.practice_window.show()
        self.practice_window.raise_()
        self.practice_window.activateWindow()

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
        self.session.start()
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_finish.setEnabled(True)
        if self.practice_window is not None:
            self.practice_window.on_session_started()

    def toggle_pause(self):
        if self.session.status == SessionStatus.RUNNING:
            self.session.pause()
            self.btn_pause.setText("Reanudar")
            if self.practice_window is not None:
                self.practice_window.on_session_paused()
        elif self.session.status == SessionStatus.PAUSED:
            self.session.resume()
            self.btn_pause.setText("Pausar")
            if self.practice_window is not None:
                self.practice_window.on_session_resumed()
        if self.practice_window is not None:
            self.practice_window.timer_label.setText(f"{self.session.format_remaining()}")

    def finish_session(self):
        if self.session.status == SessionStatus.FINISHED:
            return
        self.session.finish()
        self._finish_ui()
        if self.practice_window is not None:
            self.practice_window.timer_label.setText(f"{self.session.format_remaining()}")
            self.practice_window.result_view.append("Sesión finalizada.")

    def _finish_ui(self):
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_finish.setEnabled(False)
        self.btn_pause.setText("Pausar")

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

class PracticeWindow(QMainWindow):
    def __init__(self, config: ConfigWindow):
        super().__init__()
        self.config = config
        self.setWindowTitle("VirMetrix — Práctica")
        self.resize(1120, 260)
        self.setMinimumSize(900, 220)
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setStyleSheet(APP_QSS)

        self.orch = config.orch
        self.store = config.store
        self.exercises = config.exercises
        self.session = config.session

        self._thread: QThread | None = None
        self._worker: EvaluationWorker | None = None
        self._started_at: str | None = None
        self._start_perf: float | None = None
        self._current_step = 1

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_header())

        content = QWidget()
        content.setObjectName("content")
        root_layout.addWidget(content, 1)

        layout = QHBoxLayout(content)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)

        layout.addWidget(self._build_exercise_card(), 1)
        layout.addWidget(self._build_task_card(), 2)
        layout.addWidget(self._build_actions_card(), 1)
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card, 2)

        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._on_tick)

        self._load_current_exercise()
        self._set_step(1)

    def _build_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        row = QHBoxLayout(header)
        row.setContentsMargins(16, 8, 16, 8)
        row.setSpacing(12)

        brand = QVBoxLayout()
        brand.setSpacing(0)
        title = QLabel("VirMetrix")
        title.setObjectName("title")
        title.setAccessibleName("VirMetrix")
        subtitle = QLabel("PRÁCTICA · MOS WORD")
        subtitle.setObjectName("subtitle")
        brand.addWidget(title)
        brand.addWidget(subtitle)
        row.addLayout(brand)

        steps = QHBoxLayout()
        steps.setSpacing(6)
        self.chip_copy = _chip("1 · Copia")
        self.chip_edit = _chip("2 · Editar")
        self.chip_grade = _chip("3 · Evaluar")
        self.chip_copy.setAccessibleName("Paso 1, crear copia")
        self.chip_edit.setAccessibleName("Paso 2, editar en Word")
        self.chip_grade.setAccessibleName("Paso 3, evaluar")
        for chip in (self.chip_copy, self.chip_edit, self.chip_grade):
            steps.addWidget(chip)
        row.addLayout(steps, 1)

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
        self.exercise_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.exercise_combo.setMinimumContentsLength(12)
        for ex in self.exercises:
            self.exercise_combo.addItem(ex["title"], ex["exercise_id"])
        self.exercise_combo.setAccessibleName("Selector de ejercicio")
        self.exercise_combo.currentIndexChanged.connect(self._on_exercise_changed)
        col.addWidget(self.exercise_combo)

        nav = QHBoxLayout()
        nav.setSpacing(8)
        self.btn_prev = QPushButton("◀  Anterior")
        self.btn_prev.setObjectName("ghost")
        self.btn_prev.setAccessibleName("Ejercicio anterior")
        self.btn_prev.clicked.connect(self.prev_exercise)
        self.btn_next = QPushButton("Siguiente  ▶")
        self.btn_next.setObjectName("ghost")
        self.btn_next.setAccessibleName("Ejercicio siguiente")
        self.btn_next.clicked.connect(self.next_exercise)
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
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

        self.btn_eval = QPushButton("Evaluar ahora")
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
        self.score_caption.setAccessibleName("Detalle de puntuación")
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
        self.result_view.setMinimumHeight(52)
        self.result_view.setMaximumHeight(90)
        self.result_view.document().setDocumentMargin(6)
        self.result_view.setPlaceholderText("Los resultados de cada criterio aparecerán aquí tras evaluar.")
        col.addWidget(self.result_view)
        return card

    def _set_step(self, step: int) -> None:
        self._current_step = step
        states = {1: (("active", "idle", "idle")), 2: (("done", "active", "idle")), 3: (("done", "done", "active"))}.get(step, ("idle", "idle", "idle"))
        for chip, state in zip((self.chip_copy, self.chip_edit, self.chip_grade), states):
            chip.setProperty("state", state)
            chip.style().unpolish(chip)
            chip.style().polish(chip)

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
        self.step_label.setText(f"Ejercicio {self.session.current_index + 1} de {len(self.session.exercise_ids)}")
        self._reset_score()

    def _on_exercise_changed(self, _index):
        eid = self.exercise_combo.currentData()
        if eid and eid in self.session.exercise_ids:
            self.session.current_index = self.session.exercise_ids.index(eid)
        self.config.work_copy = None
        self._load_current_exercise()
        self._set_step(1)
        self.result_view.setPlainText(
            f"Ejercicio seleccionado: {eid}\n"
            "Crea una nueva copia de trabajo para este ejercicio."
        )

    def _auto_new_copy(self):
        ref = self.config._reference_path()
        if not ref.is_file():
            self.result_view.setPlainText("Falta el documento de referencia.")
            return
        work_dir = Path.home() / ".virmetrix" / "work"
        self.config.work_copy = self.config.orch.create_work_copy(str(ref), str(work_dir))
        try:
            self.config.open_document(self.config.work_copy)
        except Exception:
            pass
        self.result_view.setPlainText(
            f"Ejercicio cambiado. Nueva copia creada y abierta:\n{self.config.work_copy}"
        )

    def prev_exercise(self):
        if self.session.prev_exercise():
            self._load_current_exercise()
            self._set_step(1)
            self._auto_new_copy()

    def next_exercise(self):
        if self.session.next_exercise():
            self._load_current_exercise()
            self._set_step(1)
            self._auto_new_copy()

    def evaluate(self):
        wc = self.config.work_copy
        if not wc:
            QMessageBox.warning(self, "Aviso", "Primero crea una copia de trabajo.")
            return
        mode = self.config.eval_mode_combo.currentData()
        avail, reason = self.config.orch.available_modes()[mode]
        if not avail:
            QMessageBox.warning(self, "Modo no disponible", reason)
            return

        self.btn_eval.setEnabled(False)
        self.btn_cancel.setVisible(True)
        self.progress.setVisible(True)
        self.result_view.setPlainText("Evaluando… esto puede tardar unos segundos.")

        self._started_at = datetime.now(timezone.utc).isoformat()
        self._start_perf = time.perf_counter()

        self._thread = QThread()
        self._worker = EvaluationWorker(self.config.orch, mode, wc, self.exercise)
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
            return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        palette = {"pass": ("#059669", "✔"), "fail": ("#dc2626", "✖"), "not_verifiable": ("#64748b", "•"), "error": ("#b45309", "!")}
        parts = [f'<div style="color:#334155;font-weight:600;">Evaluador: {esc(result.evaluator)}</div>']
        if result.error_message:
            parts.append(f'<div style="color:#b45309;margin-top:4px;">Error: {esc(result.error_message)}</div>')
        for c in result.criteria:
            color, icon = palette.get(c.status.value, ("#64748b", "•"))
            parts.append(f'<div style="margin-top:4px;color:{color};">{icon} <b>{esc(c.criterion_id)}</b> · {esc(c.status.value)}<span style="color:#64748b;"> — {esc(c.evidence)}</span></div>')
        parts.append(f'<div style="margin-top:8px;color:#94a3b8;font-size:10px;">Historial guardado · intentos totales: {self.store.count()}</div>')
        return "".join(parts)

    def _on_finished(self, result):
        duration = None
        if self._start_perf is not None:
            duration = time.perf_counter() - self._start_perf
        try:
            self.store.save(result, version=self.exercise.version, started_at=self._started_at, finished_at=datetime.now(timezone.utc).isoformat(), duration_seconds=duration)
        except Exception as e:
            self.result_view.append(f"Aviso: no se pudo guardar historial: {e}")
        self.progress.setVisible(False)
        self.btn_cancel.setVisible(False)
        self.result_view.setHtml(self._render_result(result))
        self._show_score(result.score_earned, result.score_possible)
        self._set_step(3)
        self.btn_eval.setEnabled(True)

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

    def start_session(self):
        if self.session.status == SessionStatus.FINISHED:
            self.session = ExamSession(
                mode=self.config.mode_use_combo.currentData(),
                exercise_ids=[e["exercise_id"] for e in self.exercises],
                duration_seconds=(EXAM_DURATION_SECONDS if self.config.mode_use_combo.currentData() == ExamMode.EXAM else None),
            )
            self.config.session = self.session
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

    def on_session_started(self):
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_finish.setEnabled(True)
        if self.session.is_timed:
            self._timer.start()
        self.timer_label.setText(f"{self.session.format_remaining()}")

    def on_session_paused(self):
        self._timer.stop()
        self.btn_pause.setText("Reanudar")

    def on_session_resumed(self):
        if self.session.is_timed:
            self._timer.start()
        self.btn_pause.setText("Pausar")

# Backward compatibility alias
MainWindow = ConfigWindow

def run():
    app = QApplication(sys.argv)
    icon_path = ROOT / "virmetrix.ico"
    if icon_path.is_file():
        app.setWindowIcon(QIcon(str(icon_path)))
    w = ConfigWindow()
    if icon_path.is_file():
        w.setWindowIcon(QIcon(str(icon_path)))
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    run()
