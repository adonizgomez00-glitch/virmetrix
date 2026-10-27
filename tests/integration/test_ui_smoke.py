import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from virmetrix.ui.main_window import MainWindow


def test_ui_smoke_instantiates():
    app = QApplication.instance() or QApplication([])
    w = MainWindow()
    assert w is not None
    # mode combo populated
    assert w.mode_combo.count() == 2
