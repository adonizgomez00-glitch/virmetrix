import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication

from virmetrix.domain.ports import CriterionResult, CriterionStatus, EvaluationResult
from virmetrix.ui.main_window import MainWindow


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


@pytest.fixture(scope="module")
def window(app):
    return MainWindow()


def test_steps_track_progress(window):
    window._set_step(1)
    assert [c.property("state") for c in (window.chip_copy, window.chip_edit, window.chip_grade)] == [
        "active",
        "idle",
        "idle",
    ]
    window._set_step(2)
    assert [c.property("state") for c in (window.chip_copy, window.chip_edit, window.chip_grade)] == [
        "done",
        "active",
        "idle",
    ]
    window._set_step(3)
    assert [c.property("state") for c in (window.chip_copy, window.chip_edit, window.chip_grade)] == [
        "done",
        "done",
        "active",
    ]


def test_score_colors_and_reset(window):
    window._show_score(5, 5)
    assert window.score_label.text() == "5/5"
    assert "#059669" in window.score_label.styleSheet()
    assert window.score_label.isVisibleTo(window)

    window._show_score(1, 5)
    assert "#dc2626" in window.score_label.styleSheet()

    window._reset_score()
    assert not window.score_label.isVisibleTo(window)
    assert window.score_caption.text() == "sin evaluar"


def test_render_result_contains_statuses_and_escapes(window):
    result = EvaluationResult(
        exercise_id="word-format-001",
        evaluator="docx_portable",
        status="completed",
        criteria=[
            CriterionResult("C1", CriterionStatus.PASS, "ok"),
            CriterionResult("C2", CriterionStatus.FAIL, "esperaba <16 pt> & negrita"),
        ],
        score_earned=1,
        score_possible=2,
    )
    html = window._render_result(result)
    assert "C1" in html and "C2" in html
    assert "#059669" in html and "#dc2626" in html
    assert "&lt;16 pt&gt; &amp; negrita" in html
