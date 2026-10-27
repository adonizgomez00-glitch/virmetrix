import sys

from virmetrix.infrastructure.evaluators.word_com import WordComAdapter
from virmetrix.infrastructure.catalog.loader import load_exercise


def test_is_available_linux():
    a = WordComAdapter()
    avail, reason = a.is_available()
    if sys.platform != "win32":
        assert avail is False
        assert reason


def test_evaluate_returns_error_when_unavailable(tmp_path):
    a = WordComAdapter()
    ex = load_exercise("word-format-001")
    p = tmp_path / "x.docx"
    p.write_bytes(b"x")
    res = a.evaluate(str(p), ex)
    # On Linux, should be error status
    if sys.platform != "win32":
        assert res.status == "error"
        assert res.error_message
