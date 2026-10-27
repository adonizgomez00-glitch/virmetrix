import hashlib
from pathlib import Path

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.infrastructure.catalog.loader import load_exercise


def test_available_modes(tmp_path):
    o = EvaluationOrchestrator()
    modes = o.available_modes()
    assert "docx_portable" in modes
    assert "word_com" in modes
    assert modes["docx_portable"][0] is True


def test_create_work_copy_and_hash(tmp_path):
    ref = Path("exercises/word-format-001/reference.docx")
    assert ref.is_file()
    o = EvaluationOrchestrator()
    work = o.create_work_copy(str(ref), str(tmp_path / "work"))
    assert Path(work).is_file()
    assert Path(work).read_bytes() == ref.read_bytes()
    h = o.sha256_file(work)
    assert h == hashlib.sha256(ref.read_bytes()).hexdigest()


def test_evaluate_portable(tmp_path):
    ref = Path("exercises/word-format-001/reference.docx")
    o = EvaluationOrchestrator()
    work = o.create_work_copy(str(ref), str(tmp_path / "work"))
    ex = load_exercise("word-format-001")
    # reference is unformatted -> C2/C3 likely not_verifiable/fail, but runs
    res = o.evaluate("docx_portable", work, ex)
    assert res.evaluator == "docx_portable"
    assert res.score_possible == 5
