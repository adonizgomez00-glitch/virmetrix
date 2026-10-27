import tempfile
from pathlib import Path

from virmetrix.domain.ports import CriterionResult, CriterionStatus, EvaluationResult
from virmetrix.infrastructure.persistence.sqlite import AttemptStore


def test_save_and_list_attempt(tmp_path):
    db = tmp_path / "h.sqlite"
    store = AttemptStore(db)
    res = EvaluationResult(
        exercise_id="word-format-001",
        evaluator="docx_portable",
        status="completed",
        criteria=[CriterionResult("C1", CriterionStatus.PASS, "ok")],
        score_earned=1,
        score_possible=1,
        rules_version="1.0.0",
    )
    aid = store.save(res, version="1.0.0", duration_seconds=1.25)
    assert aid >= 1
    attempts = store.list_attempts("word-format-001")
    assert len(attempts) == 1
    a = attempts[0]
    assert a.score_earned == 1
    assert a.criteria[0]["status"] == "pass"
    assert store.count() == 1


def test_list_all_and_limit(tmp_path):
    db = tmp_path / "h.sqlite"
    store = AttemptStore(db)
    for i in range(3):
        res = EvaluationResult(
            exercise_id="word-format-001",
            evaluator="docx_portable",
            status="completed",
            criteria=[],
            score_earned=0,
            score_possible=0,
        )
        store.save(res)
    assert len(store.list_attempts(limit=2)) == 2
    assert len(store.list_attempts()) == 3