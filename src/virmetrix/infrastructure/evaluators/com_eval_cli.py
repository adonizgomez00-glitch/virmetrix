"""CLI helper: evalua un documento con Word COM e imprime JSON por stdout."""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))

from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.infrastructure.catalog.loader import load_exercise


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: com_eval_cli.py <work_file> <exercise_id>", file=sys.stderr)
        return 2
    work_file = sys.argv[1]
    exercise_id = sys.argv[2]
    ex = load_exercise(exercise_id)
    orch = EvaluationOrchestrator()
    res = orch.evaluate("word_com", work_file, ex)
    payload = {
        "exercise_id": res.exercise_id,
        "evaluator": res.evaluator,
        "status": res.status,
        "score_earned": res.score_earned,
        "score_possible": res.score_possible,
        "rules_version": res.rules_version,
        "error_message": res.error_message,
        "criteria": [
            {"criterion_id": c.criterion_id, "status": c.status.value, "evidence": c.evidence}
            for c in res.criteria
        ],
    }
    print(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
