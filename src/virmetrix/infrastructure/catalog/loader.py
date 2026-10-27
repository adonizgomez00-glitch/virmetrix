"""Carga de metadatos de ejercicio desde JSON."""

from __future__ import annotations

import json
from pathlib import Path

from virmetrix.domain.ports import ExerciseMetadata

ROOT = Path(__file__).resolve().parents[4]


def load_exercise(exercise_id: str) -> ExerciseMetadata:
    meta_path = ROOT / "exercises" / exercise_id / "metadata.json"
    data = json.loads(meta_path.read_text(encoding="utf-8"))
    return ExerciseMetadata(
        exercise_id=data["exercise_id"],
        version=data["version"],
        language=data["language"],
        competency=data["competency"],
        instructions=data["instructions"],
        criteria=data["criteria"],
        rules_version=data["rules_version"],
        target_phrase=data.get("target_phrase"),
        expected_text=data.get("expected_text"),
    )
