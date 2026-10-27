"""Orquestador de evaluación: selecciona adaptador y ejecuta."""

from __future__ import annotations

import hashlib
import shutil
from pathlib import Path

from virmetrix.domain.ports import EvaluationResult, ExerciseMetadata
from virmetrix.infrastructure.evaluators.portable import PortableDocxEvaluator
from virmetrix.infrastructure.evaluators.word_com import WordComAdapter


class EvaluationOrchestrator:
    def __init__(self):
        self.portable = PortableDocxEvaluator()
        self.word_com = WordComAdapter()

    def available_modes(self) -> dict[str, tuple[bool, str]]:
        return {
            "docx_portable": self.portable.is_available(),
            "word_com": self.word_com.is_available(),
        }

    def evaluate(self, mode: str, work_file_path: str, exercise: ExerciseMetadata) -> EvaluationResult:
        if mode == "docx_portable":
            return self.portable.evaluate(work_file_path, exercise)
        if mode == "word_com":
            return self.word_com.evaluate(work_file_path, exercise)
        raise ValueError(f"Modo desconocido: {mode}")

    @staticmethod
    def create_work_copy(reference_path: str, work_dir: str) -> str:
        ref = Path(reference_path)
        wd = Path(work_dir)
        wd.mkdir(parents=True, exist_ok=True)
        dest = wd / ref.name
        shutil.copy2(ref, dest)
        return str(dest)

    @staticmethod
    def sha256_file(path: str) -> str:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
