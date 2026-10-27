"""Adaptador Word COM (solo Windows). Importaciones bajo demanda."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from virmetrix.domain.ports import (
    CriterionResult,
    CriterionStatus,
    EvaluationResult,
    ExerciseMetadata,
)


class WordComAdapter:
    def is_available(self) -> tuple[bool, str]:
        if sys.platform != "win32":
            return False, "Word COM solo está disponible en Windows."
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except Exception as e:
            return False, f"pywin32 no disponible: {type(e).__name__}"
        # Try to connect to Word
        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            _ = word.Version
            return True, f"Word disponible (v{word.Version})."
        except Exception as e:
            return False, f"Microsoft Word no accesible: {type(e).__name__}: {e}"

    def evaluate(self, work_file_path: str, exercise: ExerciseMetadata) -> EvaluationResult:
        available, reason = self.is_available()
        if not available:
            criteria = [
                CriterionResult(cid, CriterionStatus.ERROR, reason) for cid in exercise.criteria
            ]
            return EvaluationResult(
                exercise_id=exercise.exercise_id,
                evaluator="word_com",
                status="error",
                criteria=criteria,
                score_earned=0,
                score_possible=len(exercise.criteria),
                rules_version=exercise.rules_version,
                error_message=reason,
            )

        target_phrase = getattr(exercise, "target_phrase", None)
        expected_text = getattr(exercise, "expected_text", None)
        if target_phrase is None or expected_text is None:
            sidecar = self._load_sidecar(exercise)
            target_phrase = target_phrase if target_phrase is not None else sidecar.get("target_phrase")
            expected_text = expected_text if expected_text is not None else sidecar.get("expected_text")

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(FileName=work_file_path, ReadOnly=True)

            content_text = doc.Content.Text or ""
            # Normalize trailing paragraph marks
            full_text = content_text.rstrip("\r\x0b")

            # Occurrences
            occurrences: list[int] = []
            search_from = 0
            if target_phrase:
                while True:
                    i = full_text.find(target_phrase, search_from)
                    if i == -1:
                        break
                    occurrences.append(i)
                    search_from = i + 1

            c1 = CriterionStatus.PASS if len(occurrences) == 1 else CriterionStatus.FAIL

            # Formatting of first occurrence
            bold_uniform: bool | None = None
            size_uniform: float | None = None
            mixed_bold = False
            mixed_size = False
            if occurrences and target_phrase:
                start = occurrences[0]
                end = start + len(target_phrase)
                try:
                    rng = doc.Range(start, end)
                    b = rng.Font.Bold
                    s = rng.Font.Size
                    bold_uniform = self._com_bool(b)
                    # detect mixed: win32com may return 999999
                    if b not in (-1, 0, True, False):
                        mixed_bold = True
                        bold_uniform = None
                    if s == 999999:
                        mixed_size = True
                        size_uniform = None
                    else:
                        try:
                            size_uniform = float(s)
                        except (TypeError, ValueError):
                            size_uniform = None
                except Exception:
                    bold_uniform = None
                    size_uniform = None

            if not occurrences or not target_phrase:
                c2 = CriterionStatus.FAIL
                c3 = CriterionStatus.FAIL
            else:
                if mixed_bold:
                    c2 = CriterionStatus.FAIL
                elif bold_uniform is None:
                    c2 = CriterionStatus.NOT_VERIFIABLE
                elif bold_uniform is True:
                    c2 = CriterionStatus.PASS
                else:
                    c2 = CriterionStatus.FAIL

                if mixed_size:
                    c3 = CriterionStatus.FAIL
                elif size_uniform is None:
                    c3 = CriterionStatus.NOT_VERIFIABLE
                elif abs(size_uniform - 16.0) < 1e-6:
                    c3 = CriterionStatus.PASS
                else:
                    c3 = CriterionStatus.FAIL

            # C4 neighbors
            if not occurrences or not target_phrase:
                c4 = CriterionStatus.FAIL
            else:
                start = occurrences[0]
                end = start + len(target_phrase)
                neighbor_leak = False
                try:
                    if start > 0:
                        prev_rng = doc.Range(start - 1, start)
                        pb = self._com_bool(prev_rng.Font.Bold)
                        ps = prev_rng.Font.Size
                        if pb is True and (isinstance(ps, (int, float)) and abs(float(ps) - 16.0) < 1e-6):
                            neighbor_leak = True
                    if end < len(full_text):
                        next_rng = doc.Range(end, end + 1)
                        nb = self._com_bool(next_rng.Font.Bold)
                        ns = next_rng.Font.Size
                        if nb is True and (isinstance(ns, (int, float)) and abs(float(ns) - 16.0) < 1e-6):
                            neighbor_leak = True
                except Exception:
                    pass
                c4 = CriterionStatus.FAIL if neighbor_leak else CriterionStatus.PASS

            # C5
            if expected_text is None:
                c5 = CriterionStatus.NOT_VERIFIABLE
                c5_ev = "Texto esperado no definido; no se puede verificar C5."
            else:
                c5 = CriterionStatus.PASS if full_text == expected_text else CriterionStatus.FAIL
                c5_ev = "Texto del documento coincide con el esperado." if c5 == CriterionStatus.PASS else "El texto del documento difiere del esperado."

            status_map = {"C1": c1, "C2": c2, "C3": c3, "C4": c4, "C5": c5}

            criteria = []
            for cid in exercise.criteria:
                st = status_map.get(cid, CriterionStatus.ERROR)
                if cid == "C1":
                    ev = f"Ocurrencias de '{target_phrase}': {len(occurrences)} (esperado 1)."
                elif cid == "C2":
                    ev = f"Negrita efectiva en objetivo: {bold_uniform}."
                elif cid == "C3":
                    ev = f"Tamaño (pt) en objetivo: {size_uniform}."
                elif cid == "C4":
                    ev = "Texto vecino conserva formato de referencia." if c4 == CriterionStatus.PASS else "Texto vecino presenta fuga de formato."
                elif cid == "C5":
                    ev = c5_ev
                else:
                    ev = "Evaluado por Word COM."
                criteria.append(CriterionResult(cid, st, ev))

            earned = sum(1 for c in criteria if c.status == CriterionStatus.PASS)
            return EvaluationResult(
                exercise_id=exercise.exercise_id,
                evaluator="word_com",
                status="completed",
                criteria=criteria,
                score_earned=earned,
                score_possible=len(criteria),
                rules_version=exercise.rules_version,
            )
        except Exception as e:
            criteria = [
                CriterionResult(cid, CriterionStatus.ERROR, f"Error COM: {type(e).__name__}: {e}")
                for cid in exercise.criteria
            ]
            return EvaluationResult(
                exercise_id=exercise.exercise_id,
                evaluator="word_com",
                status="error",
                criteria=criteria,
                score_earned=0,
                score_possible=len(exercise.criteria),
                rules_version=exercise.rules_version,
                error_message=f"{type(e).__name__}: {e}",
            )
        finally:
            try:
                if doc is not None:
                    doc.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def _com_bool(val: Any) -> bool | None:
        if val is True or val == -1:
            return True
        if val is False or val == 0:
            return False
        return None

    @staticmethod
    def _load_sidecar(exercise: ExerciseMetadata) -> dict:
        try:
            base = Path(__file__).resolve().parents[4] / "exercises" / exercise.exercise_id
            meta_path = base / "metadata.json"
            if meta_path.is_file():
                return json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}