"""Evaluador portable de DOCX mediante inspección OOXML (sin Word)."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

from virmetrix.domain.ports import (
    CriterionResult,
    CriterionStatus,
    EvaluationResult,
    ExerciseMetadata,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB
MAX_ENTRIES = 1000
MAX_UNCOMPRESSED = 50 * 1024 * 1024  # 50 MB


def _parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(
        resolve_entities=False, no_network=True, recover=False, huge_tree=False
    )
    return etree.fromstring(data, parser=parser)


def _bool_attr(val: str | None) -> bool | None:
    if val is None:
        return None
    v = val.strip().lower()
    if v in ("", "1", "true", "on", "yes"):
        return True
    if v in ("0", "false", "off", "no"):
        return False
    return None


def _half_points_to_pt(val: str | None) -> float | None:
    if val is None:
        return None
    try:
        return int(val) / 2.0
    except ValueError:
        return None


class PortableDocxEvaluator:
    def is_available(self) -> tuple[bool, str]:
        return True, "portable"

    def evaluate(self, work_file_path: str, exercise: ExerciseMetadata) -> EvaluationResult:
        path = Path(work_file_path)
        target_phrase = getattr(exercise, "target_phrase", None)
        expected_text = getattr(exercise, "expected_text", None)

        # Load metadata sidecar if needed
        if target_phrase is None or expected_text is None:
            sidecar = self._load_sidecar(path, exercise)
            target_phrase = sidecar.get("target_phrase", target_phrase)
            expected_text = sidecar.get("expected_text", expected_text)

        # Security / validation
        if not path.is_file():
            return self._error(exercise, "Archivo no encontrado.")
        size = path.stat().st_size
        if size > MAX_FILE_SIZE:
            return self._error(exercise, "Archivo excede el tamaño máximo permitido.")
        if size == 0:
            return self._error(exercise, "Archivo vacío.")

        # Zip validation
        if not zipfile.is_zipfile(path):
            return self._error(exercise, "El archivo no es un ZIP/DOCX válido.")

        try:
            with zipfile.ZipFile(path, "r") as zf:
                infos = zf.infolist()
                if len(infos) > MAX_ENTRIES:
                    return self._error(exercise, "Demasiadas entradas en el ZIP.")
                total_uncompressed = sum(i.file_size for i in infos)
                if total_uncompressed > MAX_UNCOMPRESSED:
                    return self._error(exercise, "Tamaño descomprimido excesivo.")
                # Path traversal check (normalise separators for cross-platform)
                for info in infos:
                    name = info.filename
                    normalized = name.replace("\\", "/")
                    if "\x00" in normalized:
                        return self._error(exercise, "Ruta inválida en el paquete (traversal).")
                    parts = normalized.split("/")
                    if normalized.startswith("/") or ".." in parts or any(p == "" and i == 0 for i, p in enumerate(parts)):
                        return self._error(exercise, "Ruta inválida en el paquete (traversal).")
                # Read document.xml
                if "word/document.xml" not in zf.namelist():
                    return self._error(exercise, "Falta word/document.xml.")
                doc_xml = zf.read("word/document.xml")
                root = _parse_xml(doc_xml)
        except (zipfile.BadZipFile, etree.XMLSyntaxError, Exception) as e:
            return self._error(exercise, f"Error al leer el DOCX: {type(e).__name__}")

        # Extract runs with positions
        runs = self._extract_runs(root)
        full_text = "".join(r["text"] for r in runs)

        # Build run index: list of (run, start, end)
        idx = []
        pos = 0
        for r in runs:
            t = r["text"]
            idx.append((r, pos, pos + len(t)))
            pos += len(t)

        # C1: target appears exactly once in full text
        occurrences = []
        search_from = 0
        while True:
            i = full_text.find(target_phrase, search_from)
            if i == -1:
                break
            occurrences.append(i)
            search_from = i + 1
        c1_status = CriterionStatus.PASS if len(occurrences) == 1 else CriterionStatus.FAIL
        c1_evidence = f"Ocurrencias de '{target_phrase}': {len(occurrences)} (esperado 1)."

        # Determine target run ranges for the (first) occurrence if any
        target_runs: list[dict[str, Any]] = []
        if occurrences:
            start = occurrences[0]
            end = start + len(target_phrase)
            for r, s, e in idx:
                if s < end and e > start:
                    target_runs.append(r)

        # C2: bold effective on all target runs
        c2_status = self._check_prop(target_runs, "bold", True, can_resolve_inherit=False)
        # C3: size 16pt
        c3_status = self._check_size(target_runs, 16.0)

        # C4: neighbors not carrying target formatting
        c4_status = self._check_neighbors(idx, occurrences, target_phrase)

        # C5: text unchanged (full text == expected_text)
        if expected_text is None:
            c5_status = CriterionStatus.NOT_VERIFIABLE
            c5_evidence = "Texto esperado no definido; no se puede verificar C5."
        else:
            c5_status = CriterionStatus.PASS if full_text == expected_text else CriterionStatus.FAIL
            c5_evidence = "Texto del documento coincide con el esperado." if c5_status == CriterionStatus.PASS else "El texto del documento difiere del esperado."

        criteria = [
            CriterionResult("C1", c1_status, c1_evidence),
            CriterionResult("C2", c2_status, self._evidence_prop(target_runs, "bold")),
            CriterionResult("C3", c3_status, self._evidence_size(target_runs)),
            CriterionResult("C4", c4_status, "Texto vecino conserva formato de referencia."),
            CriterionResult("C5", c5_status, c5_evidence),
        ]

        # Scoring: sum pass
        earned = sum(1 for c in criteria if c.status == CriterionStatus.PASS)
        possible = len(criteria)

        return EvaluationResult(
            exercise_id=exercise.exercise_id,
            evaluator="docx_portable",
            status="completed",
            criteria=criteria,
            score_earned=earned,
            score_possible=possible,
            rules_version=exercise.rules_version,
        )

    def _load_sidecar(self, path: Path, exercise: ExerciseMetadata) -> dict:
        # Try to load metadata.json from exercises/<id>/
        try:
            # exercise_id may map to exercises/<id>/metadata.json
            base = Path(__file__).resolve().parents[4] / "exercises" / exercise.exercise_id
            meta_path = base / "metadata.json"
            if meta_path.is_file():
                import json

                return json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def _extract_runs(self, root: etree._Element) -> list[dict[str, Any]]:
        runs = []
        body = root.find(f"{{{W_NS}}}body")
        if body is None:
            body = root
        for p in body.iter(f"{{{W_NS}}}p"):
            for r in p.iter(f"{{{W_NS}}}r"):
                texts = [t.text or "" for t in r.iter(f"{{{W_NS}}}t")]
                text = "".join(texts)
                rpr = r.find(f"{{{W_NS}}}rPr")
                bold = None
                sz = None
                if rpr is not None:
                    b_el = rpr.find(f"{{{W_NS}}}b")
                    if b_el is not None:
                        bold = _bool_attr(b_el.get(f"{{{W_NS}}}val")) 
                        if bold is None:
                            bold = True
                    sz_el = rpr.find(f"{{{W_NS}}}sz")
                    if sz_el is not None:
                        sz = _half_points_to_pt(sz_el.get(f"{{{W_NS}}}val"))
                runs.append({"text": text, "bold": bold, "sz": sz, "rpr": rpr})
        # If no runs found via paragraphs, fallback: collect all w:t
        if not runs:
            for t in root.iter(f"{{{W_NS}}}t"):
                runs.append({"text": t.text or "", "bold": None, "sz": None, "rpr": None})
        return runs

    def _check_prop(self, target_runs, prop, expected, can_resolve_inherit=False):
        if not target_runs:
            return CriterionStatus.FAIL
        for r in target_runs:
            val = r.get(prop)
            if val is None:
                # cannot resolve inheritance reliably
                return CriterionStatus.NOT_VERIFIABLE
            if val != expected:
                return CriterionStatus.FAIL
        return CriterionStatus.PASS

    def _check_size(self, target_runs, expected_pt):
        if not target_runs:
            return CriterionStatus.FAIL
        for r in target_runs:
            sz = r.get("sz")
            if sz is None:
                return CriterionStatus.NOT_VERIFIABLE
            if abs(sz - expected_pt) > 1e-9:
                return CriterionStatus.FAIL
        return CriterionStatus.PASS

    def _evidence_prop(self, target_runs, prop):
        if not target_runs:
            return "No se encontró el texto objetivo."
        vals = [r.get(prop) for r in target_runs]
        return f"Propiedad '{prop}' en runs del objetivo: {vals}."

    def _evidence_size(self, target_runs):
        if not target_runs:
            return "No se encontró el texto objetivo."
        vals = [r.get("sz") for r in target_runs]
        return f"Tamaño (pt) en runs del objetivo: {vals}."

    def _check_neighbors(self, idx, occurrences, target_phrase):
        if not occurrences:
            return CriterionStatus.FAIL
        start = occurrences[0]
        end = start + len(target_phrase)
        # find run index containing start and end
        # neighbor before: last run before start
        before_run = None
        after_run = None
        for r, s, e in idx:
            if e <= start:
                before_run = r
            if s >= end and after_run is None:
                after_run = r
                break
        # Check that neighbors do not carry target formatting (bold True and sz 16)
        for nr in (before_run, after_run):
            if nr is None:
                continue
            if nr.get("bold") is True and nr.get("sz") == 16.0:
                return CriterionStatus.FAIL
        return CriterionStatus.PASS

    def _error(self, exercise, msg):
        criteria = [CriterionResult(cid, CriterionStatus.ERROR, msg) for cid in exercise.criteria]
        return EvaluationResult(
            exercise_id=exercise.exercise_id,
            evaluator="docx_portable",
            status="error",
            criteria=criteria,
            score_earned=0,
            score_possible=len(exercise.criteria),
            rules_version=exercise.rules_version,
            error_message=msg,
        )
