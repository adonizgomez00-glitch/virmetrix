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
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
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

        footnotes_text = ""
        endnotes_text = ""
        has_toc = False
        footnote_ref_count = 0
        drawing_count = 0
        alt_texts: list[str] = []
        textbox_texts: list[str] = []
        header_texts: list[str] = []
        footer_texts: list[str] = []
        core_props: dict = {}
        page_orient = None
        comments_text = ""
        tracked_change_count = 0

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
                footnote_ref_count = doc_xml.decode("utf-8", errors="ignore").count("footnoteReference")
                if "word/footnotes.xml" in zf.namelist():
                    try:
                        footnotes_text = zf.read("word/footnotes.xml").decode("utf-8", errors="ignore")
                    except Exception:
                        pass
                if "word/endnotes.xml" in zf.namelist():
                    try:
                        endnotes_text = zf.read("word/endnotes.xml").decode("utf-8", errors="ignore")
                    except Exception:
                        pass
                try:
                    has_toc = "TOC" in doc_xml.decode("utf-8", errors="ignore")
                except Exception:
                    has_toc = False
                try:
                    drawing_count = len(root.findall(f".//{{{W_NS}}}drawing"))
                    for dp in root.iter(f"{{{WP_NS}}}docPr"):
                        alt_texts.append(dp.get("descr") or "")
                    for tb in root.iter(f"{{{W_NS}}}txbxContent"):
                        textbox_texts.append("".join(t.text or "" for t in tb.iter(f"{{{W_NS}}}t")))
                except Exception:
                    pass
                # headers/footers
                for name in zf.namelist():
                    if name.startswith("word/header") and name.endswith(".xml"):
                        try:
                            header_texts.append(zf.read(name).decode("utf-8", errors="ignore"))
                        except Exception:
                            pass
                    if name.startswith("word/footer") and name.endswith(".xml"):
                        try:
                            footer_texts.append(zf.read(name).decode("utf-8", errors="ignore"))
                        except Exception:
                            pass
                # comments
                if "word/comments.xml" in zf.namelist():
                    try:
                        comments_text = zf.read("word/comments.xml").decode("utf-8", errors="ignore")
                    except Exception:
                        pass
                # tracked changes count
                try:
                    tracked_change_count = doc_xml.decode("utf-8", errors="ignore").count("<w:ins ") + doc_xml.decode("utf-8", errors="ignore").count("<w:del ")
                except Exception:
                    tracked_change_count = 0

                # core properties
                if "docProps/core.xml" in zf.namelist():
                    try:
                        core_xml = zf.read("docProps/core.xml").decode("utf-8", errors="ignore")
                        core_root = etree.fromstring(core_xml.encode("utf-8"))
                        for el in core_root.iter():
                            tag = el.tag.split("}")[-1]
                            if el.text and el.text.strip():
                                core_props[tag] = el.text.strip()
                    except Exception:
                        pass
                # page orientation
                try:
                    sectPr = root.find(f".//{{{W_NS}}}sectPr")
                    if sectPr is not None:
                        pgSz = sectPr.find(f"{{{W_NS}}}pgSz")
                        if pgSz is not None:
                            orient = pgSz.get(f"{{{W_NS}}}orient")
                            if orient:
                                page_orient = orient
                            else:
                                w = pgSz.get(f"{{{W_NS}}}w")
                                h = pgSz.get(f"{{{W_NS}}}h")
                                if w and h:
                                    page_orient = "landscape" if int(w) > int(h) else "portrait"
                except Exception:
                    pass
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

        tables = self._extract_tables(root)
        list_items = self._count_list_items(root)
        specs = getattr(exercise, "criteria_specs", None) or []
        if specs:
            criteria = self._evaluate_specs(specs, full_text, runs, idx, target_phrase, expected_text, tables, list_items, footnotes_text, endnotes_text, has_toc, footnote_ref_count, drawing_count, alt_texts, textbox_texts, header_texts, footer_texts, core_props, page_orient, comments_text, tracked_change_count)
        else:
            # Legacy hardcoded criteria (backward compatibility)
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

            target_runs: list[dict[str, Any]] = []
            if occurrences:
                start = occurrences[0]
                end = start + len(target_phrase)
                for r, s, e in idx:
                    if s < end and e > start:
                        target_runs.append(r)

            c2_status = self._check_prop(target_runs, "bold", True, can_resolve_inherit=False)
            c3_status = self._check_size(target_runs, 16.0)
            c4_status = self._check_neighbors(idx, occurrences, target_phrase)

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
            from virmetrix.paths import resource_root

            base = resource_root() / "exercises" / exercise.exercise_id
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
                italic = None
                sz = None
                font = None
                color = None
                if rpr is not None:
                    b_el = rpr.find(f"{{{W_NS}}}b")
                    if b_el is not None:
                        bold = _bool_attr(b_el.get(f"{{{W_NS}}}val"))
                        if bold is None:
                            bold = True
                    i_el = rpr.find(f"{{{W_NS}}}i")
                    if i_el is not None:
                        italic = _bool_attr(i_el.get(f"{{{W_NS}}}val"))
                        if italic is None:
                            italic = True
                    sz_el = rpr.find(f"{{{W_NS}}}sz")
                    if sz_el is not None:
                        sz = _half_points_to_pt(sz_el.get(f"{{{W_NS}}}val"))
                    rfonts = rpr.find(f"{{{W_NS}}}rFonts")
                    if rfonts is not None:
                        font = rfonts.get(f"{{{W_NS}}}ascii") or rfonts.get(f"{{{W_NS}}}hAnsi") or rfonts.get(f"{{{W_NS}}}cs")
                    color_el = rpr.find(f"{{{W_NS}}}color")
                    if color_el is not None:
                        color = color_el.get(f"{{{W_NS}}}val")
                runs.append({"text": text, "bold": bold, "italic": italic, "sz": sz, "font": font, "color": color, "rpr": rpr})
        # If no runs found via paragraphs, fallback: collect all w:t
        if not runs:
            for t in root.iter(f"{{{W_NS}}}t"):
                runs.append({"text": t.text or "", "bold": None, "italic": None, "sz": None, "font": None, "color": None, "rpr": None})
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

    _PROP_KEY = {"bold": "bold", "italic": "italic", "size_pt": "sz", "font_name": "font", "color": "color"}

    def _run_prop(self, run: dict, prop: str):
        key = self._PROP_KEY.get(prop, prop)
        return run.get(key)

    @staticmethod
    def _occurrences(text: str, target: str) -> list[int]:
        res = []
        search_from = 0
        while True:
            i = text.find(target, search_from)
            if i == -1:
                break
            res.append(i)
            search_from = i + 1
        return res

    def _evaluate_specs(self, specs, full_text, runs, idx, target_phrase, expected_text, tables=None, list_items=0, footnotes_text="", endnotes_text="", has_toc=False, footnote_ref_count=0, drawing_count=0, alt_texts=None, textbox_texts=None, header_texts=None, footer_texts=None, core_props=None, page_orient=None, comments_text="", tracked_change_count=0):
        tables = tables or []
        alt_texts = alt_texts or []
        textbox_texts = textbox_texts or []
        header_texts = header_texts or []
        footer_texts = footer_texts or []
        core_props = core_props or {}
        criteria = []
        for spec in specs:
            cid = spec["id"]
            typ = spec.get("type")
            target = spec.get("target", target_phrase)
            if typ == "target_occurrences":
                expected = spec.get("expected", 1)
                occs = self._occurrences(full_text, target) if target else []
                status = CriterionStatus.PASS if len(occs) == expected else CriterionStatus.FAIL
                ev = f"Ocurrencias de '{target}': {len(occs)} (esperado {expected})."
            elif typ == "run_property":
                prop = spec["property"]
                expected = spec["expected"]
                occs = self._occurrences(full_text, target) if target else []
                target_runs = []
                if occs:
                    start = occs[0]
                    end = start + len(target)
                    for r, s, e in idx:
                        if s < end and e > start:
                            target_runs.append(r)
                if not target_runs:
                    status = CriterionStatus.FAIL
                    ev = f"No se encontró el texto objetivo '{target}'."
                else:
                    vals = [self._run_prop(r, prop) for r in target_runs]
                    if any(v is None for v in vals):
                        status = CriterionStatus.NOT_VERIFIABLE
                    elif all(v == expected for v in vals):
                        status = CriterionStatus.PASS
                    else:
                        status = CriterionStatus.FAIL
                    ev = f"Propiedad '{prop}' en objetivo: {vals}."
            elif typ == "neighbor_not_formatted":
                forbidden = spec.get("forbidden", {})
                occs = self._occurrences(full_text, target) if target else []
                if not occs:
                    status = CriterionStatus.FAIL
                    ev = f"No se encontró el texto objetivo '{target}'."
                else:
                    start = occs[0]
                    end = start + len(target)
                    before_run = None
                    after_run = None
                    for r, s, e in idx:
                        if e <= start:
                            before_run = r
                        if s >= end and after_run is None:
                            after_run = r
                            break
                    leak = False
                    for nr in (before_run, after_run):
                        if nr is None:
                            continue
                        if all(self._run_prop(nr, p) == v for p, v in forbidden.items()):
                            leak = True
                            break
                    status = CriterionStatus.FAIL if leak else CriterionStatus.PASS
                    ev = "Texto vecino conserva formato de referencia." if status == CriterionStatus.PASS else "Texto vecino presenta fuga de formato."
            elif typ == "document_text_equals":
                expected = spec.get("expected", expected_text)
                if expected is None:
                    status = CriterionStatus.NOT_VERIFIABLE
                    ev = "Texto esperado no definido; no se puede verificar."
                else:
                    norm_full = "".join(full_text.split())
                    norm_expected = "".join(str(expected).split())
                    status = CriterionStatus.PASS if norm_full == norm_expected else CriterionStatus.FAIL
                    ev = "Texto del documento coincide con el esperado." if status == CriterionStatus.PASS else "El texto del documento difiere del esperado."
            elif typ == "table_count":
                expected = spec.get("expected", 1)
                status = CriterionStatus.PASS if len(tables) == expected else CriterionStatus.FAIL
                ev = f"Tablas en el documento: {len(tables)} (esperado {expected})."
            elif typ == "table_dimensions":
                tindex = spec.get("table_index", 0)
                exp_rows = spec.get("rows")
                exp_cols = spec.get("cols")
                if tindex >= len(tables):
                    status = CriterionStatus.FAIL
                    ev = f"No existe la tabla {tindex}."
                else:
                    tbl = tables[tindex]
                    rows = len(tbl)
                    cols = len(tbl[0]) if tbl else 0
                    ok = True
                    if exp_rows is not None and rows != exp_rows:
                        ok = False
                    if exp_cols is not None and cols != exp_cols:
                        ok = False
                    status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                    ev = f"Tabla {tindex}: {rows}x{cols} (esperado {exp_rows}x{exp_cols})."
            elif typ == "table_cell_text":
                tindex = spec.get("table_index", 0)
                r = spec.get("row", 0)
                c = spec.get("col", 0)
                expected = spec.get("expected", "")
                if tindex >= len(tables):
                    status = CriterionStatus.FAIL
                    ev = f"No existe la tabla {tindex}."
                else:
                    tbl = tables[tindex]
                    if r >= len(tbl) or c >= len(tbl[r]):
                        status = CriterionStatus.FAIL
                        ev = f"Celda ({r},{c}) fuera de rango."
                    else:
                        actual = tbl[r][c]
                        status = CriterionStatus.PASS if actual == expected else CriterionStatus.FAIL
                        ev = f"Celda ({r},{c}): '{actual}' (esperado '{expected}')."
            elif typ == "list_items_count":
                expected = spec.get("expected", 1)
                status = CriterionStatus.PASS if list_items == expected else CriterionStatus.FAIL
                ev = f"Elementos de lista: {list_items} (esperado {expected})."
            elif typ == "footnote_count":
                expected = spec.get("expected", 1)
                count = footnote_ref_count
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Referencias de nota al pie: {count} (esperado >= {expected})."
            elif typ == "footnote_contains":
                expected = spec.get("expected", "")
                status = CriterionStatus.PASS if expected and expected in (footnotes_text or "") else CriterionStatus.FAIL
                ev = "La nota al pie contiene el texto esperado." if status == CriterionStatus.PASS else "La nota al pie no contiene el texto esperado."
            elif typ == "endnote_contains":
                expected = spec.get("expected", "")
                status = CriterionStatus.PASS if expected and expected in (endnotes_text or "") else CriterionStatus.FAIL
                ev = "La nota final contiene el texto esperado." if status == CriterionStatus.PASS else "La nota final no contiene el texto esperado."
            elif typ == "toc_present":
                status = CriterionStatus.PASS if has_toc else CriterionStatus.FAIL
                ev = "Se detectó tabla de contenido." if status == CriterionStatus.PASS else "No se detectó tabla de contenido."
            elif typ == "drawing_count":
                expected = spec.get("expected", 1)
                status = CriterionStatus.PASS if drawing_count == expected else CriterionStatus.FAIL
                ev = f"Imágenes/dibujos: {drawing_count} (esperado {expected})."
            elif typ == "alt_text_contains":
                expected = spec.get("expected", "")
                ok = any(expected in (a or "") for a in alt_texts)
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Texto alternativo correcto." if status == CriterionStatus.PASS else "Texto alternativo ausente o incorrecto."
            elif typ == "textbox_contains":
                expected = spec.get("expected", "")
                ok = any(expected in (t or "") for t in textbox_texts)
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Cuadro de texto contiene el texto esperado." if status == CriterionStatus.PASS else "Cuadro de texto ausente o incorrecto."
            elif typ == "header_contains":
                expected = spec.get("expected", "")
                ok = any(expected in (h or "") for h in header_texts)
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Encabezado contiene el texto esperado." if status == CriterionStatus.PASS else "Encabezado ausente o incorrecto."
            elif typ == "footer_contains":
                expected = spec.get("expected", "")
                ok = any(expected in (f or "") for f in footer_texts)
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Pie de página contiene el texto esperado." if status == CriterionStatus.PASS else "Pie de página ausente o incorrecto."
            elif typ == "doc_property":
                prop = spec.get("property", "")
                expected = spec.get("expected", "")
                actual = None
                for k, v in core_props.items():
                    if k.lower() == str(prop).lower():
                        actual = v
                        break
                status = CriterionStatus.PASS if actual == expected else CriterionStatus.FAIL
                ev = f"Propiedad '{prop}': '{actual}' (esperado '{expected}')."
            elif typ == "page_orientation":
                expected = spec.get("expected", "portrait")
                status = CriterionStatus.PASS if (page_orient or "").lower().startswith(expected.lower()) else CriterionStatus.FAIL
                ev = f"Orientación: {page_orient} (esperado {expected})."
            elif typ == "comment_contains":
                expected = spec.get("expected", "")
                status = CriterionStatus.PASS if expected and expected in (comments_text or "") else CriterionStatus.FAIL
                ev = "Comentario contiene el texto esperado." if status == CriterionStatus.PASS else "Comentario ausente o incorrecto."
            elif typ == "comment_count":
                expected = spec.get("expected", 1)
                count = (comments_text or "").count("<w:comment ")
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Comentarios: {count} (esperado >= {expected})."
            elif typ == "tracked_changes_present":
                expected = spec.get("expected", True)
                present = tracked_change_count > 0
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Se detectaron cambios registrados." if present else "No se detectaron cambios registrados."
            else:
                status = CriterionStatus.ERROR
                ev = f"Tipo de criterio desconocido: {typ}."
            criteria.append(CriterionResult(cid, status, ev))
        return criteria

    def _extract_tables(self, root: etree._Element) -> list[list[list[str]]]:
        tables = []
        body = root.find(f"{{{W_NS}}}body")
        if body is None:
            body = root
        for tbl in body.iter(f"{{{W_NS}}}tbl"):
            rows = []
            for tr in tbl.findall(f"{{{W_NS}}}tr"):
                cells = []
                for tc in tr.findall(f"{{{W_NS}}}tc"):
                    texts = [t.text or "" for t in tc.iter(f"{{{W_NS}}}t")]
                    cells.append("".join(texts))
                rows.append(cells)
            tables.append(rows)
        return tables

    @staticmethod
    def _count_list_items(root: etree._Element) -> int:
        count = 0
        body = root.find(f"{{{W_NS}}}body")
        if body is None:
            body = root
        for p in body.iter(f"{{{W_NS}}}p"):
            ppr = p.find(f"{{{W_NS}}}pPr")
            is_list = False
            if ppr is not None:
                if ppr.find(f"{{{W_NS}}}numPr") is not None:
                    is_list = True
                pstyle = ppr.find(f"{{{W_NS}}}pStyle")
                if pstyle is not None:
                    val = (pstyle.get(f"{{{W_NS}}}val") or "").lower()
                    if "list" in val:
                        is_list = True
            if is_list:
                count += 1
        return count

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
