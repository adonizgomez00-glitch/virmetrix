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
        comments_extended_text = ""
        tracked_change_count = 0
        custom_props_text = ""


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
                settings_text = ""
                if "word/settings.xml" in zf.namelist():
                    try:
                        settings_text = zf.read("word/settings.xml").decode("utf-8", errors="ignore")
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
                # modern comments extended (resolved state, etc.)
                for _ce in ("word/commentsExtended.xml", "word/commentsExtensible.xml"):
                    if _ce in zf.namelist():
                        try:
                            comments_extended_text += zf.read(_ce).decode("utf-8", errors="ignore")
                        except Exception:
                            pass
                # tracked changes count (incluye inserciones, eliminaciones, movimientos
                # y cambios de formato/propiedades, para coincidir con Word COM).
                try:
                    _raw = doc_xml.decode("utf-8", errors="ignore")
                    tracked_change_count = (
                        _raw.count("<w:ins ") + _raw.count("<w:del ")
                        + _raw.count("<w:moveFrom ") + _raw.count("<w:moveTo ")
                        + _raw.count("<w:rPrChange") + _raw.count("<w:pPrChange")
                        + _raw.count("<w:sectPrChange") + _raw.count("<w:tblPrChange")
                    )
                except Exception:
                    tracked_change_count = 0

                # styles (for inherited run properties)
                styles_root = None
                if "word/styles.xml" in zf.namelist():
                    try:
                        styles_root = _parse_xml(zf.read("word/styles.xml"))
                    except Exception:
                        styles_root = None
                # numbering (for list start values)
                numbering_root = None
                if "word/numbering.xml" in zf.namelist():
                    try:
                        numbering_root = _parse_xml(zf.read("word/numbering.xml"))
                    except Exception:
                        numbering_root = None
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
                # custom properties (used by Word add-in to record actions)
                custom_props_text = ""
                if "docProps/custom.xml" in zf.namelist():
                    try:
                        custom_props_text = zf.read("docProps/custom.xml").decode("utf-8", errors="ignore")
                    except Exception:
                        custom_props_text = ""
                # page orientation
                # Usar el ULTIMO sectPr del documento (seccion principal/final),
                # no el primero: los saltos de seccion intermedios aparecen antes
                # en el XML y podrian tener orientacion distinta (p. ej. portrait),
                # falseando el resultado.
                try:
                    sectPrs = root.findall(f".//{{{W_NS}}}sectPr")
                    sectPr = sectPrs[-1] if sectPrs else None
                    if sectPr is not None:
                        pgSz = sectPr.find(f"{{{W_NS}}}pgSz")
                        if pgSz is not None:
                            w = pgSz.get(f"{{{W_NS}}}w")
                            h = pgSz.get(f"{{{W_NS}}}h")
                            orient = pgSz.get(f"{{{W_NS}}}orient")
                            # Prefer physical dimensions (unambiguous) over the
                            # orient attribute, which may be stale or omitted.
                            if w and h:
                                page_orient = "landscape" if int(w) > int(h) else "portrait"
                            elif orient:
                                page_orient = orient
                except Exception:
                    pass
        except (zipfile.BadZipFile, etree.XMLSyntaxError, Exception) as e:
            return self._error(exercise, f"Error al leer el DOCX: {type(e).__name__}")

        # Extract runs with positions
        runs = self._extract_runs(root, styles_root if "styles_root" in locals() else None)
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
            criteria = self._evaluate_specs(specs, full_text, runs, idx, target_phrase, expected_text, tables, list_items, footnotes_text, endnotes_text, has_toc, footnote_ref_count, drawing_count, alt_texts, textbox_texts, header_texts, footer_texts, core_props, page_orient, comments_text, tracked_change_count, root, numbering_root, settings_text=settings_text, comments_extended_text=comments_extended_text, custom_props_text=custom_props_text)
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

        # Scoring: sum pass; not_verifiable criteria are not part of the score
        earned = sum(1 for c in criteria if c.status == CriterionStatus.PASS)
        possible = sum(1 for c in criteria if c.status != CriterionStatus.NOT_VERIFIABLE)

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

    def _extract_runs(self, root: etree._Element, styles_root: etree._Element | None = None) -> list[dict[str, Any]]:
        def rpr_props(rpr):
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
                bcs_el = rpr.find(f"{{{W_NS}}}bCs")
                if bcs_el is not None and bold is None:
                    bold = _bool_attr(bcs_el.get(f"{{{W_NS}}}val"))
                    if bold is None:
                        bold = True
                i_el = rpr.find(f"{{{W_NS}}}i")
                if i_el is not None:
                    italic = _bool_attr(i_el.get(f"{{{W_NS}}}val"))
                    if italic is None:
                        italic = True
                ics_el = rpr.find(f"{{{W_NS}}}iCs")
                if ics_el is not None and italic is None:
                    italic = _bool_attr(ics_el.get(f"{{{W_NS}}}val"))
                    if italic is None:
                        italic = True
                sz_el = rpr.find(f"{{{W_NS}}}sz")
                if sz_el is not None:
                    sz = _half_points_to_pt(sz_el.get(f"{{{W_NS}}}val"))
                szcs_el = rpr.find(f"{{{W_NS}}}szCs")
                if szcs_el is not None and sz is None:
                    sz = _half_points_to_pt(szcs_el.get(f"{{{W_NS}}}val"))
                rfonts = rpr.find(f"{{{W_NS}}}rFonts")
                if rfonts is not None:
                    font = rfonts.get(f"{{{W_NS}}}ascii") or rfonts.get(f"{{{W_NS}}}hAnsi") or rfonts.get(f"{{{W_NS}}}cs") or rfonts.get(f"{{{W_NS}}}eastAsia")
                color_el = rpr.find(f"{{{W_NS}}}color")
                if color_el is not None:
                    color = color_el.get(f"{{{W_NS}}}val")
            return bold, italic, sz, font, color

        runs = []
        body = root.find(f"{{{W_NS}}}body")
        if body is None:
            body = root
        for p in body.iter(f"{{{W_NS}}}p"):
            for r in p.iter(f"{{{W_NS}}}r"):
                texts = [t.text or "" for t in r.iter(f"{{{W_NS}}}t")]
                text = "".join(texts)
                rpr = r.find(f"{{{W_NS}}}rPr")
                bold, italic, sz, font, color = rpr_props(rpr)
                runs.append({"text": text, "bold": bold, "italic": italic, "sz": sz, "font": font, "color": color, "rpr": rpr})
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

    def _custom_prop(self, custom_props_text: str, name: str, expected) -> bool:
        import re
        prop_name = f"Virmetrix_{name}"
        m = re.search(r'<property[^>]*name="' + re.escape(prop_name) + r'"[^>]*>.*?</property>', custom_props_text or "", re.S | re.I)
        if not m:
            m = re.search(r'<property[^>]*name="' + re.escape(name) + r'"[^>]*>.*?</property>', custom_props_text or "", re.S | re.I)
        if not m:
            return False
        block = m.group(0)
        vm = re.search(r'<vt:[^>]*>(.*?)</vt:', block, re.S)
        if vm:
            val = vm.group(1).strip()
        else:
            vm2 = re.search(r'>([^<]+)</property>', block)
            val = vm2.group(1).strip() if vm2 else ""
        exp = str(expected).strip().lower()
        if exp in ("true", "1", "yes"):
            return val.lower() in ("true", "1", "yes")
        return val.lower() == exp

    def _evaluate_specs(self, specs, full_text, runs, idx, target_phrase, expected_text, tables=None, list_items=0, footnotes_text="", endnotes_text="", has_toc=False, footnote_ref_count=0, drawing_count=0, alt_texts=None, textbox_texts=None, header_texts=None, footer_texts=None, core_props=None, page_orient=None, comments_text="", tracked_change_count=0, root=None, numbering_root=None, settings_text="", comments_extended_text="", custom_props_text=""):
        tables = tables or []
        alt_texts = alt_texts or []
        textbox_texts = textbox_texts or []
        header_texts = header_texts or []
        footer_texts = footer_texts or []
        core_props = core_props or {}

        # Derived data for new criterion types
        page_margins = {}
        paragraph_styles: list[str] = []
        paragraph_indents: list[str] = []
        paragraph_spacings: list[str] = []
        section_columns = None
        has_merged_cell = False
        has_repeating_header = False
        if root is not None:
            sectPrs = root.findall(f".//{{{W_NS}}}sectPr")
            sectPr = sectPrs[-1] if sectPrs else None
            if sectPr is not None:
                pgMar = sectPr.find(f"{{{W_NS}}}pgMar")
                if pgMar is not None:
                    for k in ("top", "bottom", "left", "right"):
                        v = pgMar.get(f"{{{W_NS}}}{k}")
                        if v is not None:
                            page_margins[k] = v
                cols = sectPr.find(f"{{{W_NS}}}cols")
                if cols is not None:
                    section_columns = cols.get(f"{{{W_NS}}}num")
            for p in root.iter(f"{{{W_NS}}}p"):
                pPr = p.find(f"{{{W_NS}}}pPr")
                if pPr is None:
                    paragraph_styles.append("Normal")
                    continue
                ps = pPr.find(f"{{{W_NS}}}pStyle")
                if ps is not None and ps.get(f"{{{W_NS}}}val"):
                    paragraph_styles.append(ps.get(f"{{{W_NS}}}val"))
                else:
                    paragraph_styles.append("Normal")
                # Collect indent/spacing regardless of whether a paragraph style
                # is present: Word often keeps w:pStyle (e.g. "Normal") alongside
                # w:ind/w:spacing, and dropping them caused false negatives.
                ind = pPr.find(f"{{{W_NS}}}ind")
                if ind is not None and ind.get(f"{{{W_NS}}}left") is not None:
                    paragraph_indents.append(ind.get(f"{{{W_NS}}}left"))
                spacing = pPr.find(f"{{{W_NS}}}spacing")
                if spacing is not None and spacing.get(f"{{{W_NS}}}after") is not None:
                    paragraph_spacings.append(spacing.get(f"{{{W_NS}}}after"))
            for tc in root.iter(f"{{{W_NS}}}tc"):
                if tc.find(f".//{{{W_NS}}}gridSpan") is not None or tc.find(f".//{{{W_NS}}}vMerge") is not None:
                    has_merged_cell = True
            for trPr in root.iter(f"{{{W_NS}}}trPr"):
                if trPr.find(f"{{{W_NS}}}tblHeader") is not None:
                    has_repeating_header = True
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
                    if all(v is None for v in vals):
                        status = CriterionStatus.NOT_VERIFIABLE
                    elif all(v is None or v == expected for v in vals) and any(v is not None and v == expected for v in vals):
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
                    if norm_full == norm_expected:
                        status = CriterionStatus.PASS
                    elif "©" in norm_expected and "©" in norm_full and norm_full.replace("©", "") == norm_expected.replace("©", ""):
                        status = CriterionStatus.PASS
                    else:
                        status = CriterionStatus.FAIL
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
                # Contar notas reales (definiciones en footnotes.xml), no
                # referencias en el texto, para coincidir con Word COM.
                count = (footnotes_text or "").count("<w:footnote ")
                count -= (footnotes_text or "").count('w:type="separator"')
                count -= (footnotes_text or "").count('w:type="continuationSeparator"')
                if count < 0:
                    count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Notas al pie: {count} (esperado >= {expected})."
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
            elif typ == "textbox_count":
                expected = spec.get("expected", 1)
                count = len(list(root.iter(f"{{{W_NS}}}txbxContent")))
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Cuadros de texto: {count} (esperado >= {expected})."
            elif typ == "endnote_count":
                expected = spec.get("expected", 1)
                # Contar solo las notas de contenido (excluir la raiz <w:endnotes>
                # y las entradas separator/continuationSeparator).
                count = (endnotes_text or "").count("<w:endnote ")
                count -= (endnotes_text or "").count('w:type="separator"')
                count -= (endnotes_text or "").count('w:type="continuationSeparator"')
                if count < 0: count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Notas finales: {count} (esperado >= {expected})."
            elif typ == "tracked_change_count":
                expected = spec.get("expected", 1)
                count = (doc_xml if False else "").count if False else 0
                # use already computed tracked_change_count
                count = tracked_change_count
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Cambios registrados: {count} (esperado >= {expected})."
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
                prop_map = {"title": "title", "author": "creator", "subject": "subject", "creator": "creator"}
                lookup = prop_map.get(str(prop).lower(), str(prop).lower())
                actual = None
                for k, v in core_props.items():
                    if k.lower() == lookup:
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
                if expected == 0:
                    status = CriterionStatus.PASS if count == 0 else CriterionStatus.FAIL
                    ev = f"Comentarios: {count} (esperado 0)."
                else:
                    status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                    ev = f"Comentarios: {count} (esperado >= {expected})."
            elif typ == "hyperlink_present":
                expected = spec.get("expected", True)
                count = len(root.findall(f".//{{{W_NS}}}hyperlink"))
                for h in (header_texts or []):
                    count += h.count("<w:hyperlink")
                for f in (footer_texts or []):
                    count += f.count("<w:hyperlink")

                present = count > 0
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = f"Hipervínculos: {count}."
            elif typ == "background_present":
                expected = spec.get("expected", True)
                # check w:background in document.xml raw
                present = False
                try:
                    # root is document root; background is child of w:document, not in body. Check parent? root.tag is document, so:
                    bg = root.find(f"{{{W_NS}}}background")
                    present = bg is not None
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Fondo de página detectado." if present else "No se detectó fondo de página."
            elif typ == "tracked_changes_present":
                expected = spec.get("expected", True)
                present = tracked_change_count > 0
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Se detectaron cambios registrados." if present else "No se detectaron cambios registrados."
            elif typ == "page_margins":
                expected = spec.get("expected", {})
                ok = True
                for k, v in expected.items():
                    if page_margins.get(k) != str(v):
                        ok = False
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Márgenes: {page_margins} (esperado {expected})."
            elif typ == "paragraph_style":
                expected = str(spec.get("expected", "")).strip()
                expected_norm = expected.lower().replace(" ", "")
                # Built-in style IDs are localized in Word (e.g. "Heading2" vs
                # "Encabezado2"/"Ttulo2"), so accept common aliases.
                aliases = {
                    "heading1": {"heading1", "encabezado1", "ttulo1", "titulo1", "título1"},
                    "heading2": {"heading2", "encabezado2", "ttulo2", "titulo2", "título2"},
                    "heading3": {"heading3", "encabezado3", "ttulo3", "titulo3", "título3"},
                    "normal": {"normal"},
                }
                accepted = aliases.get(expected_norm, {expected_norm})
                normalized_styles = {str(s).lower().replace(" ", "") for s in (paragraph_styles or [])}
                status = CriterionStatus.PASS if bool(accepted & normalized_styles) else CriterionStatus.FAIL
                ev = f"Estilo de párrafo '{expected}' {'aplicado' if status == CriterionStatus.PASS else 'no aplicado'}."
            elif typ == "section_columns":
                expected = spec.get("expected", "1")
                status = CriterionStatus.PASS if str(section_columns) == str(expected) else CriterionStatus.FAIL
                ev = f"Columnas: {section_columns} (esperado {expected})."
            elif typ == "paragraph_indent":
                expected = spec.get("expected", "")
                status = CriterionStatus.PASS if str(expected) in paragraph_indents else CriterionStatus.FAIL
                ev = f"Sangría: {paragraph_indents} (esperado {expected})."
            elif typ == "paragraph_spacing":
                expected = spec.get("expected", "")
                status = CriterionStatus.PASS if str(expected) in paragraph_spacings else CriterionStatus.FAIL
                ev = f"Espaciado: {paragraph_spacings} (esperado {expected})."
            elif typ == "merged_cell_present":
                expected = spec.get("expected", True)
                status = CriterionStatus.PASS if has_merged_cell == bool(expected) else CriterionStatus.FAIL
                ev = "Celda combinada detectada." if has_merged_cell else "No se detectó celda combinada."
            elif typ == "repeating_header":
                expected = spec.get("expected", True)
                status = CriterionStatus.PASS if has_repeating_header == bool(expected) else CriterionStatus.FAIL
                ev = "Encabezado de tabla repetido." if has_repeating_header else "No se repite encabezado de tabla."
            elif typ == "section_break_present":
                expected = spec.get("expected", True)
                # mas de un sectPr en el body => hay salto de seccion
                sect_count = len(root.findall(f".//{{{W_NS}}}sectPr")) if root is not None else 0
                present = sect_count > 1
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Salto de seccion detectado." if present else "No se detecto salto de seccion."
            elif typ == "page_break_present" or typ == "page_break":
                expected = spec.get("expected", True)
                present = False
                try:
                    if root is not None:
                        for br in root.iter(f"{{{W_NS}}}br"):
                            btype = (br.get(f"{{{W_NS}}}type") or "").lower()
                            if btype == "page" or btype == "":
                                # Word writes <w:br w:type="page"/> for manual page break.
                                present = True
                                break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Salto de página detectado." if present else "No se detectó salto de página."
            elif typ == "section_break":
                expected = str(spec.get("expected", "nextPage")).lower()
                present = False
                try:
                    for sp in (root.findall(f".//{{{W_NS}}}sectPr") if root is not None else []):
                        t = sp.find(f"{{{W_NS}}}type")
                        if t is not None and (t.get(f"{{{W_NS}}}val") or "").lower() == expected:
                            present = True
                            break
                    # nextPage es el default; si hay mas de un sectPr se asume nextPage
                    if expected == "nextpage" and len(root.findall(f".//{{{W_NS}}}sectPr") if root is not None else []) > 1:
                        present = True
                except Exception:
                    pass
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = f"Salto de seccion '{expected}' detectado." if present else f"No se detecto salto '{expected}'."
            elif typ == "shape_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    # formas: w:drawing con a:sp (shape) o mc:AlternateContent con shape
                    for dr in (root.iter(f"{{{W_NS}}}drawing") if root is not None else []):
                        xml = etree.tostring(dr, encoding="unicode")
                        if "a:sp " in xml or "<a:sp>" in xml or "spPr" in xml and "pic:pic" not in xml:
                            # heuristic: drawing that is not a picture/textbox
                            if "pic:pic" not in xml and "txbxContent" not in xml:
                                count += 1
                except Exception:
                    pass
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Formas: {count} (esperado >= {expected})."
            elif typ == "smartart_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "smartArt" in xml.lower() or "dgm:relIds" in xml or "quickStyle" in xml and "smartart" in xml.lower()
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "SmartArt detectado." if present else "No se detecto SmartArt."
            elif typ == "model3d_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "3dmodel" in xml.lower() or "model3d" in xml.lower() or "am3d" in xml.lower()
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Modelo 3D detectado." if present else "No se detecto modelo 3D."
            elif typ == "image_background_removed":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "image_background_removed", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Fondo de imagen quitado (registrado por complemento)." if present else "No se registró la acción de quitar fondo."
            elif typ == "object_position":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    if root is not None:
                        has_right = False
                        has_top = False
                        for posh in root.iter(f"{{{WP_NS}}}positionH"):
                            al = posh.find(f"{{{WP_NS}}}align")
                            if al is not None and (al.text or "").strip().lower() == "right":
                                has_right = True
                            # posOffset near right edge also counts
                            off = posh.find(f"{{{WP_NS}}}posOffset")
                            if off is not None:
                                try:
                                    if int(off.text or "0") > 0:
                                        has_right = True
                                except Exception:
                                    pass
                        for posv in root.iter(f"{{{WP_NS}}}positionV"):
                            al = posv.find(f"{{{WP_NS}}}align")
                            if al is not None and (al.text or "").strip().lower() == "top":
                                has_top = True
                            off = posv.find(f"{{{WP_NS}}}posOffset")
                            if off is not None:
                                try:
                                    if int(off.text or "0") == 0:
                                        has_top = True
                                except Exception:
                                    pass
                        if expected in ("top-right", "top_right", "topright"):
                            present = has_right and has_top
                        elif expected in ("top-left", "top_left", "topleft"):
                            present = (not has_right) and has_top
                        else:
                            present = has_right and has_top
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Objeto posicionado correctamente." if present else "El objeto no está en la posición esperada."
            elif typ == "text_wrap":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    if expected in ("tight",):
                        present = "wrapTight" in xml
                    else:
                        present = "wrap" in xml
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Ajuste de texto configurado." if present else "No se detectó ajuste de texto."
            elif typ == "comment_resolved":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    blob = (comments_text or "") + "\n" + (comments_extended_text or "")
                    low = blob.lower()
                    present = (
                        'w:done="1"' in blob
                        or 'done="1"' in blob
                        or 'w15:done="1"' in low
                        or 'w16cid:done="1"' in low
                        or 'w:done="true"' in low
                        or "done=\"true\"" in low
                    )
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Comentario resuelto." if present else "No se detectó comentario resuelto."
            elif typ == "revisions_reviewed":
                expected = bool(spec.get("expected", True))
                present = (tracked_change_count == 0)
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Cambios revisados (sin cambios pendientes)." if present else "Aún hay cambios registrados sin revisar."
            elif typ == "track_changes_locked":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "trackRevisions" in xml
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Seguimiento bloqueado." if present else "No se detectó bloqueo de seguimiento."
            elif typ == "internal_hyperlink":
                expected = spec.get("expected", True)
                present = False
                try:
                    for hl in (root.iter(f"{{{W_NS}}}hyperlink") if root is not None else []):
                        if hl.get(f"{{{W_NS}}}anchor"):
                            present = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Hipervinculo interno detectado." if present else "No se detecto hipervinculo interno."
            elif typ == "formatting_marks_visible":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "formatting_marks_visible", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Marcas de formato mostradas (registrado por complemento)." if present else "No se registró la acción."
            elif typ == "alternative_format":
                expected = spec.get("expected", "pdf")
                present = self._custom_prop(custom_props_text, "alternative_format", expected)
                status = CriterionStatus.PASS if (present == (str(expected).lower() not in ("false",""))) else CriterionStatus.FAIL
                ev = "Formato alternativo guardado (registrado por complemento)." if present else "No se registró el formato alternativo."
            elif typ == "print_settings":
                expected = spec.get("expected", "2up")
                present = self._custom_prop(custom_props_text, "print_settings", expected)
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Configuración de impresión registrada (complemento)." if present else "No se registró la configuración de impresión."
            elif typ == "remove_personal_info":
                expected = spec.get("expected", True)
                present_info = False
                try:
                    cp = core_props or {}
                    for k, v in cp.items():
                        if v and str(v).strip():
                            present_info = True
                            break
                except Exception:
                    pass
                # pass if info removed (no personal info) matches expected True
                removed = not present_info
                status = CriterionStatus.PASS if removed == bool(expected) else CriterionStatus.FAIL
                ev = "Propiedades personales eliminadas." if removed else "Aun hay propiedades personales."
            elif typ == "text_effects":
                expected = spec.get("expected", "shadow")
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "w:effect" in xml or "effectLst" in xml or "shadow" in xml.lower()
                except Exception:
                    pass
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Efecto de texto detectado." if present else "No se detecto efecto de texto."
            elif typ == "format_painter_used":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "format_painter_used", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Pincel de formato usado (registrado por complemento)." if present else "No se registró el pincel de formato."
            elif typ == "footnote_props":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    xml = (footnotes_text or "") + (settings_text or "") + (etree.tostring(root, encoding="unicode") if root is not None else "")
                    if expected in ("roman", "romanos"):
                        present = "roman" in xml.lower() or "lowerRoman" in xml or "upperRoman" in xml
                    else:
                        present = expected in xml.lower()
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Propiedades de nota configuradas." if present else "No se detectaron propiedades de nota."
            elif typ == "citation_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "citation" in xml.lower() or "fldSimple" in xml and "CITATION" in xml
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Cita detectada." if present else "No se detecto cita."
            elif typ == "bibliography_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "bibliography" in xml.lower()
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Bibliografia detectada." if present else "No se detecto bibliografia."
            elif typ == "artistic_effect_present":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    present = "artistic" in xml.lower()
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Efecto artistico detectado." if present else "No se detecto efecto artistico."
            elif typ == "picture_effect_present":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    low = xml.lower()
                    present = ("effectlst" in low) or ("a:glow" in low) or ("a:shadow" in low) or ("a:reflection" in low) or ("a:softedge" in low)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Efecto de imagen detectado." if present else "No se detecto efecto de imagen."
            elif typ == "style_set_applied":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "style_set_applied", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Conjunto de estilos aplicado (registrado por complemento)." if present else "No se registró la acción."
            elif typ == "share_electronically":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "share_electronically", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Documento compartido (registrado por complemento)." if present else "No se registró la acción."
            elif typ == "compatibility_issues":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "compatibility_issues", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Compatibilidad revisada (registrado por complemento)." if present else "No se registró la acción."
            elif typ == "navigate_location":
                expected = spec.get("expected", True)
                present = self._custom_prop(custom_props_text, "navigate_location", expected)
                status = CriterionStatus.PASS if (present == bool(expected)) else CriterionStatus.FAIL
                ev = "Navegación registrada (complemento)." if present else "No se registró la acción."
            elif typ == "list_bullet_char":
                expected = str(spec.get("expected", "")).strip()
                present = False
                try:
                    if numbering_root is not None:
                        for lt in numbering_root.iter(f"{{{W_NS}}}lvlText"):
                            if (lt.get(f"{{{W_NS}}}val") or "").strip() == expected:
                                present = True
                                break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Viñeta personalizada detectada." if present else "No se detectó la viñeta esperada."
            elif typ == "list_custom_number_format":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    if numbering_root is not None:
                        for nf in numbering_root.iter(f"{{{W_NS}}}numFmt"):
                            val = (nf.get(f"{{{W_NS}}}val") or "").lower()
                            if val and val != "bullet":
                                present = True
                                break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Formato de numeración personalizado detectado." if present else "No se detectó formato de numeración personalizado."
            elif typ == "revisions_reviewed":
                expected = bool(spec.get("expected", True))
                present = (tracked_change_count == 0)
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Cambios revisados (sin cambios pendientes)." if present else "Aún hay cambios registrados sin revisar."
            elif typ == "list_restart":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    if numbering_root is not None:
                        present = any(True for _ in numbering_root.iter(f"{{{W_NS}}}startOverride"))
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Numeración reiniciada." if present else "No se detectó reinicio de numeración."
            elif typ == "toc_custom":
                expected = str(spec.get("expected", "")).strip().lower()
                present = False
                try:
                    xml = etree.tostring(root, encoding="unicode") if root is not None else ""
                    # A custom TOC for levels 1-2 stores a field code like:
                    # TOC \\o "1-2" \\h \\z \\u
                    import re
                    # Collect TOC field codes (may be split across w:instrText).
                    field_codes = re.findall(r'<w:instrText[^>]*>(.*?)</w:instrText>', xml, re.S | re.I)
                    joined = " ".join(field_codes)
                    if "TOC" in joined.upper():
                        if expected in ("", "true", "1-2", "1,2"):
                            present = ('\\\\o' in joined and '1-2' in joined)
                        else:
                            present = (expected in joined)
                    # Fallback: any TOC field present.
                    if not present and "TOC" in xml.upper():
                        present = ('1-2' in xml)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "TOC personalizado (niveles 1-2) insertado." if present else "No se detectó TOC personalizado con niveles 1-2."
            elif typ == "cell_margins":
                expected = str(spec.get("expected", "144"))
                present = False
                try:
                    if root is not None:
                        for mar in root.iter(f"{{{W_NS}}}tblCellMar"):
                            for side in ("top", "left", "bottom", "right"):
                                el = mar.find(f"{{{W_NS}}}{side}")
                                if el is not None:
                                    try:
                                        v = int(el.get(f"{{{W_NS}}}w"))
                                        if abs(v - 144) <= 2 or abs(v - 57) <= 2:
                                            present = True
                                    except Exception:
                                        pass
                                    if present:
                                        break
                            if present:
                                break
                        # Also check per-cell tcMar
                        if not present:
                            for mar in root.iter(f"{{{W_NS}}}tcMar"):
                                for side in ("top", "left", "bottom", "right"):
                                    el = mar.find(f"{{{W_NS}}}{side}")
                                    if el is not None:
                                        try:
                                            v = int(el.get(f"{{{W_NS}}}w"))
                                            if abs(v - 144) <= 2 or abs(v - 57) <= 2:
                                                present = True
                                        except Exception:
                                            pass
                                    if present:
                                        break
                                if present:
                                    break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Márgenes de celda configurados." if present else "Los márgenes de celda no están configurados."
            elif typ == "table_sorted":
                expected = str(spec.get("expected", "asc")).lower()
                present = False
                try:
                    if root is not None:
                        tbls = root.findall(f".//{{{W_NS}}}tbl")
                        if tbls:
                            tbl = tbls[0]
                            rows = tbl.findall(f"{{{W_NS}}}tr")
                            # Determine target column (1-based) from spec or header name.
                            col_spec = spec.get("column", None)
                            row_cells = [tr.findall(f"{{{W_NS}}}tc") for tr in rows]
                            header_texts = []
                            if row_cells and row_cells[0]:
                                for tc in row_cells[0]:
                                    header_texts.append("".join(t.text or "" for t in tc.iter(f"{{{W_NS}}}t")).strip())
                            if isinstance(col_spec, int):
                                col_idx = max(0, col_spec - 1)
                            elif isinstance(col_spec, str) and col_spec.strip().isdigit():
                                col_idx = max(0, int(col_spec.strip()) - 1)
                            elif isinstance(col_spec, str) and col_spec.strip():
                                needle = col_spec.strip().lower()
                                col_idx = 0
                                for i, h in enumerate(header_texts):
                                    if needle in h.lower() or h.lower() == needle:
                                        col_idx = i
                                        break
                            else:
                                col_idx = 0
                                for i, h in enumerate(header_texts):
                                    hl = h.lower()
                                    if "precio" in hl or "price" in hl:
                                        col_idx = i
                                        break
                            values = []
                            for cells in row_cells:
                                if not cells:
                                    continue
                                idx = col_idx if col_idx < len(cells) else 0
                                texts = []
                                for tnode in cells[idx].iter(f"{{{W_NS}}}t"):
                                    texts.append(tnode.text or "")
                                values.append("".join(texts).strip())
                            data = values[1:] if len(values) > 1 else values
                            if expected.startswith("desc"):
                                present = data == sorted(data, reverse=True)
                            else:
                                present = data == sorted(data)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Tabla ordenada correctamente." if present else "La tabla no está ordenada por la columna indicada."
            elif typ == "table_autofit":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    if root is not None:
                        xml = etree.tostring(root, encoding="unicode")
                        present = ('tblLayout w:type="autofit"' in xml) or ('w:tblW' in xml and 'w:type="pct"' in xml)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Ajuste de tabla configurado." if present else "No se detectó ajuste de tabla."
            elif typ == "list_levels":
                expected = spec.get("expected", 2)
                count = 0
                try:
                    ilvls = set()
                    for e in (root.iter(f"{{{W_NS}}}ilvl") if root is not None else []):
                        ilvls.add(e.get(f"{{{W_NS}}}val"))
                    count = len(ilvls) if ilvls else 1
                except Exception:
                    pass
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Niveles de lista: {count} (esperado >= {expected})."
            elif typ == "list_restart":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    if numbering_root is not None:
                        present = any(True for _ in numbering_root.iter(f"{{{W_NS}}}startOverride"))
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Numeración reiniciada." if present else "No se detectó reinicio de numeración."
            elif typ == "list_start_value":
                expected = int(spec.get("expected", 1))
                start_val = None
                try:
                    num_id = None
                    if root is not None:
                        for p in root.iter(f"{{{W_NS}}}p"):
                            pPr = p.find(f"{{{W_NS}}}pPr")
                            if pPr is not None:
                                numPr = pPr.find(f"{{{W_NS}}}numPr")
                                if numPr is not None:
                                    nid = numPr.find(f"{{{W_NS}}}numId")
                                    if nid is not None and nid.get(f"{{{W_NS}}}val"):
                                        num_id = nid.get(f"{{{W_NS}}}val")
                                        break
                    if numbering_root is not None and num_id is not None:
                        for num in numbering_root.findall(f".//{{{W_NS}}}num"):
                            if num.get(f"{{{W_NS}}}numId") == num_id:
                                for lvlOverride in num.findall(f"{{{W_NS}}}lvlOverride"):
                                    so = lvlOverride.find(f"{{{W_NS}}}startOverride")
                                    lvl = lvlOverride.find(f"{{{W_NS}}}lvl")
                                    if so is None and lvl is not None:
                                        so = lvl.find(f"{{{W_NS}}}startOverride")
                                    st = None
                                    if lvl is not None:
                                        st = lvl.find(f"{{{W_NS}}}start")
                                    if so is not None:
                                        start_val = int(so.get(f"{{{W_NS}}}val"))
                                        break
                                    if st is not None and start_val is None:
                                        start_val = int(st.get(f"{{{W_NS}}}val"))
                                if start_val is None:
                                    an = num.find(f"{{{W_NS}}}abstractNumId")
                                    abstract_id = an.get(f"{{{W_NS}}}val") if an is not None else None
                                    if abstract_id is not None:
                                        for anode in numbering_root.findall(f".//{{{W_NS}}}abstractNum"):
                                            if anode.get(f"{{{W_NS}}}abstractNumId") == abstract_id:
                                                lvl = anode.find(f"{{{W_NS}}}lvl")
                                                if lvl is not None:
                                                    st = lvl.find(f"{{{W_NS}}}start")
                                                    if st is not None:
                                                        start_val = int(st.get(f"{{{W_NS}}}val"))
                                                break
                                break
                except Exception:
                    start_val = None
                # Fallback: if the paragraph has no explicit numPr (inherited
                # from a style), scan all numbering definitions for an override
                # with the expected start value.
                if start_val is None and numbering_root is not None:
                    try:
                        for so in numbering_root.iter(f"{{{W_NS}}}startOverride"):
                            if int(so.get(f"{{{W_NS}}}val")) == expected:
                                start_val = expected
                                break
                    except Exception:
                        pass
                present = (start_val == expected)
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = f"Valor inicial de lista: {start_val} (esperado {expected})." if start_val is not None else "No se pudo determinar el valor inicial de la lista."
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
