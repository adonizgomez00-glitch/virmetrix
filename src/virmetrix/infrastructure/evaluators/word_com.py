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
        # Try to connect to Word (use DispatchEx to avoid attaching to a user instance)
        import pythoncom

        pythoncom.CoInitialize()
        word = None
        try:
            import win32com.client

            word = win32com.client.DispatchEx("Word.Application")
            version = word.Version
            return True, f"Word disponible (v{version})."
        except Exception as e:
            return False, f"Microsoft Word no accesible: {type(e).__name__}: {e}"
        finally:
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

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

            specs = getattr(exercise, "criteria_specs", None) or []
            if specs:
                criteria = self._evaluate_specs_com(specs, doc, full_text, target_phrase, expected_text, occurrences)
            else:
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

    @staticmethod
    def _com_range_prop(rng, prop: str):
        try:
            if prop == "bold":
                b = rng.Font.Bold
                if b in (-1, True, 1):
                    return True
                if b in (0, False):
                    return False
                return None
            if prop == "italic":
                it = rng.Font.Italic
                if it in (-1, True, 1):
                    return True
                if it in (0, False):
                    return False
                return None
            if prop == "size_pt":
                s = rng.Font.Size
                if s == 999999:
                    return None
                return float(s)
            if prop == "font_name":
                return rng.Font.Name
            if prop == "color":
                c = rng.Font.Color
                # Font.Color es un entero OLE_COLOR (BGR), no una cadena hex.
                # Convertir a cadena hex RGB (p. ej. 255 -> "FF0000") para
                # comparar con el expected del ejercicio.
                if isinstance(c, int) and c >= 0:
                    r = c & 0xFF
                    g = (c >> 8) & 0xFF
                    b = (c >> 16) & 0xFF
                    return f"{r:02X}{g:02X}{b:02X}"
                return None
        except Exception:
            return None
        return None

    def _com_range_prop_all(self, rng, prop, expected):
        """Evalúa una propiedad en todos los caracteres del rango (ignora espacios)."""
        vals = []
        try:
            chars = rng.Characters
            count = chars.Count
        except Exception:
            count = 0
        if count and count < 100000:
            for i in range(1, count + 1):
                try:
                    ch = chars(i)
                    txt = getattr(ch, "Text", "") or ""
                    if txt.strip() == "":
                        continue
                    vals.append(self._com_range_prop(ch, prop))
                except Exception:
                    vals.append(None)
        else:
            vals.append(self._com_range_prop(rng, prop))
        if not vals:
            vals.append(self._com_range_prop(rng, prop))
        if all(v is None for v in vals):
            return None
        if all(v is None or v == expected for v in vals) and any(v is not None and v == expected for v in vals):
            return True
        return False

    def _evaluate_specs_com(self, specs, doc, full_text, target_phrase, expected_text, occurrences):
        criteria = []
        for spec in specs:
            cid = spec["id"]
            typ = spec.get("type")
            target = spec.get("target", target_phrase)
            occs = self._occurrences(full_text, target) if target else []
            if typ == "target_occurrences":
                expected = spec.get("expected", 1)
                status = CriterionStatus.PASS if len(occs) == expected else CriterionStatus.FAIL
                ev = f"Ocurrencias de '{target}': {len(occs)} (esperado {expected})."
            elif typ == "run_property":
                prop = spec["property"]
                expected = spec["expected"]
                if not occs:
                    status = CriterionStatus.FAIL
                    ev = f"No se encontró el texto objetivo '{target}'."
                else:
                    start = occs[0]
                    end = start + len(target)
                    rng = doc.Range(start, end)
                    val = self._com_range_prop(rng, prop)
                    agg = self._com_range_prop_all(rng, prop, expected)
                    if agg is None:
                        status = CriterionStatus.NOT_VERIFIABLE
                    elif agg:
                        status = CriterionStatus.PASS
                    else:
                        status = CriterionStatus.FAIL
                    ev = f"Propiedad '{prop}' en objetivo: {val}."
            elif typ == "neighbor_not_formatted":
                forbidden = spec.get("forbidden", {})
                if not occs:
                    status = CriterionStatus.FAIL
                    ev = f"No se encontró el texto objetivo '{target}'."
                else:
                    start = occs[0]
                    end = start + len(target)
                    leak = False
                    try:
                        if start > 0:
                            prev_rng = doc.Range(start - 1, start)
                            if all(self._com_range_prop(prev_rng, p) == v for p, v in forbidden.items()):
                                leak = True
                        if end < len(full_text):
                            next_rng = doc.Range(end, end + 1)
                            if all(self._com_range_prop(next_rng, p) == v for p, v in forbidden.items()):
                                leak = True
                    except Exception:
                        pass
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
                try:
                    count = doc.Tables.Count
                except Exception:
                    count = 0
                status = CriterionStatus.PASS if count == expected else CriterionStatus.FAIL
                ev = f"Tablas en el documento: {count} (esperado {expected})."
            elif typ == "table_dimensions":
                tindex = spec.get("table_index", 0)
                exp_rows = spec.get("rows")
                exp_cols = spec.get("cols")
                try:
                    tbl = doc.Tables(tindex + 1)
                    rows = tbl.Rows.Count
                    cols = tbl.Columns.Count
                except Exception:
                    rows = 0
                    cols = 0
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
                try:
                    tbl = doc.Tables(tindex + 1)
                    actual = tbl.Cell(r + 1, c + 1).Range.Text
                    actual = actual.rstrip("\r\x07")
                except Exception:
                    actual = None
                if actual is None:
                    status = CriterionStatus.FAIL
                    ev = f"Celda ({r},{c}) no accesible."
                else:
                    status = CriterionStatus.PASS if actual == expected else CriterionStatus.FAIL
                    ev = f"Celda ({r},{c}): '{actual}' (esperado '{expected}')."
            elif typ == "list_items_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    for p in doc.Paragraphs:
                        try:
                            lf = p.Range.ListFormat
                            if getattr(lf, "ListType", 0) != 0:
                                count += 1
                                continue
                            if getattr(lf, "ListValue", 0) not in (None, 0):
                                count += 1
                                continue
                            style = getattr(p, "Style", None)
                            name = str(getattr(style, "Name", "") or getattr(style, "NameLocal", "") or "")
                            if "list" in name.lower():
                                count += 1
                        except Exception:
                            pass
                except Exception:
                    pass
                status = CriterionStatus.PASS if count == expected else CriterionStatus.FAIL
                ev = f"Elementos de lista: {count} (esperado {expected})."
            elif typ == "footnote_count":
                expected = spec.get("expected", 1)
                try:
                    count = doc.Footnotes.Count
                except Exception:
                    count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Notas al pie: {count} (esperado >= {expected})."
            elif typ == "footnote_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for fn in doc.Footnotes:
                        if expected in (fn.Range.Text or ""):
                            ok = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "La nota al pie contiene el texto esperado." if status == CriterionStatus.PASS else "La nota al pie no contiene el texto esperado."
            elif typ == "endnote_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for en in doc.Endnotes:
                        if expected in (en.Range.Text or ""):
                            ok = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "La nota final contiene el texto esperado." if status == CriterionStatus.PASS else "La nota final no contiene el texto esperado."
            elif typ == "toc_present":
                ok = False
                try:
                    ok = doc.TablesOfContents.Count > 0
                except Exception:
                    try:
                        ok = "TOC" in (doc.Content.Text or "")
                    except Exception:
                        ok = False
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Se detectó tabla de contenido." if status == CriterionStatus.PASS else "No se detectó tabla de contenido."
            elif typ == "drawing_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    count = doc.InlineShapes.Count + doc.Shapes.Count
                except Exception:
                    try:
                        count = doc.InlineShapes.Count
                    except Exception:
                        count = 0
                status = CriterionStatus.PASS if count == expected else CriterionStatus.FAIL
                ev = f"Imágenes/dibujos: {count} (esperado {expected})."
            elif typ == "alt_text_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for shp in doc.InlineShapes:
                        if expected in (getattr(shp, "AlternativeText", "") or ""):
                            ok = True
                            break
                except Exception:
                    pass
                try:
                    for shp in doc.Shapes:
                        if expected in (getattr(shp, "AlternativeText", "") or ""):
                            ok = True
                            break
                except Exception:
                    pass
                # Fallback: inspeccionar el XML (docPr/@descr)
                if not ok:
                    try:
                        xml = str(getattr(doc, "WordOpenXML", "") or "")
                        if expected and expected.lower() in xml.lower():
                            ok = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Texto alternativo correcto." if status == CriterionStatus.PASS else "Texto alternativo ausente o incorrecto."
            elif typ == "textbox_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for shp in doc.Shapes:
                        try:
                            if expected in (shp.TextFrame.TextRange.Text or ""):
                                ok = True
                                break
                        except Exception:
                            pass
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Cuadro de texto contiene el texto esperado." if status == CriterionStatus.PASS else "Cuadro de texto ausente o incorrecto."
            elif typ == "textbox_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    xml = str(getattr(doc, "WordOpenXML", "") or "")
                    count = xml.count("<w:txbxContent")
                except Exception:
                    count = 0
                # Fallback: contar shapes de tipo cuadro de texto
                if count == 0:
                    try:
                        for shp in doc.Shapes:
                            if getattr(shp, "Type", None) == 17:
                                count += 1
                    except Exception:
                        pass
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Cuadros de texto: {count} (esperado >= {expected})."
            elif typ == "endnote_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    count = doc.Endnotes.Count
                except Exception:
                    count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Notas finales: {count} (esperado >= {expected})."
            elif typ == "tracked_change_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    count = doc.Revisions.Count
                except Exception:
                    count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Cambios registrados: {count} (esperado >= {expected})."
            elif typ == "header_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for si in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(si)
                        for hi in (1, 2, 3):
                            try:
                                txt = sec.Headers(hi).Range.Text or ""
                                if expected in txt:
                                    ok = True
                                    break
                            except Exception:
                                continue
                        if ok:
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Encabezado contiene el texto esperado." if status == CriterionStatus.PASS else "Encabezado ausente o incorrecto."
            elif typ == "footer_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for si in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(si)
                        for fi in (1, 2, 3):
                            try:
                                txt = sec.Footers(fi).Range.Text or ""
                                if expected in txt:
                                    ok = True
                                    break
                            except Exception:
                                continue
                        if ok:
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Pie de página contiene el texto esperado." if status == CriterionStatus.PASS else "Pie de página ausente o incorrecto."
            elif typ == "doc_property":
                prop = spec.get("property", "")
                expected = spec.get("expected", "")
                # Nombres canonicos (EN) + equivalentes localizados (ES) + indice
                # WdBuiltInProperty, porque en Word localizado el Name puede
                # estar traducido y el acceso por cadena inglesa fallar.
                canon = {"title": (1, ["Title", "Título", "Titulo"]),
                        "subject": (2, ["Subject", "Asunto"]),
                        "author": (3, ["Author", "Autor"]),
                        "creator": (3, ["Author", "Autor"]),
                        "keywords": (4, ["Keywords", "Palabras clave"]),
                        "comments": (5, ["Comments", "Comentarios"]),
                        "category": (18, ["Category", "Categoría", "Categoria"]),
                        "manager": (20, ["Manager", "Administrador"]),
                        "company": (21, ["Company", "Empresa"])}
                key = str(prop).lower()
                idx, names = canon.get(key, (None, [str(prop)]))
                actual = None
                # 1) Por indice WdBuiltInProperty (mas robusto)
                if idx is not None:
                    try:
                        actual = doc.BuiltInDocumentProperties(idx).Value
                    except Exception:
                        actual = None
                # 2) Por nombre (EN/ES)
                if actual is None:
                    for nm in names:
                        try:
                            actual = doc.BuiltInDocumentProperties(nm).Value
                            break
                        except Exception:
                            continue
                # 3) Custom properties
                if actual is None:
                    for nm in names + [str(prop)]:
                        try:
                            actual = doc.CustomDocumentProperties(nm).Value
                            break
                        except Exception:
                            continue
                # Normalizar comparacion (str, strip, case-insensitive)
                status = CriterionStatus.PASS if str(actual).strip().lower() == str(expected).strip().lower() else CriterionStatus.FAIL
                ev = f"Propiedad '{prop}': '{actual}' (esperado '{expected}')."
            elif typ == "page_orientation":
                expected = spec.get("expected", "portrait").lower()
                actual_str = None
                try:
                    sections = getattr(doc, "Sections", None)
                    ps = sections(sections.Count).PageSetup if (sections is not None and sections.Count) else doc.PageSetup
                    # Prefer physical page dimensions (unambiguous) over the
                    # Orientation enum, whose numeric value varies by Word version.
                    try:
                        if ps.PageWidth and ps.PageHeight:
                            actual_str = "landscape" if float(ps.PageWidth) > float(ps.PageHeight) else "portrait"
                    except Exception:
                        actual_str = None
                    if actual_str is None:
                        orient = ps.Orientation
                        if orient in (0, "wdOrientPortrait", "Portrait"):
                            actual_str = "portrait"
                        elif orient in (1, 2, "wdOrientLandscape", "Landscape"):
                            actual_str = "landscape"
                except Exception:
                    try:
                        ps = doc.PageSetup
                        if ps.PageWidth and ps.PageHeight:
                            actual_str = "landscape" if float(ps.PageWidth) > float(ps.PageHeight) else "portrait"
                    except Exception:
                        actual_str = None
                if actual_str is None:
                    actual_str = "unknown"
                # Resolve ambiguity for orient==1: if dimensions are available we
                # already set actual_str; otherwise assume modern mapping (1=landscape).
                status = CriterionStatus.PASS if actual_str == expected else CriterionStatus.FAIL
                ev = f"Orientación: {actual_str} (esperado {expected})."
            elif typ == "comment_contains":
                expected = spec.get("expected", "")
                ok = False
                try:
                    for cm in doc.Comments:
                        if expected in (cm.Range.Text or ""):
                            ok = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = "Comentario contiene el texto esperado." if status == CriterionStatus.PASS else "Comentario ausente o incorrecto."
            elif typ == "comment_count":
                expected = spec.get("expected", 1)
                try:
                    count = doc.Comments.Count
                except Exception:
                    count = 0
                if expected == 0:
                    status = CriterionStatus.PASS if count == 0 else CriterionStatus.FAIL
                    ev = f"Comentarios: {count} (esperado 0)."
                else:
                    status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                    ev = f"Comentarios: {count} (esperado >= {expected})."
            elif typ == "hyperlink_present":
                expected = spec.get("expected", True)
                count = 0
                try:
                    count = doc.Hyperlinks.Count
                except Exception:
                    count = 0
                # Incluir hipervinculos en encabezados/pies de pagina
                try:
                    for si in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(si)
                        for hi in (1, 2, 3):
                            try:
                                count += sec.Headers(hi).Range.Hyperlinks.Count
                            except Exception:
                                pass
                            try:
                                count += sec.Footers(hi).Range.Hyperlinks.Count
                            except Exception:
                                pass
                except Exception:
                    pass
                present = count > 0
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = f"Hipervínculos: {count}."
            elif typ == "background_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    xml = str(getattr(doc, "WordOpenXML", "") or "")
                    present = "<w:background" in xml.lower()
                except Exception:
                    present = False
                if not present:
                    try:
                        present = bool(getattr(doc, "Background", None)) and getattr(doc.Background, "Fill", None) is not None
                    except Exception:
                        present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Fondo de página detectado." if present else "No se detectó fondo de página."
            elif typ == "tracked_changes_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    present = doc.Revisions.Count > 0
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Se detectaron cambios registrados." if present else "No se detectaron cambios registrados."
            elif typ == "page_margins":
                expected = spec.get("expected", {})
                ok = True
                try:
                    sections = getattr(doc, "Sections", None)
                    ps = sections(sections.Count).PageSetup if (sections is not None and sections.Count) else doc.PageSetup
                    if "top" in expected and abs(ps.TopMargin - float(expected["top"]) / 20.0) > 1:
                        ok = False
                    if "bottom" in expected and abs(ps.BottomMargin - float(expected["bottom"]) / 20.0) > 1:
                        ok = False
                    if "left" in expected and abs(ps.LeftMargin - float(expected["left"]) / 20.0) > 1:
                        ok = False
                    if "right" in expected and abs(ps.RightMargin - float(expected["right"]) / 20.0) > 1:
                        ok = False
                except Exception:
                    ok = False
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Márgenes verificados (esperado {expected})."
            elif typ == "paragraph_style":
                expected = spec.get("expected", "")
                ok = False
                norm_expected = str(expected).lower().replace(" ", "")
                try:
                    for p in doc.Paragraphs:
                        style = getattr(p, "Style", None)
                        if style is None:
                            continue
                        for name in (getattr(style, "Name", None), getattr(style, "NameLocal", None)):
                            if name and norm_expected in str(name).lower().replace(" ", ""):
                                ok = True
                                break
                        if ok:
                            break
                except Exception:
                    pass
                # Fallback locale-independent: inspect the paragraph XML for a
                # w:pStyle w:val matching the expected style id (e.g. Heading1).
                if not ok:
                    try:
                        import re
                        for p in doc.Paragraphs:
                            xml = str(getattr(p.Range, "XML", "") or "")
                            m = re.search(r'w:pStyle\s+w:val="([^"]+)"', xml)
                            if m and m.group(1).lower().replace(" ", "") == norm_expected:
                                ok = True
                                break
                    except Exception:
                        pass
                if not ok:
                    try:
                        xml = str(getattr(doc, "WordOpenXML", "") or "")
                        if norm_expected and norm_expected in xml.lower().replace(" ", ""):
                            # ensure it is a pStyle val, not arbitrary text
                            import re
                            for m in re.finditer(r'w:pStyle\s+w:val="([^"]+)"', xml):
                                if m.group(1).lower().replace(" ", "") == norm_expected:
                                    ok = True
                                    break
                    except Exception:
                        pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Estilo '{expected}' {'aplicado' if ok else 'no aplicado'}."
            elif typ == "section_columns":
                expected = spec.get("expected", "1")
                ok = False
                try:
                    sections = getattr(doc, "Sections", None)
                    sec = sections(sections.Count) if (sections is not None and sections.Count) else None
                    ps = sec.PageSetup if sec is not None else doc.PageSetup
                    ok = str(ps.TextColumns.Count) == str(expected)
                except Exception:
                    try:
                        ok = str(doc.PageSetup.TextColumns.Count) == str(expected)
                    except Exception:
                        ok = False
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Columnas: {expected if ok else '?'} (esperado {expected})."
            elif typ == "paragraph_indent":
                expected = spec.get("expected", "")
                ok = False
                actual_seen = None
                try:
                    target_pt = float(expected) / 20.0
                except Exception:
                    target_pt = None
                if target_pt is not None:
                    for p in doc.Paragraphs:
                        try:
                            li = p.ParagraphFormat.LeftIndent
                            actual_seen = li
                            if li is None:
                                continue
                            if abs(float(li) - target_pt) < 2:
                                ok = True
                                break
                        except Exception:
                            continue
                # Fallback locale/unit-independent: inspect paragraph/document XML
                # for an indent of the expected twips (w:left="360").
                if not ok:
                    try:
                        import re
                        needles = [str(expected)]
                        for p in doc.Paragraphs:
                            xml = str(getattr(p.Range, "XML", "") or "")
                            if ('w:left="' + str(expected) + '"') in xml:
                                ok = True
                                break
                        if not ok:
                            for attr in ("WordOpenXML",):
                                xml = str(getattr(doc, attr, "") or "")
                                if ('w:left="' + str(expected) + '"') in xml:
                                    ok = True
                                    break
                        if not ok:
                            xml = str(getattr(doc.Content, "XML", "") or "")
                            if ('w:left="' + str(expected) + '"') in xml:
                                ok = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Sangría verificada (esperado {expected} twips / {float(expected)/20:.1f} pt / {float(expected)/1440*2.54:.3f} cm; leído {actual_seen})."
            elif typ == "paragraph_spacing":
                expected = spec.get("expected", "")
                ok = False
                actual_seen = None
                try:
                    target_pt = float(expected) / 20.0
                except Exception:
                    target_pt = None
                if target_pt is not None:
                    for p in doc.Paragraphs:
                        try:
                            sa = p.ParagraphFormat.SpaceAfter
                            actual_seen = sa
                            if sa is None:
                                continue
                            if abs(float(sa) - target_pt) < 2:
                                ok = True
                                break
                        except Exception:
                            continue
                if not ok:
                    try:
                        for p in doc.Paragraphs:
                            xml = str(getattr(p.Range, "XML", "") or "")
                            if ('w:after="' + str(expected) + '"') in xml:
                                ok = True
                                break
                        if not ok:
                            xml = str(getattr(doc, "WordOpenXML", "") or "")
                            if ('w:after="' + str(expected) + '"') in xml:
                                ok = True
                        if not ok:
                            xml = str(getattr(doc.Content, "XML", "") or "")
                            if ('w:after="' + str(expected) + '"') in xml:
                                ok = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if ok else CriterionStatus.FAIL
                ev = f"Espaciado verificado (esperado {expected} twips / {float(expected)/20:.1f} pt / {float(expected)/1440*2.54:.3f} cm; leído {actual_seen})."
            elif typ == "merged_cell_present":
                expected = spec.get("expected", True)
                ok = False
                try:
                    for t in doc.Tables:
                        for row in t.Rows:
                            for cell in row.Cells:
                                # merged cells share same Range start; heuristic: count cells < rows*cols
                                pass
                    # heuristic: if any table has fewer cells than rows*cols
                    for t in doc.Tables:
                        if t.Rows.Count * t.Columns.Count > sum(1 for _ in t.Range.Cells):
                            ok = True
                except Exception:
                    pass
                if not ok:
                    try:
                        xml = str(getattr(doc, "WordOpenXML", "") or "")
                        if "gridSpan" in xml or "vMerge" in xml:
                            ok = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if ok == bool(expected) else CriterionStatus.FAIL
                ev = "Celda combinada detectada." if ok else "No se detectó celda combinada."
            elif typ == "repeating_header":
                expected = spec.get("expected", True)
                ok = False
                try:
                    for t in doc.Tables:
                        if t.Rows(1).HeadingFormat:
                            ok = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if ok == bool(expected) else CriterionStatus.FAIL
                ev = "Encabezado de tabla repetido." if ok else "No se repite encabezado de tabla."
            elif typ == "section_break_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    present = doc.Sections.Count > 1
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Salto de seccion detectado." if present else "No se detecto salto de seccion."
            elif typ == "section_break":
                expected = str(spec.get("expected", "nextPage")).lower()
                present = False
                try:
                    present = doc.Sections.Count > 1
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = f"Salto de seccion '{expected}' detectado." if present else f"No se detecto salto '{expected}'."
            elif typ == "page_break_present" or typ == "page_break":
                expected = spec.get("expected", True)
                present = False
                try:
                    text = getattr(doc.Content, "Text", "") or ""
                    if "\x0c" in text:
                        present = True
                    else:
                        # Fallback: search manual page break character.
                        import win32com.client  # noqa: F401
                        rng = doc.Content
                        find = rng.Find
                        find.Text = "^m"
                        find.Forward = True
                        find.Wrap = 1
                        present = bool(find.Execute())
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Salto de página detectado." if present else "No se detectó salto de página."
            elif typ == "shape_count":
                expected = spec.get("expected", 1)
                count = 0
                try:
                    count = doc.Shapes.Count
                except Exception:
                    count = 0
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Formas: {count} (esperado >= {expected})."
            elif typ == "smartart_present":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    for sh in doc.Shapes:
                        if getattr(sh, "HasSmartArt", False):
                            present = True
                            break
                    if not present:
                        for sh in doc.InlineShapes:
                            if getattr(sh, "HasSmartArt", False):
                                present = True
                                break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "SmartArt insertado." if present else "No se detectó SmartArt."
            elif typ == "model3d_present":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    for sh in doc.Shapes:
                        if "3D" in str(getattr(sh, "Name", "")).upper():
                            present = True
                            break
                    if not present:
                        for sh in doc.InlineShapes:
                            if "3D" in str(getattr(sh, "AlternativeText", "")).upper():
                                present = True
                                break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Modelo 3D insertado." if present else "No se detectó modelo 3D."
            elif typ == "image_background_removed":
                expected = bool(spec.get("expected", True))
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Imagen de fondo eliminable: no verificable desde el archivo."
            elif typ == "object_position":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    page_w = float(doc.PageSetup.PageWidth)
                    page_h = float(doc.PageSetup.PageHeight)
                    for sh in doc.Shapes:
                        cx = float(sh.Left) + float(sh.Width) / 2.0
                        cy = float(sh.Top) + float(sh.Height) / 2.0
                        if expected in ("top-right", "top_right", "topright"):
                            if cx > page_w / 2.0 and cy < page_h / 2.0:
                                present = True
                                break
                        elif expected in ("top-left", "top_left", "topleft"):
                            if cx < page_w / 2.0 and cy < page_h / 2.0:
                                present = True
                                break
                        else:
                            present = True
                            break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Objeto posicionado correctamente." if present else "El objeto no está en la posición esperada."
            elif typ == "text_wrap":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    for sh in doc.Shapes:
                        wt = getattr(sh.WrapFormat, "Type", None)
                        if expected == "tight" and wt == 1:
                            present = True
                            break
                        if expected and wt is not None:
                            present = True
                            break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Ajuste de texto configurado." if present else "No se detectó ajuste de texto."
            elif typ == "comment_resolved":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    for c in doc.Comments:
                        if getattr(c, "Resolved", False):
                            present = True
                            break
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Comentario resuelto." if present else "No se detectó comentario resuelto."
            elif typ == "revisions_reviewed":
                expected = spec.get("expected", True)
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Cambios revisados no verificable desde el archivo."
            elif typ == "track_changes_locked":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    present = bool(getattr(doc, "TrackRevisions", False))
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Seguimiento bloqueado." if present else "No se detectó bloqueo de seguimiento."
            elif typ == "internal_hyperlink":
                expected = spec.get("expected", True)
                present = False
                try:
                    for hl in doc.Hyperlinks:
                        if getattr(hl, "SubAddress", None):
                            present = True
                            break
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Hipervinculo interno detectado." if present else "No se detecto hipervinculo interno."
            elif typ == "formatting_marks_visible":
                # Las marcas de formato son una configuracion de la aplicacion
                # (vista), no del documento; Word COM no expone de forma fiable
                # el estado de la ventana del usuario (instancia separada), por
                # lo que se considera no automatizable.
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Marcas de formato: verificacion visual/manual (no automatizable)."
            elif typ == "alternative_format":
                expected = spec.get("expected", True)
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Formato alternativo no verificable."
            elif typ == "print_settings":
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Configuracion de impresion no verificable desde el archivo."
            elif typ == "remove_personal_info":
                expected = spec.get("expected", True)
                present_info = False
                try:
                    for prop in ("Title", "Subject", "Author", "Keywords", "Comments", "Category"):
                        try:
                            v = doc.BuiltInDocumentProperties(prop).Value
                            if v and str(v).strip():
                                present_info = True
                                break
                        except Exception:
                            pass
                except Exception:
                    pass
                removed = not present_info
                status = CriterionStatus.PASS if removed == bool(expected) else CriterionStatus.FAIL
                ev = "Propiedades personales eliminadas." if removed else "Aun hay propiedades personales."
            elif typ == "text_effects":
                expected = spec.get("expected", "shadow")
                present = False
                try:
                    # 1) Sombra legada (boton "Sombra del texto")
                    present = bool(getattr(doc.Content.Font, "Shadow", False)) or bool(getattr(doc.Content.Font, "Outline", False))
                except Exception:
                    present = False
                # 2) Sombras avanzadas (exterior/interior/perspectiva) no activan
                #    Font.Shadow; se detectan inspeccionando el XML del documento.
                if not present:
                    try:
                        xml = ""
                        try:
                            xml = str(getattr(doc.Content, "WordOpenXML", "") or "")
                        except Exception:
                            xml = ""
                        if not xml:
                            try:
                                xml = str(getattr(doc, "WordOpenXML", "") or "")
                            except Exception:
                                xml = ""
                        low = xml.lower()
                        if "shadow" in low or "effectlst" in low or "<w:effect" in low:
                            present = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Efecto de texto detectado." if present else "No se detecto efecto de texto."
            elif typ == "format_painter_used":
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Pincel de formato no verificable desde el archivo."
            elif typ == "footnote_props":
                expected = str(spec.get("expected", "")).lower()
                present = False
                try:
                    ns = None
                    try:
                        ns = doc.Footnotes
                    except Exception:
                        try:
                            ns = doc.Sections(1).Footnotes
                        except Exception:
                            ns = None
                    style = getattr(ns, "NumberStyle", None) if ns is not None else None
                    # WdNoteNumberStyle: 0=Arabic, 1=UppercaseRoman, 2=LowercaseRoman
                    if expected in ("roman", "romanos"):
                        present = style in (1, 2)
                    else:
                        present = style is not None
                except Exception:
                    present = False
                # Fallback: inspeccionar el XML (footnotes.xml / settings) por "roman"
                if not present and expected in ("roman", "romanos"):
                    try:
                        xml = ""
                        try:
                            xml = str(getattr(doc, "WordOpenXML", "") or "")
                        except Exception:
                            xml = ""
                        low = xml.lower()
                        if "roman" in low or "lowerroman" in low or "upperroman" in low:
                            present = True
                    except Exception:
                        pass
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Propiedades de nota configuradas." if present else "No se detectaron propiedades de nota."
            elif typ == "citation_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    present = doc.Fields.Count > 0
                except Exception:
                    pass
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Cita detectada." if present else "No se detecto cita."
            elif typ == "bibliography_present":
                expected = spec.get("expected", True)
                present = False
                try:
                    # 1) Campos de Word con codigo BIBLIOGRAPHY
                    try:
                        for f in doc.Fields:
                            try:
                                if "BIBLIOGRAPHY" in str(getattr(f, "Code", "")).upper():
                                    present = True
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass
                    # 2) Texto visible que contenga "Bibliography" (resultado del campo)
                    if not present:
                        try:
                            if "bibliography" in (doc.Content.Text or "").lower():
                                present = True
                        except Exception:
                            pass
                    # 3) Fallback: inspeccionar el XML del documento
                    if not present:
                        try:
                            xml = str(getattr(doc, "WordOpenXML", "") or "")
                            if "bibliography" in xml.lower():
                                present = True
                        except Exception:
                            pass
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present == bool(expected) else CriterionStatus.FAIL
                ev = "Bibliografia detectada." if present else "No se detecto bibliografia."
            elif typ == "toc_custom":
                expected = str(spec.get("expected", ""))
                present = False
                try:
                    present = doc.TablesOfContents.Count > 0
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "TOC personalizado insertado." if present else "No se detectó TOC personalizado."
            elif typ == "cell_margins":
                expected_twips = int(spec.get("expected", 144))
                expected_pt = expected_twips / 20.0
                present = False
                try:
                    if doc.Tables.Count > 0:
                        t = doc.Tables(1)
                        pads = [float(t.TopPadding), float(t.BottomPadding), float(t.LeftPadding), float(t.RightPadding)]
                        target_pts = (expected_twips / 20.0, 2.834645669)
                        present = all(any(abs(p - tp) < 0.5 for tp in target_pts) for p in pads)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Márgenes de celda configurados." if present else "Los márgenes de celda no están configurados."
            elif typ == "table_sorted":
                expected = str(spec.get("expected", "asc")).lower()
                present = False
                try:
                    if doc.Tables.Count > 0:
                        t = doc.Tables(1)
                        values = []
                        for r in range(1, t.Rows.Count + 1):
                            try:
                                txt = t.Cell(r, 1).Range.Text
                                txt = (txt or "").replace("\r", "").replace("\x07", "").strip()
                                values.append(txt)
                            except Exception:
                                values.append("")
                        data = values[1:] if len(values) > 1 else values
                        if expected.startswith("desc"):
                            present = data == sorted(data, reverse=True)
                        else:
                            present = data == sorted(data)
                except Exception:
                    present = False
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = "Tabla ordenada correctamente." if present else "La tabla no está ordenada por la primera columna."
            elif typ == "table_autofit":
                expected = bool(spec.get("expected", True))
                present = False
                try:
                    if doc.Tables.Count > 0:
                        t = doc.Tables(1)
                        present = bool(getattr(t, "AllowAutoFit", False)) or getattr(t, "AutoFitBehavior", None) == 2
                except Exception:
                    present = False
                status = CriterionStatus.PASS if (present == expected) else CriterionStatus.FAIL
                ev = "Ajuste de tabla configurado." if present else "No se detectó ajuste de tabla."
            elif typ == "list_levels":
                expected = spec.get("expected", 2)
                count = 0
                try:
                    levels = set()
                    for p in doc.Paragraphs:
                        try:
                            levels.add(p.Range.ListFormat.ListLevelNumber)
                        except Exception:
                            pass
                    count = len(levels) if levels else 1
                except Exception:
                    pass
                status = CriterionStatus.PASS if count >= expected else CriterionStatus.FAIL
                ev = f"Niveles de lista: {count} (esperado >= {expected})."
            elif typ == "list_restart":
                expected = bool(spec.get("expected", True))
                status = CriterionStatus.NOT_VERIFIABLE
                ev = "Reinicio de lista no verificable desde el archivo."
            elif typ == "list_start_value":
                expected = int(spec.get("expected", 1))
                start_val = None
                try:
                    for p in doc.Paragraphs:
                        lf = p.Range.ListFormat
                        if getattr(lf, "ListType", 0) not in (0, None):
                            try:
                                start_val = int(lf.ListValue)
                            except Exception:
                                try:
                                    start_val = int(lf.ListTemplate.ListLevels(1).StartAt)
                                except Exception:
                                    start_val = None
                            break
                except Exception:
                    start_val = None
                present = (start_val == expected)
                status = CriterionStatus.PASS if present else CriterionStatus.FAIL
                ev = f"Valor inicial de lista: {start_val} (esperado {expected})." if start_val is not None else "No se pudo determinar el valor inicial de la lista."
            else:
                status = CriterionStatus.ERROR
                ev = f"Tipo de criterio desconocido: {typ}."
            criteria.append(CriterionResult(cid, status, ev))
        return criteria

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
            from virmetrix.paths import resource_root

            base = resource_root() / "exercises" / exercise.exercise_id
            meta_path = base / "metadata.json"
            if meta_path.is_file():
                return json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}