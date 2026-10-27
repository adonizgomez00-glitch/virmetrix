import io
import tempfile
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt

from virmetrix.infrastructure.catalog.loader import list_exercises, load_exercise
from virmetrix.infrastructure.evaluators.portable import PortableDocxEvaluator

ev = PortableDocxEvaluator()


def _correct_bytes(ex):
    prop = None
    size = None
    for s in ex.criteria_specs:
        if s["type"] == "run_property" and s["property"] in ("bold", "italic"):
            prop = s["property"]
        if s["type"] == "run_property" and s["property"] == "size_pt":
            size = s["expected"]
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("El presente ")
    r = p.add_run("Informe de resultados")
    if prop == "bold":
        r.bold = True
    if prop == "italic":
        r.italic = True
    r.font.size = Pt(size)
    p.add_run(" resume los hallazgos del ejercicio.")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


@pytest.mark.parametrize("exinfo", list_exercises(), ids=lambda e: e["exercise_id"])
def test_catalog_exercise_correct_passes(exinfo, tmp_path):
    ex = load_exercise(exinfo["exercise_id"])
    data = _correct_bytes(ex)
    path = tmp_path / f"{ex.exercise_id}.docx"
    path.write_bytes(data)
    res = ev.evaluate(str(path), ex)
    assert res.status == "completed"
    assert res.score_earned == res.score_possible
    assert all(c.status.value == "pass" for c in res.criteria)


def test_list_exercises_not_empty():
    assert len(list_exercises()) >= 4