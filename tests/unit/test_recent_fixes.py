import io
import json
import zipfile
from pathlib import Path

import pytest

from virmetrix.infrastructure.catalog.loader import load_exercise
from virmetrix.infrastructure.evaluators.portable import PortableDocxEvaluator

ev = PortableDocxEvaluator()


def _rewrite_docx(src: Path, dst: Path, edits):
    """edits: dict filename -> transform callable(bytes)->bytes"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename in edits:
                data = edits[info.filename](data)
            zout.writestr(info, data)


def test_d_table_003_sorted_by_price_passes(tmp_path):
    from docx import Document
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.rows[0].cells[0].text = "Producto"
    t.rows[0].cells[1].text = "Precio"
    for a, b in [("Uva", "8"), ("Manzana", "5"), ("Pera", "3")]:
        cells = t.add_row().cells
        cells[0].text = a
        cells[1].text = b
    p = tmp_path / "sorted.docx"
    doc.save(p)
    ex = load_exercise("d-table-003")
    res = ev.evaluate(str(p), ex)
    assert res.score_earned == res.score_possible
    assert all(c.status.value == "pass" for c in res.criteria)


def test_d_table_003_wrong_column_fails(tmp_path):
    from docx import Document
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.rows[0].cells[0].text = "Producto"
    t.rows[0].cells[1].text = "Precio"
    for a, b in [("Manzana", "5"), ("Pera", "3"), ("Uva", "8")]:
        cells = t.add_row().cells
        cells[0].text = a
        cells[1].text = b
    p = tmp_path / "unsorted.docx"
    doc.save(p)
    ex = load_exercise("d-table-003")
    res = ev.evaluate(str(p), ex)
    assert res.score_earned != res.score_possible


def test_d_collab_001_modern_comment_resolved(tmp_path):
    src = Path("exercises/d-collab-001/reference.docx")
    dst = tmp_path / "resolved.docx"

    def add_extended(data):
        text = data.decode("utf-8")
        # comments.xml: mark legacy done too (harmless)
        text = text.replace("<w:comment w:id=\"0\"", "<w:comment w:id=\"0\" w:done=\"1\"")
        return text.encode("utf-8")

    def add_ext_file(data):
        return data

    # Build commentsExtended.xml content
    ext_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w15:commentEx w15:paraId="00000001" w15:done="1"/>'
        '</w15:commentsEx>'
    )

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/comments.xml":
                data = data.replace(b"<w:comment w:id=\"0\"", b"<w:comment w:id=\"0\"")
            zout.writestr(info, data)
        # add extended part
        zout.writestr("word/commentsExtended.xml", ext_xml.encode("utf-8"))

    ex = load_exercise("d-collab-001")
    res = ev.evaluate(str(dst), ex)
    assert res.score_earned == res.score_possible


def test_d_refs_001_toc_custom_levels(tmp_path):
    src = Path("exercises/d-refs-001/reference.docx")
    dst = tmp_path / "toc.docx"

    def inject_toc(data):
        text = data.decode("utf-8")
        # Insert a TOC field code into the first paragraph if not present
        if "TOC" not in text:
            text = text.replace("</w:body>", '<w:p><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r></w:p></w:body>')
        else:
            if '1-2' not in text:
                text = text.replace("TOC", 'TOC \\\\o "1-2"')
        return text.encode("utf-8")

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = inject_toc(data)
            zout.writestr(info, data)

    ex = load_exercise("d-refs-001")
    res = ev.evaluate(str(dst), ex)
    assert res.score_earned == res.score_possible