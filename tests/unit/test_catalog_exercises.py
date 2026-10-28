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


@pytest.mark.parametrize("exinfo", [e for e in list_exercises() if e["exercise_id"].startswith("word-format")], ids=lambda e: e["exercise_id"])
def test_catalog_exercise_correct_passes(exinfo, tmp_path):
    ex = load_exercise(exinfo["exercise_id"])
    data = _correct_bytes(ex)
    path = tmp_path / f"{ex.exercise_id}.docx"
    path.write_bytes(data)
    res = ev.evaluate(str(path), ex)
    assert res.status == "completed"
    assert res.score_earned == res.score_possible
    assert all(c.status.value == "pass" for c in res.criteria)


def test_table_001_correct(tmp_path):
    from docx import Document
    ex = load_exercise("table-001")
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Producto"
    table.cell(0, 1).text = "Cantidad"
    table.cell(1, 0).text = "Teclado"
    table.cell(1, 1).text = "5"
    buf = io.BytesIO()
    doc.save(buf)
    path = tmp_path / "table.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_list_001_correct(tmp_path):
    from docx import Document
    ex = load_exercise("list-001")
    doc = Document()
    doc.add_paragraph("Lista de compras")
    for t in ["Leche", "Pan", "Huevos"]:
        doc.add_paragraph(t, style="List Bullet")
    buf = io.BytesIO()
    doc.save(buf)
    path = tmp_path / "list.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_refs_001_footnote(tmp_path):
    import io, zipfile
    ex = load_exercise("refs-001")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body><w:p><w:r><w:t>Texto principal</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body>
</w:document>'''
    footnotes = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>Nota al pie de ejemplo.</w:t></w:r></w:p></w:footnote>
</w:footnotes>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/footnotes.xml", footnotes)
    path = tmp_path / "fn.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_graphics_001_alt_text(tmp_path):
    import io, zipfile
    ex = load_exercise("graphics-001")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:wp="{WP}">
  <w:body><w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture" descr="Logo de la empresa"/></wp:inline></w:drawing></w:r></w:p></w:body>
</w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
    path = tmp_path / "g1.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_graphics_002_textbox(tmp_path):
    import io, zipfile
    ex = load_exercise("graphics-002")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body><w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>Nota lateral</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p></w:body>
</w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
    path = tmp_path / "g2.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_doc_manage_001_header(tmp_path):
    import io, zipfile
    ex = load_exercise("doc-manage-001")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body><w:sectPr><w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr></w:body>
</w:document>'''
    header = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Encabezado principal</w:t></w:r></w:p></w:hdr>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/header1.xml", header)
    path = tmp_path / "h.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_doc_manage_002_orientation(tmp_path):
    import io, zipfile
    ex = load_exercise("doc-manage-002")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr></w:body>
</w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
    path = tmp_path / "o.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_doc_manage_003_property(tmp_path):
    import io, zipfile
    ex = load_exercise("doc-manage-003")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    core = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Informe final</dc:title>
</cp:coreProperties>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("docProps/core.xml", core)
    path = tmp_path / "p.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_collab_001_comment(tmp_path):
    import io, zipfile
    ex = load_exercise("collab-001")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'''
    comments = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{W}">
  <w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Revisar este párrafo.</w:t></w:r></w:p></w:comment>
</w:comments>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/comments.xml", comments)
    path = tmp_path / "c.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_collab_002_tracked(tmp_path):
    import io, zipfile
    ex = load_exercise("collab-002")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:p><w:ins w:id="1" w:author="A" w:date="2020-01-01T00:00:00Z"><w:r><w:t>insertado</w:t></w:r></w:ins></w:p></w:body></w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
    path = tmp_path / "c2.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_refs_002_toc(tmp_path):
    import io, zipfile
    ex = load_exercise("refs-002")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body><w:p><w:fldSimple w:instr="TOC \\\\o \\\\h \\\\z \\\\u"><w:r><w:t>TOC</w:t></w:r></w:fldSimple></w:p></w:body>
</w:document>'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
    path = tmp_path / "toc.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_list_exercises_not_empty():
    assert len(list_exercises()) >= 29