import io, zipfile
from pathlib import Path

import pytest

from virmetrix.infrastructure.catalog.loader import load_exercise
from virmetrix.infrastructure.evaluators.portable import PortableDocxEvaluator

ev = PortableDocxEvaluator()
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _zip(parts):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def test_doc_manage_004_footer(tmp_path):
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:sectPr><w:footerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr></w:body></w:document>'''
    footer = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Pie de página principal</w:t></w:r></w:p></w:ftr>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/_rels/document.xml.rels": doc_rels, "word/document.xml": document, "word/footer1.xml": footer})
    path = tmp_path / "f.docx"; path.write_bytes(data)
    ex = load_exercise("doc-manage-004")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_doc_manage_005_property(tmp_path):
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
  <cp:subject>Borrador</cp:subject>
</cp:coreProperties>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document, "docProps/core.xml": core})
    path = tmp_path / "s.docx"; path.write_bytes(data)
    ex = load_exercise("doc-manage-005")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_text_para_003_indent(tmp_path):
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
<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document})
    path = tmp_path / "i.docx"; path.write_bytes(data)
    ex = load_exercise("text-para-003")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_text_para_004_style(tmp_path):
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
<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document})
    path = tmp_path / "s.docx"; path.write_bytes(data)
    ex = load_exercise("text-para-004")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_table_003_merged(tmp_path):
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
<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document})
    path = tmp_path / "mc.docx"; path.write_bytes(data)
    ex = load_exercise("table-003")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_list_003_count(tmp_path):
    from docx import Document
    ex = load_exercise("list-003")
    doc = Document()
    for x in ["a", "b", "c", "d"]:
        doc.add_paragraph(x, style="List Bullet")
    buf = io.BytesIO()
    doc.save(buf)
    path = tmp_path / "l3.docx"
    path.write_bytes(buf.getvalue())
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_refs_005_endnote(tmp_path):
    import io, zipfile
    ex = load_exercise("refs-005")
    Wn = W
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{Wn}"><w:body><w:p><w:r><w:t>x</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:body></w:document>'''
    endnotes = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="{Wn}">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:id="1"><w:p><w:r><w:t>Nota final avanzada.</w:t></w:r></w:p></w:endnote>
</w:endnotes>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/_rels/document.xml.rels": doc_rels, "word/document.xml": document, "word/endnotes.xml": endnotes})
    path = tmp_path / "en.docx"; path.write_bytes(data)
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_refs_006_footnotes(tmp_path):
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
<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>x</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p></w:body></w:document>'''
    footnotes = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>a</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>b</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="3"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:footnote>
</w:footnotes>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/_rels/document.xml.rels": doc_rels, "word/document.xml": document, "word/footnotes.xml": footnotes})
    path = tmp_path / "fn3.docx"; path.write_bytes(data)
    ex = load_exercise("refs-006")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_graphics_005_drawings(tmp_path):
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
  <w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="P" descr="x"/></wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:docPr id="2" name="P" descr="y"/></wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:docPr id="3" name="P" descr="z"/></wp:inline></w:drawing></w:r></w:p>
  </w:body>
</w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document})
    path = tmp_path / "g5.docx"; path.write_bytes(data)
    ex = load_exercise("graphics-005")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_graphics_006_textbox(tmp_path):
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
  <w:body><w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>Nota lateral avanzada</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p></w:body>
</w:document>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/document.xml": document})
    path = tmp_path / "g6.docx"; path.write_bytes(data)
    ex = load_exercise("graphics-006")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_collab_005_comments(tmp_path):
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
  <w:comment w:id="0" w:author="A"><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1" w:author="A"><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:comment>
  <w:comment w:id="2" w:author="A"><w:p><w:r><w:t>c3</w:t></w:r></w:p></w:comment>
</w:comments>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/_rels/document.xml.rels": doc_rels, "word/document.xml": document, "word/comments.xml": comments})
    path = tmp_path / "c5.docx"; path.write_bytes(data)
    ex = load_exercise("collab-005")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible


def test_collab_006_comment(tmp_path):
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
  <w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Aprobar cambios</w:t></w:r></w:p></w:comment>
</w:comments>'''
    data = _zip({"[Content_Types].xml": content_types, "_rels/.rels": rels, "word/_rels/document.xml.rels": doc_rels, "word/document.xml": document, "word/comments.xml": comments})
    path = tmp_path / "c6.docx"; path.write_bytes(data)
    ex = load_exercise("collab-006")
    res = ev.evaluate(str(path), ex)
    assert res.score_earned == res.score_possible