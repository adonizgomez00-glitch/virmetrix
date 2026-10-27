"""Genera el DOCX de referencia para word-format-001 de forma reproducible."""

import hashlib
import json
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
META = json.loads((ROOT / "exercises/word-format-001/metadata.json").read_text(encoding="utf-8"))
OUT = ROOT / "exercises/word-format-001/reference.docx"

TARGET = META["target_phrase"]
BEFORE = "El presente "
AFTER = " resume los hallazgos del ejercicio."

doc = Document()
p = doc.add_paragraph()
p.add_run(BEFORE)
p.add_run(TARGET)  # sin formato especial: el estudiante aplica negrita y 16pt
p.add_run(AFTER)
doc.save(OUT)

h = hashlib.sha256(OUT.read_bytes()).hexdigest()
META["reference_sha256"] = h
(ROOT / "exercises/word-format-001/metadata.json").write_text(
    json.dumps(META, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
)
print("Generado:", OUT)
print("SHA256:", h)
