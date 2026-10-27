from virmetrix.infrastructure.catalog.loader import load_exercise


def test_exercise_metadata_has_target_fields():
    ex = load_exercise("word-format-001")
    assert ex.target_phrase == "Informe de resultados"
    assert ex.expected_text == "El presente Informe de resultados resume los hallazgos del ejercicio."
    assert ex.criteria == ["C1", "C2", "C3", "C4", "C5"]