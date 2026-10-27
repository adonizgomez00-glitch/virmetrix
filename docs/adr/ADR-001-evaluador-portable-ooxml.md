# ADR-001: Evaluador portable basado en inspección OOXML

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

El vertical slice `word-format-001` debe evaluar documentos DOCX sin Microsoft Word. `python-docx` abstrae parte del XML y oculta detalles de runs, herencia de estilos y propiedades directas. Los criterios C2–C4 exigen inspección fiable de negrita, tamaño y vecinos, incluyendo texto fragmentado en varios runs.

## Decisión

Implementar el evaluador portable mediante inspección directa del paquete OOXML (`zipfile` + parser XML), sin depender únicamente de `python-docx`. `python-docx` podrá usarse como ayuda para generación de documentos de referencia, pero la evaluación se basa en el XML.

## Consecuencias

- Mayor control y fidelidad para los criterios OOXML.
- Mayor complejidad de parsing y validación de seguridad (ZIP, XML).
- Riesgo de no resolver herencia de estilos: se marca `not_verifiable` con motivo.

## Alternativas

- Usar solo `python-docx`: descartado por pérdida de detalle de runs/estilos.
- Usar LibreOffice como motor de renderizado: descartado (no es sustituto implícito de Word).

## Referencias

- `VERTICAL_SLICE_001.md` (C1–C5)
- `ARCHITECTURE.md` (Evaluador portable DOCX)
