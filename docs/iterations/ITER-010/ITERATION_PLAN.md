# ITERATION_PLAN — VirMetrix ITER-010

**Iteración:** ITER-010  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** administrar documentos (Dominio 1)  
**Modo SAQI:** estándar.

## Objetivo

Añadir ejercicios y evaluación para encabezados, orientación de página y propiedades del documento.

## Alcance

- `doc-manage-001`: encabezado con texto.
- `doc-manage-002`: orientación horizontal.
- `doc-manage-003`: propiedad Title.
- Criterios declarativos: `header_contains`, `page_orientation`, `doc_property`.
- Evaluador portable y COM generalizados.
- Pruebas unitarias (3).

## Criterios de aceptación

- AC-01: encabezado correcto → puntuación completa (portable).
- AC-02: orientación correcta → puntuación completa (portable).
- AC-03: propiedad correcta → puntuación completa (portable).
- AC-04: COM evalúa en Windows (pendiente).
- AC-05: catálogo incluye los nuevos ejercicios.

## Estado

Implementado en Linux: 51 tests pass. Pendiente validación COM en Windows.