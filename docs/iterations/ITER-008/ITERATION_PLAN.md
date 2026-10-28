# ITERATION_PLAN — VirMetrix ITER-008

**Iteración:** ITER-008  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** referencias (Dominio 4)  
**Modo SAQI:** estándar.

## Objetivo

Añadir ejercicios y evaluación para notas al pie y tabla de contenido.

## Alcance

- `refs-001`: nota al pie con texto esperado.
- `refs-002`: tabla de contenido.
- Criterios declarativos: `footnote_count`, `footnote_contains`, `toc_present`.
- Evaluador portable y COM generalizados.
- Pruebas unitarias (2).

## Criterios de aceptación

- AC-01: nota al pie correcta → puntuación completa (portable).
- AC-02: TOC presente → puntuación completa (portable).
- AC-03: COM evalúa referencias en Windows (pendiente).
- AC-04: catálogo incluye los nuevos ejercicios.

## Estado

Implementado en Linux: 46 tests pass. Pendiente validación COM en Windows.