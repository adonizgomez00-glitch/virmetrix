# ITERATION_PLAN — VirMetrix ITER-007

**Iteración:** ITER-007  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** tablas y listas (Dominio 3)  
**Modo SAQI:** estándar.

## Objetivo

Añadir ejercicios y evaluación portable/COM para tablas y listas, primer dominio nuevo tras texto.

## Alcance

- `table-001`: tabla 3x2 con texto en celdas.
- `list-001`: lista con viñetas de 3 elementos.
- Criterios declarativos: `table_count`, `table_dimensions`, `table_cell_text`, `list_items_count`.
- Evaluador portable y COM generalizados.
- Pruebas unitarias (2).

## Criterios de aceptación

- AC-01: tabla correcta → puntuación completa (portable).
- AC-02: lista correcta → puntuación completa (portable).
- AC-03: COM evalúa tablas y listas en Windows (pendiente).
- AC-04: catálogo incluye los nuevos ejercicios.

## Estado

Implementado en Linux: 44 tests pass. Pendiente validación COM en Windows.