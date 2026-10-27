# ITERATION_PLAN — VirMetrix ITER-003

**Iteración:** ITER-003  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** catálogo de ejercicios + criterios declarativos  
**Modo SAQI:** estándar.

## Objetivo

Añadir un catálogo de ejercicios originales y un modelo declarativo de criterios (JSON) para que el evaluador portable sea genérico y escalable, manteniendo compatibilidad con `word-format-001`.

## Alcance

- `exercises/catalog.json` con 4 ejercicios.
- Metadatos con `criteria_specs` declarativos.
- Evaluador portable genérico (`_evaluate_specs`) con tipos: `target_occurrences`, `run_property`, `neighbor_not_formatted`, `document_text_equals`.
- Evaluador COM genérico (`_evaluate_specs_com`).
- UI con selector de ejercicio desde el catálogo.
- Pruebas parametrizadas del catálogo (4 pass).

## Criterios de aceptación

- AC-01: catálogo lista ≥4 ejercicios.
- AC-02: cada ejercicio correcto obtiene puntuación completa en modo portable.
- AC-03: criterios declarativos cubren C1–C5.
- AC-04: UI permite seleccionar ejercicio y crear copia.
- AC-05: COM evalúa ejercicios declarativos en Windows (pendiente validación).

## Estado

Implementado en Linux: 38 tests pass. Pendiente validación COM de nuevos ejercicios en Windows.