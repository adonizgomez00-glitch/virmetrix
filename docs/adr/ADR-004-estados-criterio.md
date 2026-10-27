# ADR-004: Estados de criterio y regla de puntuación

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

La evaluación debe diferenciar criterios satisfechos, no satisfechos, no verificables y con error de ejecución. Contar `not_verifiable` como aprobado falsea la puntuación.

## Decisión

Estados permitidos por criterio: `pass`, `fail`, `not_verifiable`, `error`. La puntuación suma solo `pass`. `not_verifiable` y `error` no suman, y se reportan con motivo/evidencia. La regla de puntuación es determinista y versionada con el ejercicio.

## Consecuencias

- Puntuación honesta y reproducible (AC-11).
- Obliga a evidencia por criterio (AC-10).

## Referencias

- `VERTICAL_SLICE_001.md` (Contrato de estados, AC-06, AC-07, AC-10)
