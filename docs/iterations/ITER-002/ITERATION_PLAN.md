# ITERATION_PLAN — VirMetrix ITER-002

**Iteración:** ITER-002
**Fase SAQI:** 4 — Build (parcial) + QA
**Fecha:** 2026-10-27
**Vertical slice:** word-format-001 (extendido)
**Modo SAQI:** estándar; riguroso en COM y parsing.

## Objetivo

Añadir persistencia local de intentos, evaluación no bloqueante en UI, adaptador Word COM con evaluación real y corregir deuda técnica (metadatos tipados, seguridad ZIP multiplataforma).

## Alcance

- Persistencia SQLite: `AttemptStore` con historial por ejercicio.
- Threading: `EvaluationWorker` + QThread; botón deshabilitado durante evaluación.
- Word COM: evaluación real mediante `Word.Application` (bold, tamaño, texto, vecinos); cleanup; no matar procesos ajenos.
- Deuda: `ExerciseMetadata.target_phrase`/`expected_text`; normalización de rutas en ZIP.

## Criterios de aceptación

- AC-P1: cada evaluación se persiste y se puede listar.
- AC-P2: la UI no se bloquea durante la evaluación.
- AC-P3: Word COM devuelve estados C1–C5 reales en Windows (no ejecutado en Linux).
- AC-P4: metadatos tipados cargados desde JSON.
- AC-P5: ZIP con separadores Windows detectado como traversal.

## Estado

Implementado en Linux (33 tests pass). COM pendiente de validación en Windows.