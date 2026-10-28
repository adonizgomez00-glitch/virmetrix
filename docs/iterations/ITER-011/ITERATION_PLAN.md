# ITERATION_PLAN — VirMetrix ITER-011

**Iteración:** ITER-011  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** colaboración (Dominio 6)  
**Modo SAQI:** estándar.

## Objetivo

Añadir ejercicios y evaluación para comentarios y control de cambios.

## Alcance

- `collab-001`: comentario con texto.
- `collab-002`: cambios registrados.
- Criterios declarativos: `comment_count`, `comment_contains`, `tracked_changes_present`.
- Evaluador portable y COM generalizados.
- Pruebas unitarias (2).

## Criterios de aceptación

- AC-01: comentario correcto → puntuación completa (portable).
- AC-02: cambios registrados → puntuación completa (portable).
- AC-03: COM evalúa colaboración en Windows (pendiente).
- AC-04: catálogo incluye los nuevos ejercicios.

## Estado

Implementado en Linux: 53 tests pass. Pendiente validación COM en Windows.  
Cobertura: los seis dominios de MO-100 ahora tienen al menos un ejercicio.