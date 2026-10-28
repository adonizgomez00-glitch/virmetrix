# ITERATION_PLAN — VirMetrix ITER-009

**Iteración:** ITER-009  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** elementos gráficos (Dominio 5)  
**Modo SAQI:** estándar.

## Objetivo

Añadir ejercicios y evaluación para imágenes (texto alternativo) y cuadros de texto.

## Alcance

- `graphics-001`: imagen con texto alternativo.
- `graphics-002`: cuadro de texto con contenido.
- Criterios declarativos: `drawing_count`, `alt_text_contains`, `textbox_contains`.
- Evaluador portable y COM generalizados.
- Pruebas unitarias (2).

## Criterios de aceptación

- AC-01: imagen con alt text correcto → puntuación completa (portable).
- AC-02: cuadro de texto correcto → puntuación completa (portable).
- AC-03: COM evalúa gráficos en Windows (pendiente).
- AC-04: catálogo incluye los nuevos ejercicios.

## Estado

Implementado en Linux: 48 tests pass. Pendiente validación COM en Windows.