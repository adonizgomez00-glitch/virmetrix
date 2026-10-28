# ITERATION_PLAN — VirMetrix ITER-016

**Iteración:** ITER-016  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** indicador de progreso y cancelación  
**Modo SAQI:** estándar.

## Objetivo

Añadir barra de progreso indeterminada y botón de cancelación durante la evaluación.

## Alcance

- QProgressBar indeterminado visible durante evaluación.
- Botón "Cancelar" que termina el hilo y restaura la UI.
- Pruebas existentes siguen pasando.

## Criterios de aceptación

- AC-01: la barra aparece al evaluar y desaparece al terminar.
- AC-02: el botón cancelar restaura la UI.
- AC-03: 65 tests pass.

## Estado

Implementado en Linux: 65 tests pass. Pendiente validación en Windows.