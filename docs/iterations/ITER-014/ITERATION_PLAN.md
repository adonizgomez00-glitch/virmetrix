# ITERATION_PLAN — VirMetrix ITER-014

**Iteración:** ITER-014  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** abrir documento en Word  
**Modo SAQI:** estándar.

## Objetivo

Permitir abrir la copia de trabajo en Microsoft Word (o programa predeterminado) para que el usuario pueda realizar las tareas manualmente.

## Alcance

- `EvaluationOrchestrator.open_document()` multiplataforma.
- Botón "Abrir documento" en la UI.
- Apertura automática tras crear la copia de trabajo.
- Pruebas existentes siguen pasando.

## Criterios de aceptación

- AC-01: crear copia abre el documento en Word en Windows.
- AC-02: botón "Abrir documento" funciona.
- AC-03: en Linux usa xdg-open (si está disponible).
- AC-04: pruebas automatizadas pasan.

## Estado

Implementado en Linux: 53 tests pass. Pendiente validación en Windows.