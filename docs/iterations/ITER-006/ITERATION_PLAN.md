# ITERATION_PLAN — VirMetrix ITER-006

**Iteración:** ITER-006  
**Fase SAQI:** 4 — Build / QA  
**Fecha:** 2026-10-27  
**Vertical slice:** accesibilidad y pulido  
**Modo SAQI:** estándar.

## Objetivo

Mejorar la accesibilidad (navegación por teclado, nombres accesibles, tooltips, mensajes claros) y el pulido de la UI.

## Alcance

- Nombres accesibles (`setAccessibleName`) en controles.
- Tooltips descriptivos.
- Botón por defecto (Evaluar) y atajos básicos.
- Diálogo "Acerca de" con versión y licencia.
- Mensajes de error claros y no ambiguos.
- Pruebas de humo de UI actualizadas.

## Criterios de aceptación

- AC-01: todos los controles interactivos tienen nombre accesible.
- AC-02: la UI es navegable por teclado (Tab/Enter).
- AC-03: los botones tienen tooltips.
- AC-04: existe diálogo "Acerca de".
- AC-05: las pruebas de UI siguen pasando.

## Estado

Implementado en Linux. Pendiente validación con lector de pantalla en Windows.