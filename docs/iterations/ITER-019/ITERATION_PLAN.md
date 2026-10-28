# ITERATION_PLAN — VirMetrix ITER-019

**Iteración:** ITER-019  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** mantener VirMetrix encima de Word  
**Modo SAQI:** estándar.

## Objetivo

Evitar que Microsoft Word tape la ventana de VirMetrix durante la práctica.

## Alcance

- `WindowStaysOnTopHint` por defecto.
- `_bring_to_front()` con `raise_`, `activateWindow` y `SetWindowPos(HWND_TOPMOST)` en Windows.
- Llamada tras crear copia y abrir documento.
- Pruebas: 65 pass.

## Estado

Implementado en Linux: 65 tests pass. Pendiente validación en Windows.