# ITERATION_PLAN — VirMetrix ITER-017

**Iteración:** ITER-017  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** rediseño de UI estilo gmetrix  
**Modo SAQI:** estándar.

## Objetivo

Mejorar la interfaz para que el flujo sea claro (pasos 1-2-3), con panel de instrucciones y resultados, tamaño de ventana adecuado y navegación visible.

## Alcance

- Ventana redimensionable (960x720, mínimo 820x620).
- Panel izquierdo: controles y pasos numerados.
- Panel derecho: temporizador y resultados.
- Etiqueta de paso actual (Ejercicio X de N).
- Pruebas existentes pasan.

## Criterios de aceptación

- AC-01: la UI muestra instrucciones y resultados lado a lado.
- AC-02: el usuario ve claramente el paso actual y navegación.
- AC-03: 65 tests pass.

## Estado

Implementado en Linux: 65 tests pass. Pendiente validación en Windows.