# ITERATION_PLAN — VirMetrix ITER-021

**Iteración:** ITER-021  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** interfaz moderna y llamativa  
**Modo SAQI:** estándar.

## Objetivo

Reemplazar la apariencia genérica por una interfaz moderna tipo panel de examen
(GMetrix), legible sobre Word y con feedback visual inmediato.

## Alcance

- Cabecera con degradado, marca y temporizador grande.
- Chips de progreso: 1 · Copia → 2 · Editar → 3 · Evaluar.
- Tarjetas con bordes redondeados y sombra suave.
- Tarjeta de tarea con acento azul lateral.
- Puntuación grande con color según rendimiento (verde/ámbar/rojo).
- Resultados por criterio con iconos y color.
- Cancelar visible solo durante la evaluación.
- Scroll con desplazamiento automático al resultado.
- Corrección de ancho: los combos ya no desbordan la ventana.

## Criterios de aceptación

- AC-01: la UI cabe en 460 px sin recortes horizontales.
- AC-02: el paso actual se refleja en los chips.
- AC-03: la puntuación y los criterios se distinguen por color.
- AC-04: `pytest -q` → 68 passed; cobertura ≥ 80%.

## Evidencia Linux

- `pytest -q` → 68 passed.
- `pytest --cov=src/virmetrix` → 80.52%.
- Capturas offscreen revisadas: parte superior, resultado y scroll inferior.

## Estado

Implementado y validado en Linux y Windows.

## Validación Windows (2026-10-29)

- Word COM: evaluación completada sin bloqueo; score 1/2 con colores.
- Historial guardado; barra de progreso oculta al finalizar; ventana topmost.
- Evidencia: `e2e_evidence/ui_iter021_windows.png`.
- 68 tests pass en Windows.

## Accesibilidad (preparación Narrador)

- Orden de tabulación explícito en `_setup_tab_order()`.
- Chips de paso con `setAccessibleName` y estado accesible (actual/completado/pendiente).
- Puntuación y su detalle con nombres accesibles.
- Checklist actualizada en `ACCESSIBILITY_CHECKLIST.md`.
