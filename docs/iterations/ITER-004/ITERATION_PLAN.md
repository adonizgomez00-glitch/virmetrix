# ITERATION_PLAN — VirMetrix ITER-004

**Iteración:** ITER-004  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** temporizador y modo examen  
**Modo SAQI:** estándar.

## Objetivo

Añadir modo práctica/examen, temporizador cronometrado y navegación entre ejercicios, sin bloquear la UI ni duplicar resultados.

## Alcance

- `ExamSession` con estados idle/running/paused/finished.
- Modo práctica (sin tiempo) y examen (300 s).
- Pausa/reanudación y finalización idempotente.
- Navegación anterior/siguiente.
- Temporizador con QTimer.
- Pruebas unitarias de sesión (4).

## Criterios de aceptación

- AC-01: iniciar/pausar/reanudar/finalizar sesión.
- AC-02: temporizador decrementa y finaliza automáticamente en 0.
- AC-03: no se duplican resultados al finalizar dos veces.
- AC-04: navegación entre ejercicios actualiza la copia de trabajo.
- AC-05: UI no se congela durante evaluación.

## Estado

Implementado en Linux: 42 tests pass. Pendiente validación en Windows.