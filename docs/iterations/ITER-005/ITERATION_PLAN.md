# ITERATION_PLAN — VirMetrix ITER-005

**Iteración:** ITER-005  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** empaquetado Windows  
**Modo SAQI:** estándar.

## Objetivo

Generar un ejecutable empaquetado (PyInstaller) que funcione sin entorno de desarrollo, incluyendo ejercicios y recursos, y documentar el instalador.

## Alcance

- `main.py` como punto de entrada.
- `virmetrix.spec` con datos `exercises/`.
- `src/virmetrix/paths.py` para rutas en modo frozen.
- Scripts de build para Linux y Windows.
- README con instrucciones de empaquetado.
- Validación de arranque del bundle en Linux.

## Criterios de aceptación

- AC-01: el bundle arranca sin error en Linux (QT_QPA_PLATFORM=offscreen).
- AC-02: el catálogo carga desde el bundle.
- AC-03: la evaluación portable funciona en el bundle.
- AC-04: historial SQLite se guarda fuera del bundle (home del usuario).
- AC-05: instrucciones de build para Windows documentadas.

## Estado

Build Linux validado. Pendiente build y validación en Windows (instalador .exe).