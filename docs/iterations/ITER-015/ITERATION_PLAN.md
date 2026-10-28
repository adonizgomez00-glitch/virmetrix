# ITERATION_PLAN — VirMetrix ITER-015 (Gate QA final)

**Iteración:** ITER-015  
**Fase SAQI:** 4/11 — QA y verificación  
**Fecha:** 2026-10-27  
**Modo SAQI:** riguroso.

## Objetivo

Cerrar el MVP con un gate de QA reproducible, comparable en flujo a herramientas de práctica MOS (abrir tarea, editar en Word, evaluar, retroalimentar), sin copiar contenido propietario.

## Alcance

- Ejecutar suite completa (53 tests) y cobertura.
- Validar bundle PyInstaller en Linux (arranque + catálogo + evaluación).
- Checklist de accesibilidad (teclado, nombres accesibles, contraste básico).
- Revisión de seguridad (ZIP, rutas, sin macros).
- Revisión de licencias y avisos legales.
- Documentación final (README, WINDOWS_VALIDATION, RELEASE_CANDIDATE_APPROVAL).
- Evidencia en QA_RESULTS.md y SESSION.md.

## Criterios de aceptación

- AC-01: 53 tests pass, 0 fail.
- AC-02: bundle arranca y evalúa en Linux.
- AC-03: sin defectos críticos/altos abiertos.
- AC-04: documentación de instalación Windows completa.
- AC-05: aprobación de release candidate registrada.

## Estado

En progreso.