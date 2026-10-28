# ITERATION_PLAN — VirMetrix ITER-012

**Iteración:** ITER-012  
**Fase SAQI:** 4 — Build  
**Fecha:** 2026-10-27  
**Vertical slice:** ampliación de cobertura por dominio (2 ejercicios por dominio)  
**Modo SAQI:** estándar.

## Objetivo

Añadir 12 ejercicios adicionales (2 por dominio) con criterios evaluables en portable y COM.

## Alcance

Nuevos ejercicios:
- doc-manage-004, doc-manage-005
- text-para-001, text-para-002
- table-002, list-002
- refs-003, refs-004
- graphics-003, graphics-004
- collab-003, collab-004

Total catálogo: 28 ejercicios.

## Criterios de aceptación

- AC-01: todos los nuevos ejercicios pasan en portable (sintéticos validados).
- AC-02: COM evalúa en Windows (pendiente).
- AC-03: pruebas automatizadas pasan.

## Estado

Implementado en Linux: 53 tests pass. Pendiente validación COM en Windows.