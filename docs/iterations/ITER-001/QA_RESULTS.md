# QA Results

**Estado:** no se han ejecutado pruebas de software; solo se ha creado documentación inicial.

## Resumen

| Área | Estado | Evidencia |
|---|---|---|
| Revisión de requisitos | En curso | Documentación inicial |
| Build | No ejecutado | No hay implementación |
| Unitarias | No ejecutadas | No hay código |
| Integración | No ejecutada | No hay código |
| E2E Linux | No ejecutada | No hay UI |
| E2E Windows + Word | No ejecutada | No hay adaptador COM |
| QA adversarial | No ejecutado | Plan en `TEST_STRATEGY.md` |

## Registro de defectos

No se han registrado defectos de implementación porque aún no existe implementación. Los riesgos de diseño se registran en `PROJECT_STATE.md` y `SECURITY.md`.

## Protocolo de actualización

Para cada iteración incluir:

- Fecha y commit o identificador de versión.
- Entorno y versiones relevantes.
- Comandos ejecutados.
- Conteo de pruebas aprobadas, fallidas y no ejecutadas.
- Evidencia resumida.
- Defectos encontrados y su estado.
- Regresión ejecutada después de los arreglos.

Nunca reemplazar “no ejecutado” por “aprobado” por inferencia.

## ITER-001 — Evidencia Fase 4 (Build) — 2026-10-27

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2, python-docx, lxml 5.2.1.
- Comando: `python3 -m pytest -q`
- Resultado: 15 passed, 0 failed.
- Desglose: 10 pruebas unitarias (evaluador portable: correcto, negrita ausente, tamaño incorrecto, objetivo ausente, fragmentado, texto cambiado, heredado no verificable, vecino con fuga, múltiples ocurrencias) + 5 adversariales (traversal, ZIP inválido, falta document.xml, archivo vacío, inexistente).
- Cobertura/mutation: no ejecutadas (MVP-First; diferidas a macro-fase QA final).
- Word COM: no ejecutado (entorno Linux; se registra como no ejecutado, no como aprobado).
- E2E UI: no ejecutado aún.

## ITER-001 — Evidencia Fase 5 (Unit & Integration) — 2026-10-27

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, coverage 7.17.0.
- Comando: `python3 -m pytest --cov=src/virmetrix --cov-report=term -q`
- Resultado: 22 passed, 0 failed.
- Cobertura: 92.36% (umbral ≥80% alcanzado). Excluidos de cobertura: UI (requiere display) y rama COM de éxito (solo Windows).
- Contract tests: `EvaluationAdapter` implementado por ambos adaptadores (runtime_checkable).
- Defectos: 0 nuevos.

## ITER-001 — Evidencia Fase 6/7 (Functional & Adversarial) — 2026-10-27

- Funcionales (L3): 4 pass (referencia sin formato, documento correcto, reproducibilidad, modo Word no disponible).
- UI smoke (xvfb): 1 pass.
- Adversariales: 7 pass (traversal, ZIP inválido, falta document.xml, archivo vacío, inexistente, filename Unicode, not_verifiable no contado).
- Defectos abiertos: 0.
- Word COM: no ejecutado (Linux) — registrado como no ejecutado.
