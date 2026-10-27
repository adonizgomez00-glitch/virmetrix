# ITERATION_PLAN — VirMetrix

**Iteración:** ITER-001  
**Fase SAQI:** 1 — Plan  
**Fecha:** 2026-10-27  
**Estado:** Borrador para aprobación del humano (PO/Tech Lead)  
**Vertical slice:** `word-format-001` (ver `VERTICAL_SLICE_001.md`)  
**Modo SAQI:** Estándar con macro-fase QA final (MVP-First para Build; ver §5)

---

## 1. Objetivo de la iteración

Demostrar el recorrido completo de un ejercicio de certificación MOS Word (MO-100) desde seleccionar un ejercicio hasta recibir una evaluación verificable, tanto sin Word (`docx_portable`) como con Word real en Windows (`word_com`), con evidencia reproducible y QA adversarial.

No se incluye en esta iteración: temporizador de examen completo, login/cuenta, nube, sincronización, editor propio, catálogo amplio, persistencia de historial completa (diferida a ITER-002).

## 2. Alcance y criterios de aceptación

Se adoptan íntegramente los **AC-01..AC-12** de `VERTICAL_SLICE_001.md`:

- AC-01 Inicio offline
- AC-02 Copia de trabajo (no modificar original)
- AC-03 Modo portable (estados C1–C5, puntuación, evidencia)
- AC-04 Detección explícita de Word (sin bloqueo ni selección silenciosa)
- AC-05 Modo Word real (`evaluator=word_com`)
- AC-06 Corrección positiva
- AC-07 Corrección negativa
- AC-08 Texto fragmentado en varios runs
- AC-09 Archivo inválido (error controlado, sin contenido activo)
- AC-10 Evidencia de resultados (ID/versión ejercicio, adaptador, fecha, estados, versión de reglas)
- AC-11 Reproducibilidad (resultados idénticos salvo fecha)
- AC-12 Regresión automatizada (unit + COM en runner Windows o “no ejecutado”, nunca “aprobado”)

Criterios de evaluación del ejercicio: **C1** integridad del contenido, **C2** negrita, **C3** tamaño 16 pt (directo + herencia soportada), **C4** no afectar texto vecino, **C5** no alterar plantilla.

Estados por criterio: `pass`, `fail`, `not_verifiable`, `error`. `not_verifiable` no suma puntos.

## 3. Historias (backlog de la iteración)

| ID | Historia | Prioridad | Estimación | Dependencias |
|---|---|---|---|---|
| H1 | Metadatos del ejercicio `word-format-001` (id, versión, idioma, competencia, instrucciones, criterios declarativos) + generador reproducible del DOCX de referencia | P0 | M | — |
| H2 | Evaluador portable OOXML: C1 (contenido), C2 (negrita), C3 (tamaño 16 pt con directo+herencia soportada), C4 (vecinos), C5 (plantilla intacta) | P0 | L | H1 |
| H3 | Modelo de resultado estructurado + reglas de puntuación deterministas (tratamiento de `not_verifiable`) | P0 | S | H2 |
| H4 | UI mínima (PySide6): instrucciones, selección de modo, apertura de copia de trabajo, evaluación, resultados/evidencia | P1 | M | H1, H3 |
| H5 | Adaptador Word COM (Windows) tras interfaz común + detección de capacidad + timeouts/limpieza | P1 | L | H3 |
| H6 | Pruebas automatizadas: unitarias (correcto, negrita ausente, tamaño incorrecto, objetivo ausente, fragmentado, herencia, corrupto, tamaño excesivo) + adversariales (ZIP, XML, traversal, doble clic, concurrencia) | P0 | M | H2, H3 |
| H7 | Documentación: `QA_RESULTS.md` con evidencia, `ARCHITECTURE.md`/`CONTEXT.md` actualizados, `SESSION.md` checkpoint | P2 | S | H2–H6 |

**Leyenda estimación:** S = pequeño, M = mediano, L = grande.

**Diferido a ITER-002:** persistencia SQLite del historial de intentos, temporizador de examen, empaquetado PyInstaller, catálogo amplio.

## 4. Definition of Done (iteración)

### Modo MVP-First (Fase 4)
- Cada historia H1–H6 está **funcional y verificable**: sus AC/tests funcionales pasan; `typecheck`/`py_compile` OK; lint OK. No se exige cobertura ≥80% ni mutation ≥70% como salida de Fase 4.
- 0 errores de lint/tipado en el código nuevo.

### Macro-fase QA final (antes del candidato de release)
- Pruebas unitarias/integración requeridas pasan (H6).
- Flujo E2E portable pasa en Linux y Windows.
- Flujo COM pasa en entorno Windows con Word o se registra “no ejecutado” (nunca “aprobado”).
- QA adversarial completado; defectos críticos/altos cerrados.
- Gate de release (`TEST_STRATEGY.md`) cumplido.
- Documentación viva actualizada (H7).

### Criterio de salida Fase 1
- Plan aprobado por humano (PO/Tech Lead).
- ACs testables definidos (AC-01..AC-12).
- DoD clara (este documento).
- Riesgos mitigados o aceptados (§6).

## 5. Modo SAQI y política de entrega

- **Modo adaptativo:** estándar para el núcleo de evaluación (riesgo medio-alto por seguridad de archivos); riguroso en el adaptador COM y en el parsing OOXML.
- **MVP-First** aplicado a Fase 4 (Build): se prioriza funcionalidad validada con tests funcionales; la QA estricta (cobertura/mutation/regresión/adversarial/seguridad) se condensa en la macro-fase final antes del release.
- La Fase 7 (Break / QA adversarial) **no se omite** antes de una versión candidata.

## 6. Riesgos y dependencias

| Riesgo | Probabilidad | Impacto | Mitigación |
|---|---|---|---|
| El evaluador portable no puede resolver herencia de estilos de forma fiable | Media | Medio | Marcar `not_verifiable` con motivo; no aprobar por suposición (AC-06/AC-07) |
| Word COM no disponible en el entorno de desarrollo Linux | Alta | Medio | Aislar adaptador; detección explícita; runner Windows para COM |
| Vulnerabilidades de parsing OOXML (ZIP bomb, XXE, traversal) | Media | Alto | Límites de tamaño/entradas, validación de rutas, sin ejecución de contenido activo |
| `python-docx` insuficiente para inspección de runs/estilos | Media | Medio | Evaluador OOXML propio con `zipfile` + `lxml`/`xml.etree` |
| Inconsistencia entre modos portable y Word | Media | Medio | Comparación documentada; `not_verifiable` explícito |
| UI bloqueante durante evaluación | Baja | Medio | Evaluación en hilo/worker; pruebas de no congelamiento |

**Dependencias técnicas:** Python 3.x, PySide6, `python-docx`, `lxml` (a validar), `pytest`, `pywin32` (solo Windows, importación bajo demanda).

## 7. Skills seleccionadas (Fase 3 — preliminar)

Nivel A (obligatorias): `saqi-lean-code`, `saqi-spec-arch`, `saqi-qa-harness`, `saqi-context-manager`, `saqi-testing`, `saqi-harness`, `saqi-metrics`, `saqi-token-optimizer`, `saqi-adaptive-mode`, `saqi-vertical-slice`, `saqi-qa-cycles`, `saqi-skills-learning` + las globales `A-sdd`, `A-testing`, `A-secure-coding`, `D-git-workflow` (si disponibles).

Nivel B/C según stack: Python/PySide6 no tienen skill específico; se aplican convenciones de `saqi-lean-code` y `A-secure-coding`. Se revalidará en Fase 3 con versiones exactas.

## 8. Métricas de iteración (SAQI-007)

- ACs cumplidos / ACs totales (objetivo: 12/12 portable; COM según entorno).
- Tests: total, pass rate, no ejecutados (COM).
- Defectos por severidad (P0/P1/P2/P3) y por categoría.
- `not_verifiable` count y su impacto en puntuación.
- Tokens/tiempo por historia (opcional, si se instrumenta).

## 9. Aprobación

- [ ] Aprobado por humano (PO/Tech Lead): ______ Fecha: ______
- [ ] Riesgos aceptados: ______

> Tras la aprobación, proceder a Fase 2 (Architect): ADRs, contratos de puertos, diagramas C4 y actualización de `ARCHITECTURE.md`.
