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

## ITER-002 — Evidencia parcial (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `python3 -m pytest -q`
- Resultado: 33 passed, 0 failed.
- Cambios: persistencia SQLite (AttemptStore), evaluación en hilo (QThread worker), adaptador Word COM con evaluación real por COM (no ejecutable en Linux), metadatos tipados, seguridad ZIP multiplataforma.
- Word COM: no ejecutado (Linux) — registrado como no ejecutado.
- UI smoke: pass.
- Defectos abiertos: 0.

## ITER-002 — Validación Windows (2026-10-27)

- Entorno: Windows (VM), Python 3.x, Word v16.0, PySide6 6.11.2, pytest 9.1.1.
- Comando: `pytest -q` → **33 passed, 0 failed**.
- COM: disponible (v16.0); sin proceso residual tras corrección `DispatchEx` + `Quit`.
- Referencia (sin formato): 3/5 — C1 pass, C2 fail, C3 fail, C4 pass, C5 pass.
- Documento correcto: 5/5 — todos pass.
- Documento incorrecto: 3/5 — C2 fail, C3 fail, resto pass.
- UI: sin bloqueo; historial persistente (4 intentos tras reabrir).
- Defecto corregido: `is_available()` dejaba proceso Word residual; corregido con `DispatchEx` + `Quit` + `CoInitialize`/`CoUninitialize`.
- App en ejecución: PID 5108, ventana “VirMetrix - word-format-001”.

## ITER-003 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **38 passed, 0 failed**.
- Catálogo: 4 ejercicios (`word-format-001..004`).
- Criterios declarativos: soportados en portable y COM.
- Pruebas parametrizadas: 4 ejercicios correctos → puntuación completa.
- COM de nuevos ejercicios: no ejecutado (Linux); pendiente validación Windows.

## ITER-004 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **42 passed, 0 failed**.
- Sesión: modo práctica/examen, temporizador 300 s, pausa/reanudación, finalización idempotente.
- Navegación anterior/siguiente entre ejercicios.
- UI: sin bloqueo (evaluación en QThread).
- COM: pendiente validación en Windows.

## ITER-005 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, PyInstaller 6.22.3.
- Comando: `pyinstaller --clean --noconfirm virmetrix.spec` → build exitoso.
- Bundle: `dist/VirMetrix/VirMetrix` arranca sin error (QT_QPA_PLATFORM=offscreen).
- Catálogo: carga desde el bundle (`_internal/exercises/catalog.json`).
- Pruebas: `pytest -q` → 42 passed, 0 failed.
- Windows: pendiente build y validación de instalador .exe.

## ITER-006 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **42 passed, 0 failed**.
- Accesibilidad: nombres accesibles, tooltips, botón por defecto, diálogo "Acerca de".
- UI smoke: pass.
- Lector de pantalla: pendiente validación en Windows.

## ITER-007 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **44 passed, 0 failed**.
- Nuevos ejercicios: `table-001`, `list-001`.
- Criterios: tablas (dimensiones y texto de celdas) y listas (conteo).
- COM: pendiente validación en Windows.

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

## ITER-002 � Validaci�n Windows + Word COM (2026-10-27)

- Entorno: Windows 10 Pro 10.0.19045, Python 3.12.10, PySide6 6.11.2, python-docx 1.2.0, lxml 5.2.1, pywin32 312, pytest 9.1.1.
- Microsoft Word: v16.0 (COM disponible).

### 1. Pruebas autom�ticas
- Comando: `$env:PYTHONPATH="src"; pytest -q` (ejecutado desde ra�z del proyecto)
- Resultado: **33 passed, 0 failed** (duraci�n 6.18 s).

### 2. Disponibilidad de Word COM
- Comando: `python -c "from virmetrix.infrastructure.evaluators.word_com import WordComAdapter; print(WordComAdapter().is_available())"`
- Resultado: `(True, 'Word disponible (v16.0).')`.
- Proceso residual: ninguno (tras correcci�n de `is_available` con `DispatchEx` + `Quit`).

### 3. Evaluaci�n COM sobre documento sin formato (referencia)
- Documento: `exercises/word-format-001/reference.docx` (copia de trabajo).
- Resultado: `word_com 3 / 5`
  - C1: pass � Ocurrencias de 'Informe de resultados': 1 (esperado 1).
  - C2: fail � Negrita efectiva en objetivo: False.
  - C3: fail � Tama�o (pt) en objetivo: 11.0.
  - C4: pass � Texto vecino conserva formato de referencia.
  - C5: pass � Texto del documento coincide con el esperado.
- Excepci�n: ninguna. Proceso Word residual: ninguno.

### 4. Evaluaci�n COM sobre documento correcto
- Documento: `~/.virmetrix/work/correct.docx` (frase en negrita, 16 pt).
- Resultado: `word_com 5 / 5`
  - C1..C5: pass.

### 5. Evaluaci�n COM sobre documento incorrecto
- Documento: `~/.virmetrix/work/incorrect.docx` (sin negrita, 12 pt).
- Resultado: `word_com 3 / 5`
  - C1: pass, C2: fail, C3: fail, C4: pass, C5: pass.

### 6. UI (no bloqueo + historial)
- Comando: `python -m virmetrix.ui.main_window`.
- Modo `word_com`: evaluaci�n completada sin congelar la ventana; resultado mostrado.
- Historial: "Historial guardado. Intentos totales: N" aparece tras la evaluaci�n.
- Persistencia SQLite verificada: el conteo se mantiene tras cerrar y reabrir (4 intentos registrados).
- Proceso Word residual tras evaluaci�n UI: ninguno.

### Defectos
- Defecto corregido: `WordComAdapter.is_available()` usaba `Dispatch` y no llamaba `Quit`, dejando un proceso WINWORD residual. Corregido usando `DispatchEx` + `Quit`. Estado: cerrado.
- Sin defectos abiertos tras la validaci�n.

## ITER-003 � Validaci�n Windows + Word COM (2026-10-27)

- Entorno: Windows 10 Pro 10.0.19045, Python 3.12.10, Word v16.0, PySide6 6.11.2, pytest 9.1.1.
- Pruebas autom�ticas: 38 passed, 0 failed.

### Validaci�n COM por ejercicio

Referencia (sin formato):
- word-format-001: 3/5 (C1 pass, C2 fail, C3 fail, C4 pass, C5 pass)
- word-format-002: 3/5 (C1 pass, C2 fail, C3 fail, C4 pass, C5 pass)
- word-format-003: 3/5 (C1 pass, C2 fail, C3 fail, C4 pass, C5 pass)
- word-format-004: 4/5 (C1 pass, C2 fail, C3 pass, C4 pass, C5 pass)

Documento correcto (formato aplicado):
- word-format-001 (bold 16 pt): 5/5
- word-format-002 (italic 14 pt): 5/5
- word-format-003 (bold 18 pt): 5/5
- word-format-004 (italic 11 pt): 5/5

Documento incorrecto (sin formato, 12 pt):
- word-format-001: 3/5 (C2 fail, C3 fail)
- word-format-002: 3/5 (C2 fail, C3 fail)
- word-format-003: 3/5 (C2 fail, C3 fail)
- word-format-004: 3/5 (C2 fail, C3 fail)

### Defectos
- Ninguno nuevo. La detecci�n de cursiva (run_property italic) funciona correctamente en COM.
- Sin procesos Word residuales tras las evaluaciones.
- Ajuste post-validación: `word-format-004` tamaño esperado 11→12 pt para evitar falso positivo por fuente por defecto de Word (la validación original se realizó con 11 pt).
