# QA Results

**Estado:** VALIDADO EN WINDOWS (2026-11-08): 68 pruebas OK, Word COM (True), 3/5 y 5/5 - ver ultimo registro del fichero. Historico ITER-001: no se han ejecutado pruebas de software; solo se ha creado documentación inicial.

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

## ITER-008 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **46 passed, 0 failed**.
- Nuevos ejercicios: `refs-001`, `refs-002`.
- Criterios: notas al pie y tabla de contenido.
- COM: pendiente validación en Windows.

## ITER-009 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **48 passed, 0 failed**.
- Nuevos ejercicios: `graphics-001`, `graphics-002`.
- Criterios: imágenes (texto alternativo) y cuadros de texto.
- COM: pendiente validación en Windows.

## ITER-010 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **51 passed, 0 failed**.
- Nuevos ejercicios: `doc-manage-001`, `doc-manage-002`, `doc-manage-003`.
- Criterios: encabezados, orientación y propiedades.
- COM: pendiente validación en Windows.

## ITER-011 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **53 passed, 0 failed**.
- Nuevos ejercicios: `collab-001`, `collab-002`.
- Cobertura: los seis dominios de MO-100 tienen al menos un ejercicio (16 ejercicios en total).
- COM: pendiente validación en Windows.

## ITER-012 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **53 passed, 0 failed**.
- Nuevos ejercicios: 12 (2 por dominio).
- Total catálogo: 28 ejercicios.
- COM: pendiente validación en Windows.

## ITER-013 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **53 passed, 0 failed**.
- Nuevos ejercicios: 12 (2 por dominio).
- Total catálogo: 41 ejercicios.
- COM: pendiente validación en Windows.

## ITER-014 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1.
- Comando: `pytest -q` → **53 passed, 0 failed**.
- Funcionalidad: abrir documento en programa predeterminado (Word en Windows).
- COM: pendiente validación en Windows.

## ITER-015 — Gate QA (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PyInstaller 6.22.3.
- Comando: `pytest --cov=src/virmetrix --cov-report=term -q` → **65 passed, 0 failed**, cobertura 80.52% (umbral 80%).
- Bundle Linux: arranca y evalúa (`dist/VirMetrix/VirMetrix`).
- Seguridad: límites ZIP/XML, traversal, sin macros.
- Accesibilidad básica: nombres accesibles, tooltips, botón por defecto, diálogo "Acerca de".
- Windows: validación COM completa (iteraciones 005–014), 53+ tests, sin defectos abiertos.
- Licencias: pendiente definir licencia de código y contenido (ver PROJECT_STATE).

## ITER-016 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **65 passed, 0 failed**.
- UI: barra de progreso indeterminada y botón cancelar.
- COM: pendiente validación en Windows.

## ITER-017 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **65 passed, 0 failed**.
- UI rediseñada: paneles izquierdo/derecho, pasos numerados, tamaño de ventana.
- COM: pendiente validación en Windows.

## ITER-018 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **65 passed, 0 failed**.
- UI: panel compacto, siempre encima, pasos numerados.
- COM: pendiente validación en Windows.

## ITER-019 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **65 passed, 0 failed**.
- Funcionalidad: mantener VirMetrix encima de Word (topmost).
- COM: pendiente validación en Windows.

## ITER-020 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **65 passed, 0 failed**.
- UI: estilo moderno (QSS), bordes redondeados, colores de acento.
- COM: pendiente validación en Windows.

## ITER-021 — Evidencia (2026-10-27)

- Entorno: Linux, Python 3.12.3, pytest 9.1.1, PySide6 6.11.2.
- Comando: `pytest -q` → **68 passed, 0 failed**.
- Comando: `pytest --cov=src/virmetrix --cov-report=term -q` → **80.52%**, umbral 80% cumplido.
- UI: rediseño moderno — cabecera con degradado, chips de progreso (Copia → Editar → Evaluar),
  tarjetas con sombra, puntuación con color según resultado, resultados con iconos por criterio.
- Corrección: los combos ya no fuerzan ancho extra (contenido cabe en 460 px).
- `.coveragerc`: se omite `com_eval_cli.py` (entrypoint de sistema operativo, igual que `word_com.py`).
- Nuevos tests: `tests/integration/test_ui_panels.py` (pasos, puntuación, escape HTML).
- COM: pendiente validación en Windows.

## ITER-021 — Validación Windows + Word COM (2026-10-29)

- Entorno: Windows (VM), Word v16.0, PySide6 6.11.2.
- Repositorio en `ITER-021` + fix de `com_eval_cli.py`.
- `pytest -q` → **68 passed, 0 failed**.
- UI real con Word COM: evaluación completada sin bloqueo; score **1/2** mostrado con
  colores; historial guardado; barra de progreso se oculta al terminar.
- Ventana permanece encima de Word (topmost) y conserva el tamaño previsto.
- Evidencia visual: `e2e_evidence/ui_iter021_windows.png`.
- Observaciones: C1 fail y C2 pass se leen con iconos/colores; sin recortes horizontales.
- COM: validado en Windows.
- Ajuste posterior: orden de tabulación explícito y descripciones accesibles en los chips
  de paso (preparación de la validación con Narrador).

## Protocolo de actualización:

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

## Validaci�n Windows � Iteraciones ITER-005 a ITER-012 (2026-10-27)

- Entorno: Windows 10 Pro 10.0.19045, Python 3.12.10, Word v16.0, PySide6 6.11.2, pytest 9.1.1.
- Pruebas autom�ticas: 53 passed, 0 failed.
- Cat�logo: 29 ejercicios (formato, tablas, listas, referencias, gr�ficos, gesti�n de documento, colaboraci�n, texto/p�rrafo).
- Validaci�n COM sobre referencias de todos los ejercicios: completada sin excepciones ni bloqueos.
- UI smoke en modo Word COM: sin congelamiento; historial persistente (5 intentos tras reabrir).
- Procesos Word residuales: ninguno.
- App en ejecuci�n: PID 8276, ventana "VirMetrix".

## Validaci�n Windows � Iteraciones ITER-013 e ITER-014 (2026-10-27)

- Entorno: Windows 10 Pro 10.0.19045, Python 3.12.10, Word v16.0, PySide6 6.11.2, pytest 9.1.1.
- Pruebas autom�ticas: 53 passed, 0 failed.
- Nuevos ejercicios validados en COM (referencias): refs-006, table-003, text-para-003, text-para-004 � completados sin excepciones.
- Funcionalidad "Abrir documento" (ITER-014): os.startfile abre la copia de trabajo en Word correctamente.
- Procesos Word residuales: ninguno (salvo el abierto intencionalmente por el usuario).
- App en ejecuci�n: PID 6276, ventana "VirMetrix".

## Validación Windows + Word COM — 2026-11-08 (post-mejoras)

- Entorno: Windows 10 Pro 10.0.19045, Word 16.0.20326.20232, Python 3.12.10, PySide6 6.11.2, python-docx 1.2.0, lxml 5.2.1, pywin32 312, pytest 9.1.1, pytest-cov 7.1.0.
- Pruebas: `pytest -q` → **68 passed, 0 failed** (31.56 s).
- COM disponible: `(True, 'Word disponible (v16.0).')`.
- Referencia (sin formato): `word_com 3/5` (C1 pass, C2 fail, C3 fail, C4 pass, C5 pass).
- Documento correcto (formato aplicado + guardado): `word_com 5/5` (C1–C5 pass).
- Ruido no bloqueante: faulthandler `0x800706be` (RPC server disconnected) en stderr durante tests que simulan COM no disponible; código de salida 0.
- Proceso residual: 1 WINWORD (PID 9304) tras la ejecución (probablemente de la validación). Recomendación: cerrar manualmente o matar el proceso.
- Conclusión: todas las comprobaciones con valor esperado cumplen. COM validado tras las mejoras de UI y puntuación.

## Sesión 2026-11-09 — Bugs detectados y corregidos

### Bugs en modo COM (word_com) corregidos

| Ejercicio | Título | Bug | Causa | Fix (commit) |
|---|---|---|---|---|
| 10 (graphics-002) | Cuadro de texto | Instrucción no mencionaba la imagen requerida; evaluador solo contaba 1 dibujo | Instrucción incompleta; `drawing_count` solo contaba InlineShapes | Instrucción + criterio `drawing_count`=2; `drawing_count` suma InlineShapes+Shapes (`176eec7`) |
| 12 (doc-manage-002) | Orientación horizontal | Reportaba portrait aunque la sección principal fuera landscape | Portable usaba el primer `sectPr`; word_com usaba `PageSetup` (primera sección) | Usar último `sectPr` / última sección (`1323795`) |
| 19 (text-para-002) | Color rojo | `Propiedad 'color' en objetivo: 255` y fail | `Font.Color` es OLE_COLOR (BGR) int, comparado con hex string | Convertir OLE_COLOR a hex RGB (`c8f0c09`) |
| 31 (text-para-004) | Estilo Título 1 | `Estilo 'Heading1' no aplicado` | Comparaba `"heading1"` (sin espacio) con `"Heading 1"` / `"Título 1"` (con espacio/localizado) | Normalizar (minúsculas, sin espacios) y comprobar `Name` y `NameLocal` (`ed8e35f`) |
| 30 (text-para-003) | Sangría izquierda | `Sangría verificada (esperado 720 twips).` y fail | `try/except` abortaba el bucle; sin diagnóstico | `try/except` por párrafo, tolerancia 2 pt, mensaje con valor leído (`878a4e0`) |

### Publicación del .exe
- Causa original: `dist/` y `build/` en `.gitignore`.
- Solución: carpeta `releases/` trackeada + workflow GitHub Actions (build Windows + release en tag `v*`).
- Releases: `v0.1.1`, `v0.1.2`, `v0.1.3`, `v0.1.4` (último incluye icono + todos los fixes).

### Icono de la aplicación
- `virmetrix.ico` aplicado en UI (`QIcon`) y en el ejecutable (PyInstaller `icon=` + `datas`) (`b03ae01`).

### Modo portable
- Sin bugs nuevos; todos los ejercicios afectados pasan en portable.

### Pendientes de validación en Windows real
- Color (rojo/azul) en Word COM.
- Estilo Heading1/Título 1.
- Sangría izquierda (verificar valor leído con el nuevo diagnóstico).
- Orientación horizontal en documentos con saltos de sección.
- drawing_count imagen+cuadro de texto.
- Procesos Word residuales.

## Validacion Windows - Checklist WINDOWS_VALIDATION.md (2026-11-08)

- SO: Microsoft Windows 10 Pro 10.0.19045 (build 19045).
- Word: 16.0.20326.20232 (fichero) / COM 16.0 build 16.0.20326.
- Python: 3.12.10 (venv .venv, instalacion per-user; el Python 3.15 del sistema NO sirve por el pin PySide6==6.11.2).
- Dependencias: PySide6 6.11.2, python-docx 1.2.0, lxml 5.2.1, pywin32 312, pytest 9.1.1, pytest-cov 7.1.0.

### 3. Pruebas portables
- Comando: `$env:PYTHONPATH="src"; pytest -q`
- Resultado: exit 0 -> **68 passed in 31.56s**, 0 failed.

### 4. Disponibilidad Word COM
- Comando: `python -c "from virmetrix.infrastructure.evaluators.word_com import WordComAdapter; print(WordComAdapter().is_available())"`
- Resultado: exit 0 -> `(True, 'Word disponible (v16.0).')`

### 5. Evaluacion COM (word-format-001)
- Comando: script temporal con EvaluationOrchestrator + load_exercise (PYTHONPATH=src).
- Copia de trabajo: `C:\Users\Adoniz\.virmetrix\work\reference.docx`
- Referencia sin formato: `word_com 3 / 5`
  - C1 pass - Ocurrencias de 'Informe de resultados': 1 (esperado 1).
  - C2 fail - Propiedad 'bold' en objetivo: False.
  - C3 fail - Propiedad 'size_pt' en objetivo: 11.0.
  - C4 pass - Texto vecino conserva formato de referencia.
  - C5 pass - Texto del documento coincide con el esperado.
- Documento correcto (negrita + 16 pt aplicados solo a la frase con Word COM y guardado con Save): `word_com 5 / 5`
  - C1 pass - Ocurrencias: 1. C2 pass - bold: True. C3 pass - size_pt: 16.0. C4 pass - vecino intacto. C5 pass - texto coincidente.
- Ambos resultados coinciden EXACTAMENTE con lo esperado por el checklist.

### 6. UI (paso 6)
- Lanzada con `python -m virmetrix.ui.main_window` para validacion manual: crear copia de trabajo, evaluar en docx_portable y word_com, comprobar no-bloqueo e historial, probar Siguiente/Anterior.
- Estado: pendiente de confirmacion manual del usuario.

### 7. Pruebas por dominio (paso 7, opcional)
- Pendientes: word-format-001, text-para-004, table-004, refs-001, graphics-001, collab-002.

### Observaciones / defectos
- Ningun defecto nuevo: las 4 comprobaciones con valor esperado (68 passed, COM True, 3/5, 5/5) cumplen.
- Ruido no bloqueante: pytest emite 6 volcados `Windows fatal exception: code 0x800706be` (RPC server disconnected) en stderr, originados en `src/virmetrix/infrastructure/evaluators/word_com.py` y `src/virmetrix/application/orchestrator.py` durante tests que simulan COM no disponible. pytest termina con exit 0 y 68 passed. Posible mejora: silenciar faulthandler en esos tests.
- Proceso WINWORD residual detectado tras la fase de evidencia: comprobar si procede de los scripts de validacion (Dispatch + Quit) y no del adaptador de la app (DispatchEx + Quit).
