
## Actualización — 2026-11-10 (v0.4.0 publicada)

- Evaluación D: 20 ejercicios nuevos similares a B, registrados en `catalog.json` (192 total: A 89, B 33, C 50, D 20).
- `reference.docx` D: corregidos b-collab-001/002 (comentario base) y d-list-001..003 (`numPr`/`ilvl`); metadata de `d-table-004` con `cell_margins` plano (57 twips).
- Auditoría A/B/C (192): sin referencias corruptas ni tipos no implementados; 2 issues corregidos (3 refs regeneradas: b-refs-001, b-refs-004, text-para-011; 7 criterios triviales marcados `not_verifiable` igual en portable y COM: formatting_marks_visible, print_settings, format_painter_used, image_background_removed, revisions_reviewed, list_restart; alternative_format ya lo estaba).
- Guía `EXERCISE_GENERATION_GUIDELINES.md`: ámbito A/B/C/D; documentada la excepción de `page_margins.expected` como dict intencional.
- `APP_VERSION = 0.4.0` y diálogo "Acerca de" coherente con el tag.
- Push a origin/main (6d81df2) + tag `v0.4.0`; workflow Build Windows genera el .exe para la release.
- Tests: 68 pass. Evidencia Windows pendiente: los 20 ejercicios D — validar COM(referencia≠max, correcto=max).
- Deuda: text-para-011 (`document_text_equals` ignora formato) sigue trivial; requiere rediseño del criterio (formato) o registro como verificación manual.
- Pendiente: revisar release v0.4.0 en GitHub (exe adjunto por CI).

## Actualización — 2026-10-29 (ITER-021 validada en Windows)

- Validación real en Windows con Word COM: evaluación sin bloqueo, score 1/2 con colores.
- Historial guardado y barra de progreso oculta al terminar; ventana topmost.
- Evidencia: `e2e_evidence/ui_iter021_windows.png`.
- 68 tests pass; cobertura 80.52%.
- Preparación accesibilidad: orden de tabulación explícito y chips con descripción de estado.
- Siguiente: validación con Narrador (ver `ACCESSIBILITY_CHECKLIST.md`).

## Actualización — 2026-10-27 (ITER-021 Build)

- UI modernizada: cabecera con degradado, chips de progreso, tarjetas con sombra.
- Puntuación con color (verde/ámbar/rojo) y criterios con iconos.
- Cancelar visible solo durante la evaluación; scroll automático al resultado.
- Corrección de ancho de combos (sin recortes en 460 px).
- Pruebas: 68 pass; cobertura 80.52%.
- Pendiente: validación visual en Windows (fuentes, sombras, topmost).

## Checkpoint — 2026-10-27 (fin de sesión)

- Repositorio sincronizado con origin/main (commit 8dd48e4).
- UI compacta estilo gmetrix, topmost, estilo moderno.
- COM en proceso separado (com_eval_cli.py) con ruta corregida.
- 65 tests pass, cobertura 80.52%.
- Pendiente: validación en Windows (topmost, COM sin bloqueo, evaluación completa).
- Próxima acción: validar en VM Windows y reportar resultados.

# SESSION — Checkpoint

**Fecha:** 2026-10-27  
**Proyecto:** VirMetrix  
**Fase SAQI estimada:** Fase 1 — Plan / definición de requisitos.

## Objetivo de esta sesión

Convertir las primeras decisiones del usuario en documentos de proyecto utilizables por agentes y desarrolladores.

## Decisiones del usuario

- Aplicación de escritorio integrada con Word real.
- Debe permitir practicar sin Word instalado.
- Windows es la plataforma principal; Linux será el entorno principal de desarrollo.

## Artefactos creados

- `README.md`
- `CONTEXT.md`
- `PROJECT_STATE.md`
- `ARCHITECTURE.md`
- `REQUIREMENTS.md`
- `WORD_COMPETENCIES.md`
- `ROADMAP.md`
- `TEST_STRATEGY.md`
- `QA_RESULTS.md`
- `SECURITY.md`
- `DECISIONS.md`
- `SESSION.md`

## Pendiente de resolver con el usuario

1. ¿Qué versión de MOS Word Associate quieres priorizar? Si no tienes preferencia, podemos usar la versión vigente de los objetivos públicos que confirmemos antes de construir el catálogo.
2. ¿Quieres que la aplicación se distribuya como instalador `.exe` para Windows, además de permitir ejecutar el código fuente en Linux?
3. ¿Quieres que el proyecto sea 100 % offline en su primera versión, sin cuentas, telemetría ni servicios web?

## Próxima tarea técnica recomendada

Crear el vertical slice `word-format-001`:
1. Documento DOCX original de práctica.
2. Especificación de criterios y puntuación.
3. Evaluador portable de OOXML.
4. Pruebas unitarias con DOCX sintéticos.
5. Adaptador COM mínimo en Windows.
6. Comparación de resultados y actualización de `QA_RESULTS.md`.

## Regla de reanudación

Antes de implementar, leer `CONTEXT.md`, `PROJECT_STATE.md`, `ARCHITECTURE.md` y `QA_RESULTS.md`. No asumir que un elemento de roadmap está completado sin evidencia.

## Decisiones confirmadas en esta iteración

- Examen objetivo: MOS Word Associate (Office 2019), MO-100, como referencia vigente según Microsoft Learn consultado el 2026-10-27.
- Distribución: instalador `.exe` para Windows y código fuente multiplataforma.
- Conectividad: MVP 100 % offline.

## Fuente oficial consultada

https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

La página publica seis dominios: administrar documentos (20–25 %); insertar/formatear texto, párrafos y secciones (20–25 %); tablas y listas (20–25 %); referencias (5–10 %); elementos gráficos (15–20 %); colaboración en documentos (5–10 %).

## Siguiente tarea técnica: vertical slice `word-format-001`

Implementar una tarea original de formato de texto, con evaluación portable y adaptador COM en Windows. Ver los criterios de aceptación concretos en `VERTICAL_SLICE_001.md`.

## Próximos pasos

1. Revisar el esquema oficial de aptitudes MO-100 para completar la trazabilidad.
2. Aprobar stack con un prototipo.
3. Implementar el vertical slice.
4. Ejecutar QA portable en Linux y Windows; después, validar COM en Windows con Word instalado.

## Decisiones confirmadas en esta iteración

- Nombre de la aplicación: **VirMetrix**.
- Stack inicial aprobado: Python, PySide6, `python-docx`, evaluador OOXML, `pywin32` en Windows, SQLite y pytest; PyInstaller como candidato para empaquetado Windows.
- SAQI es la metodología de desarrollo, no el nombre del producto.

## Próximo trabajo

Implementar `word-format-001` con este stack, empezando por el evaluador DOCX portable y sus pruebas. Validar empaquetado y automatización COM en Windows en entornos adecuados.

## Corrección de versión de examen

El usuario especificó **Office Word 2019 Associate**. La referencia correcta es **Microsoft Office Specialist: Word Associate (Office 2019), examen MO-100**; MO-110 queda descartado para este proyecto.

Fuente oficial: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

Microsoft Learn lista seis dominios, duración de 50 minutos e idioma español entre los disponibles. La página consultada no muestra fecha de retirada. Siguiente paso de documentación: revisar el esquema descargable MO-100 para asegurar que el mapa de ejercicios coincide con la versión correcta.


## Revisión oficial de aptitudes completada

Se consultó el PDF oficial de MO-100 y se añadieron los dominios, pesos y subobjetivos a `WORD_COMPETENCIES.md`. La ponderación de tablas y listas es 15–20 %. Fuente: https://arch-center.azureedge.net/Learning/Credentials/mo-100-microsoft-word-2019-skills-measured.pdf

## Actualización — 2026-10-27 (Fase 1 Plan, ITER-001)

- Se creó `ITERATION_PLAN.md` (borrador) para el vertical slice `word-format-001`.
- Se actualizaron `PROJECT_STATE.md` con la iteración actual.
- Se instalaron las skills de SAQI globalmente en `~/.config/opencode/skills/`.
- Pendiente: aprobación humana del plan antes de Fase 2 (Architect).

### Contexto para próxima sesión
- Fase actual: 1 (Plan) — pendiente aprobación.
- Tarea en curso: ninguna implementación; solo planificación.
- Próxima acción: tras aprobación, Fase 2 (Architect): ADRs, contratos de puertos, diagramas C4, actualizar `ARCHITECTURE.md`.

## Actualización — 2026-10-27 (Fase 2 Architect, ITER-001)

- Creados 6 ADRs en `docs/adr/` (ADR-001..ADR-006).
- Creado contrato de puerto `src/domain/ports/evaluation.py`.
- Actualizado `ARCHITECTURE.md` con ADRs, puertos y estructura de código.
- Pendiente: validación contra Skills Nivel A y aprobación de ADRs por humano.

## Actualización — 2026-10-27 (Fase 3 Skill Selection, ITER-001)

- Creado `SKILL_SELECTION.md` con 12 skills Nivel A (saqi-*) + inactivas por stack.
- Registradas versiones 1.0.0.
- Nota: skills globales A-* referenciadas por SAQI no están instaladas; las saqi-* consolidadas las cubren.
- Pendiente: aprobación humana de la selección.

## Actualización — 2026-10-27 (Fase 4 Build, ITER-001)

- Implementado evaluador portable OOXML (`src/virmetrix/infrastructure/evaluators/portable.py`).
- Implementado adaptador Word COM base (`word_com.py`) con detección de disponibilidad.
- Orquestador (`application/orchestrator.py`) y UI mínima PySide6 (`ui/main_window.py`).
- Metadatos y generador de referencia para `word-format-001`.
- Contrato de puertos y modelos de dominio.
- Pruebas: 15 pass (10 unitarias + 5 adversariales). pytest.ini con pythonpath=src.
- Pendiente: UI threading, persistencia (diferida), COM completo en Windows, cobertura/mutation en macro-fase final.

## Actualización — 2026-10-27 (Fase 5 Test Unit & Integration, ITER-001)

- 22 tests pass; cobertura 92.36% (umbral 80%).
- Contract tests OK.
- Listo para Fase 6 (Functional & E2E) / o macro-fase QA final.

## Actualización — 2026-10-27 (Fase 6/7, ITER-001)

- Funcionales 4 pass + UI smoke 1 pass + adversariales 7 pass.
- Total suite: 27+ pass.
- Defectos abiertos: 0.
- Pendiente: Fase 11 (Verify) sign-off y documentación final (Fase 12/13).

## Actualización — 2026-10-27 (Fase 11 Verify, ITER-001)

- Creado `ITERATION_REPORT.md`.
- Sign-off pendiente del humano (PO + Tech Lead + QA Lead).

## Actualización — 2026-10-27 (Fases 12-14, ITER-001)

- Fase 12 Document: `CHANGELOG.md` actualizado, docs archivados en `docs/iterations/ITER-001/`, `CONTEXT.md` actualizado.
- Fase 13 Learn: 4 docs vivos actualizados; checkpoint registrado.
- Fase 14 Improve Skills: propuestas en `SKILL_CHANGELOG.md` (P1-P3) pendientes de gobernanza.
- Release candidate aprobado por el usuario.

## Actualización — 2026-10-27 (ITER-020 Build)

- Estilo visual moderno (QSS).
- Pruebas: 65 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-019 Build)

- Mantener VirMetrix encima de Word (topmost).
- Pruebas: 65 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-018 Build)

- UI compacta estilo gmetrix (siempre encima, pasos numerados).
- Pruebas: 65 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-017 Build)

- UI rediseñada: paneles, pasos numerados, tamaño de ventana.
- Pruebas: 65 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-016 Build)

- Barra de progreso y cancelación en UI.
- Pruebas: 65 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-015 Gate QA)

- Gate QA: 65 tests pass, cobertura 80.52%.
- Bundle Linux validado.
- Seguridad y accesibilidad básicas verificadas.
- Pendiente: licencia, instalador Windows final, validación en VM limpia.

## Actualización — 2026-10-27 (ITER-014 Build)

- Botón "Abrir documento" y apertura automática tras crear copia.
- Pruebas: 53 pass.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-013 Build)

- 12 ejercicios adicionales (2 por dominio).
- Total catálogo: 41.
- Pruebas: 53 pass.
- Pendiente: validación COM en Windows.

## Actualización — 2026-10-27 (ITER-012 Build)

- 12 ejercicios adicionales (2 por dominio).
- Total catálogo: 28.
- Pruebas: 53 pass.
- Pendiente: validación COM en Windows, más ejercicios si se desea.

## Actualización — 2026-10-27 (ITER-011 Build)

- Ejercicios de colaboración (collab-001, collab-002).
- Evaluador portable y COM generalizados.
- Pruebas: 53 pass.
- Cobertura: seis dominios MO-100 con al menos un ejercicio (16 total).
- Pendiente: validación COM en Windows, ampliación de ejercicios por dominio.

## Actualización — 2026-10-27 (ITER-010 Build)

- Ejercicios de administración de documentos (doc-manage-001..003).
- Evaluador portable y COM generalizados.
- Pruebas: 51 pass.
- Pendiente: validación COM en Windows.

## Actualización — 2026-10-27 (ITER-009 Build)

- Ejercicios de elementos gráficos (graphics-001, graphics-002).
- Evaluador portable y COM generalizados.
- Pruebas: 48 pass.
- Pendiente: validación COM en Windows.

## Actualización — 2026-10-27 (ITER-008 Build)

- Ejercicios de referencias (refs-001, refs-002).
- Evaluador portable y COM generalizados.
- Pruebas: 46 pass.
- Pendiente: validación COM en Windows.

## Actualización — 2026-10-27 (ITER-007 Build)

- Ejercicios de tablas y listas (table-001, list-001).
- Evaluador portable y COM generalizados.
- Pruebas: 44 pass.
- Pendiente: validación COM en Windows.

## Actualización — 2026-10-27 (ITER-006 Build)

- Nombres accesibles, tooltips, botón por defecto, diálogo "Acerca de".
- Pruebas: 42 pass.
- Pendiente: validación con lector de pantalla en Windows.

## Actualización — 2026-10-27 (ITER-005 Build)

- PyInstaller spec, main.py, paths.py, scripts de build.
- Bundle Linux validado (arranque + catálogo).
- Pruebas: 42 pass.
- Pendiente: build/validación en Windows + instalador.

## Actualización — 2026-10-27 (ITER-004 Build)

- Modo práctica/examen, temporizador 300 s, pausa/reanudación, finalización idempotente.
- Navegación entre ejercicios.
- Pruebas: 42 pass, 0 fail.
- Pendiente: validación en Windows.

## Actualización — 2026-10-27 (ITER-003 Build)

- Catálogo de 4 ejercicios con criterios declarativos (JSON).
- Evaluador portable y COM generalizados (`_evaluate_specs`, `_evaluate_specs_com`).
- UI con selector de ejercicio.
- Pruebas: 38 pass, 0 fail.
- Pendiente: validación COM de nuevos ejercicios en Windows.

## Actualización — 2026-10-27 (ITER-002 Build parcial)

- Implementado: persistencia SQLite (`AttemptStore`), evaluación en hilo (QThread + `EvaluationWorker`), adaptador Word COM con evaluación real por COM (Windows), metadatos tipados (`target_phrase`, `expected_text`), seguridad ZIP multiplataforma.
- Pruebas: 33 pass, 0 fail (Linux). Word COM validado en Windows (5/5 documento correcto; 3/5 referencia; 3/5 incorrecto).
- Pendiente: validar COM en Windows; catálogo amplio; empaquetado; temporizador.

## Contexto para Próxima Sesión (Checkpoint final ITER-001)

- Fase actual: 14 completada (ITER-001 cerrado, candidato aprobado).
- Tarea en curso: ninguna.
- Archivos clave: ver `context-checkpoints/ITER-001-20261027-checkpoint.md`.
- Próxima acción: definir alcance de ITER-002 con el usuario y ejecutar Fase 1 (Plan).
- Tests: 27+ pass; cobertura 92.36%; defectos 0.
- Skills: saqi-* v1.0.0 activas globalmente.

## Actualizaci�n � 2026-10-27 (Validaci�n Windows + Word COM, ITER-002)

- Gate COM aprobado en Windows 10 Pro (Word v16.0).
- Pruebas: 33 passed, 0 failed (`pytest -q`).
- COM: disponible; sin proceso residual tras correcci�n de `is_available()` (`DispatchEx` + `Quit`).
- Referencia (sin formato): 3/5 (C1 pass, C2 fail, C3 fail, C4 pass, C5 pass).
- Documento correcto: 5/5.
- Documento incorrecto: 3/5 (C2 fail, C3 fail, resto pass).
- UI: sin bloqueo; historial persistente (4 intentos tras reabrir).
- Defecto corregido: `is_available()` dejaba proceso Word residual.
- App en ejecuci�n: PID 5108, ventana �VirMetrix - word-format-001�.
- Pr�xima acci�n: definir ITER-003 (cat�logo de ejercicios / temporizador / empaquetado) y a�adir prueba de regresi�n para residual de proceso.

## Actualización — 2026-11-07 (Sesión de mejora UI y puntuación)

- UI rediseñada en dos ventanas:
  - `ConfigWindow` (vertical): modo, evaluador, copia de trabajo, abrir documento, sesión.
  - `PracticeWindow` (horizontal): ejercicio, tarea, evaluación, resultado, temporizador.
  - Botón “Iniciar práctica” abre la ventana horizontal.
  - Ventana de práctica se ancla horizontalmente (ancho completo, parte inferior).
- Automatización de copias de trabajo:
  - Al cambiar de ejercicio (Anterior/Siguiente) se cierra la copia anterior y se crea/abre automáticamente la nueva.
  - Se resetea la copia de trabajo al cambiar de ejercicio.
- Corrección de puntuación (bold + tamaño de fuente):
  - Portable: soporte de variantes `w:bCs`/`w:iCs`/`w:szCs`; tolerancia a runs fragmentados.
  - Word COM: evaluación a nivel de caracteres; acepta valores `1` además de `-1`/`True`; `not_verifiable` cuando no se resuelve.
- UX:
  - Recordatorio de guardar documento (`Ctrl+S`) antes de evaluar (causa de `not_verifiable`).
  - Botón “Abrir en Word” corregido.
  - Al iniciar práctica se muestra el primer ejercicio sin necesidad de navegar.
- Pruebas: 68 passed, 0 failed; cobertura 80.52%.
- Build Linux: `dist/VirMetrix/VirMetrix` generado y ejecutable.
- Pendiente Windows: revalidar COM/accesibilidad (Narrador) e instalador tras estos cambios.


## Checkpoint — 2026-11-07 (fin de sesión)

### Trabajo realizado
- UI en dos ventanas (ConfigWindow vertical + PracticeWindow horizontal).
- Automatización de copias de trabajo al cambiar de ejercicio.
- Corrección de puntuación (bold/size) en portable y Word COM.
- 35 ejercicios nuevos añadidos (total 89) para cobertura ~90-95% de subobjetivos verificables.
- Tipos de criterio nuevos: hyperlink_present, background_present, textbox_count, endnote_count, tracked_change_count.
- Licencia GPL v2 añadida.
- Tutorial PDF de resolución de ejercicios generado y subido.
- 68 tests pass, cobertura 80.52%.

### Próximos pasos (mañana)
1. Validar en VM Windows 10:
   - Ejecutar `pytest -q` y verificar 68+ pass.
   - Probar los 89 ejercicios en modo `word_com` (especialmente los nuevos de gráficos, referencias y colaboración).
   - Verificar que los ejercicios con criterios no verificables en portable (impresión, accesibilidad, fondos, efectos) se comportan correctamente en COM.
   - Comprobar que no haya procesos Word residuales tras evaluar.
2. Accesibilidad (Narrador): validar nombres, orden de foco, anuncio de temporizador/puntuación (pendiente).
3. Empaquetado Windows: build con PyInstaller + instalador (.exe) y validación de arranque.
4. Revisar y corregir cualquier ejercicio que falle en Windows.
5. Actualizar `QA_RESULTS.md` con evidencia reproducible.

### Archivos clave
- `src/virmetrix/ui/main_window.py` (dos ventanas)
- `src/virmetrix/infrastructure/evaluators/portable.py`
- `src/virmetrix/infrastructure/evaluators/word_com.py`
- `exercises/` (89 ejercicios)
- `docs/Tutorial_Ejercicios_VirMetrix.pdf`
- `LICENSE` (GPL v2)

---

## Checkpoint — 2026-11-09 (sesión actual)

### Trabajo realizado

#### A. Publicación del .exe en GitHub
- Causa: `dist/` y `build/` estaban en `.gitignore`, por eso el binario no aparecía en GitHub.
- Solución:
  - Carpeta `releases/` trackeada (con `.gitkeep` + `README.md`) para subir binarios manualmente.
  - `.github/workflows/build-windows.yml`: compila el `.exe` en Windows (GitHub Actions) en cada push a `main` (artefacto) y publica un **GitHub Release** al crear un tag `v*`.
  - Excepción en `.gitignore` para `releases/*.exe`/`*.zip`.
  - README actualizado con instrucciones de descarga.
- Releases publicados: `v0.1.1`, `v0.1.2`, `v0.1.3`, `v0.1.4` (cada uno con `VirMetrix-windows.zip`).

#### B. Bugs detectados y corregidos (modo COM / Word)

1. **Ejercicio 10 (graphics-002) — instrucción incompleta.**
   - Antes: "Inserta un cuadro de texto con el texto 'Nota lateral'."
   - Ahora: "Inserta una imagen y, a su lado, un cuadro de texto con el texto 'Nota lateral'."
   - Criterios: `drawing_count` esperado 1 → 2 (imagen + cuadro); `textbox_contains` = "Nota lateral".
   - `word_com` `drawing_count`: ahora suma `InlineShapes.Count + Shapes.Count` (antes solo InlineShapes) para contar imagen + cuadro de texto.
   - Título en catálogo: "Imagen y cuadro de texto".
   - Tests actualizados y pasando.

2. **Ejercicio 12 (doc-manage-002) — orientación de página falsea portrait.**
   - Causa: portable usaba el primer `sectPr` (`root.find`); en documentos con saltos de sección intermedios (portrait) se reportaba portrait aunque la sección principal fuera landscape. word_com usaba `doc.PageSetup` (primera sección).
   - Fix: portable usa el último `sectPr` (`findall(...)[-1]`); word_com usa la última sección (`Sections(Sections.Count)`).
   - Reproducido con `multi.docx` (1ª portrait + 2ª landscape) → ahora reporta landscape.

3. **Ejercicio 19 (text-para-002) y 30? no — text-para-008 — color.**
   - Ejercicios afectados: `text-para-002` (rojo, esperado `FF0000`) y `text-para-008` (azul, esperado `0000FF`).
   - Causa: word_com devolvía `Font.Color` como entero OLE_COLOR (BGR) (p. ej. 255) y se comparaba con la cadena hex `"FF0000"` → siempre fail. Mensaje: `Propiedad 'color' en objetivo: 255`.
   - Fix: convertir el entero OLE_COLOR a cadena hex RGB (`r = c & 0xFF`, etc.). 255 → `"FF0000"`.
   - Modo portable ya funcionaba (extrae `w:color/@w:val`).

4. **Ejercicio 31 (text-para-004) — estilo "Heading1" no aplicado.**
   - Causa: word_com comparaba `expected.lower()` (`"heading1"`) como subcadena de `Style.NameLocal` (`"Título 1"` en español o `"Heading 1"` en inglés, con espacio) → nunca coincidía.
   - Fix: normalizar (minúsculas, sin espacios) y comprobar tanto `Style.Name` como `Style.NameLocal`.

5. **Ejercicio 30 (text-para-003) — sangría izquierda.**
   - Causa: el `try/except` envolvía todo el bucle; si un párrafo lanzaba excepción, se abortaba y `ok` quedaba `False`. Mensaje: `Sangría verificada (esperado 720 twips).`
   - Fix: `try/except` por párrafo; tolerancia ampliada a 2 pt; mensaje incluye el valor real leído (`leído X`) para diagnóstico.

#### C. Icono de la aplicación
- Añadido `virmetrix.ico` (raíz del repo) a la UI (`run()` aplica `QIcon` a la app y a la ventana) y al empaquetado (PyInstaller `icon=` + `datas`).

#### D. Commits de esta sesión (en orden)
- `407904b` feat: permitir publicar .exe en GitHub (carpeta releases + workflow build Windows)
- `c522107` fix: añadir permissions contents:write para crear release
- `176eec7` fix: ejercicio 10 (graphics-002) pide imagen + cuadro de texto 'Nota lateral'
- `1323795` fix: orientacion de pagina usa la ultima seccion
- `c8f0c09` fix: color en word_com se convierte de OLE_COLOR (BGR) a hex RGB
- `ed8e35f` fix: paragraph_style en word_com compara nombre de estilo normalizado
- `878a4e0` fix: paragraph_indent en word_com no aborta el bucle y diagnostica el valor leido
- `b03ae01` feat: icono de la app (virmetrix.ico) en UI y en el ejecutable

#### E. Tests
- 68 passed, 0 failed tras cada cambio.

#### F. Pendientes
- Validar los fixes de COM en VM Windows (Word real): color, estilo Heading1, sangría, orientación, drawing_count imagen+textbox.
- Comprobar que no haya procesos Word residuales tras evaluar.
- Accesibilidad (Narrador) pendiente.
- Revisar si Word guarda el color como nombre (`w:val="red"`) en vez de hex en algunos casos (portable podría necesitar normalización).

