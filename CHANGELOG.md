# Changelog

## 2026-11-11 — v0.4.2 (corrección tipográfica)

- `b-doc-002`: corregido "Símbodos" → "Símbolos" en metadata y catálogo.

## 2026-11-11 — v0.4.2

### Fixed
- UI: selector de evaluación incluye el módulo D (faltaba "Evaluación D" y "Todo" no lo sumaba).
- `d-table-003`: evaluador `table_sorted` ahora ordena por la columna indicada (Precio), no solo por la primera columna; soporte de `column` en el spec (portable y COM).
- `d-list-001/002/003`: `reference.docx` con `numPr` válido (`w:ilvl w:val` / `w:numId w:val`) para abrir en Word COM.
- `d-collab-001`: detección de comentario resuelto en comentarios modernos (Office 2024) leyendo `word/commentsExtended.xml` (`w15:done` / `w16cid:done`).
- `paragraph_style`: acepta alias localizados (Heading2 / Encabezado2 / Ttulo2, etc.) en portable y COM.

### Changed
- `APP_VERSION` actualizada a 0.4.2.

## 2026-11-10 — v0.4.0

### Added
- Módulo D: 20 ejercicios nuevos similares a B (texto/párrafos, tablas, listas, referencias, gráficos, colaboración y documento), con `metadata.json`, `reference.docx` válidos y referencias sin solución preaplicada.
- Todos los ejercicios D incluidos en `exercises/catalog.json` (la app los carga al iniciar).

### Fixed
- `d-table-004`: `expected` de `cell_margins` normalizado a formato plano (57 twips).
- `d-collab-001` / `d-collab-002`: `reference.docx` con el comentario correcto (sin resolver / presente) según la guía de generación.
- `d-list-001` / `d-list-002` / `d-list-003`: referencia con número y `ilvl` consistentes para la tarea de niveles/reinicio/inicio.
- Guidelines `EXERCISE_GENERATION_GUIDELINES.md`: aplicables a A, B, C y D; documentada la excepción intencional de `page_margins` (dict).
- `d-refs-001`, `d-refs-004`, `b-refs-001`, `b-refs-004`, `text-para-011`: referencias sin pasar trivialmente; se mantienen como tarea activa.
- `portable.py` / `word_com.py`: criterios triviales marcados `not_verifiable`: `alternative_format`, `formatting_marks_visible`, `print_settings`, `format_painter_used`, `image_background_removed`, `revisions_reviewed`, `list_restart`.
- `main_window.py`: diálogo "Acerca de" muestra `APP_VERSION` (igual a la tag); APP_VERSION 0.4.0.

### Releases
- Tag `v0.4.0`: workflow `.github/workflows/build-windows.yml` genera el ejecutable y lo adjunta a la release.

## 2026-10-27 — Corrección del examen objetivo

- Cambiado el objetivo de MO-110 (Microsoft 365 Apps) a MO-100 (Office 2019).
- Actualizadas las referencias oficiales y las instrucciones para OpenCode.
- Se mantiene el stack Python + PySide6 + python-docx + OOXML + pywin32 (Windows) + SQLite + pytest.
- Se mantiene el requisito de MVP 100 % offline y el instalador para Windows.
- Pendiente: verificar el archivo oficial “MO-100 Exam Skills Outline” antes de fijar subobjetivos y pesos.

## 2026-10-27 — ITER-001: vertical slice word-format-001 (candidato)

### Added
- Evaluador portable de DOCX basado en OOXML (`src/virmetrix/infrastructure/evaluators/portable.py`).
- Adaptador Word COM base con detección de disponibilidad (`word_com.py`).
- Orquestador de evaluación y copia de trabajo (`application/orchestrator.py`).
- UI mínima PySide6 (`ui/main_window.py`).
- Metadatos y generador del DOCX de referencia (`exercises/word-format-001/`).
- Contrato de puerto `EvaluationAdapter` y modelos de dominio (`domain/ports.py`).
- Pruebas: 27+ pass (unitarias, integración, funcionales, adversariales, UI smoke).
- Documentación SAQI: `ITERATION_PLAN.md`, `SKILL_SELECTION.md`, ADRs, `ITERATION_REPORT.md`.

### Changed
- `ARCHITECTURE.md`: ADRs, puertos, estructura de código.
- `CONTEXT.md`, `PROJECT_STATE.md`, `QA_RESULTS.md`, `SESSION.md` actualizados.

### Fixed
- Ninguno (0 defectos registrados).

## 2026-11-09 — Sesión: publicación del .exe y corrección de bugs COM

### Added
- Carpeta `releases/` trackeada para publicar binarios en el repositorio.
- Workflow `.github/workflows/build-windows.yml` (build Windows en GitHub Actions + release en tag `v*`).
- Icono de la aplicación `virmetrix.ico` en UI y en el ejecutable (PyInstaller).
- Excepción en `.gitignore` para `releases/*.exe`/`*.zip`.

### Changed
- Ejercicio 10 (graphics-002): instrucción ahora pide imagen + cuadro de texto "Nota lateral"; `drawing_count` esperado 2; título "Imagen y cuadro de texto".
- `word_com.drawing_count` suma `InlineShapes + Shapes`.

### Fixed
- Orientación de página (ej. 12): portable/word_com usan la última sección en vez de la primera (falseaba portrait).
- Color en word_com (ej. 19, text-para-008): conversión de OLE_COLOR (BGR) a hex RGB.
- Estilo "Heading1" en word_com (ej. 31): comparación normalizada (minúsculas, sin espacios, Name y NameLocal).
- Sangría izquierda en word_com (ej. 30): `try/except` por párrafo, tolerancia 2 pt, diagnóstico del valor leído.

### Releases
- `v0.1.1`, `v0.1.2`, `v0.1.3`, `v0.1.4` publicados con `VirMetrix-windows.zip`.
