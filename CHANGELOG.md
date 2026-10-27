# Changelog

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
