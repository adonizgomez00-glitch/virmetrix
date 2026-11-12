# ITERATION_PLAN — VirMetrix ITER-022

**Iteración:** ITER-022
**Fase SAQI:** 4 — Build
**Modo SAQI:** estándar.
**Vertical slice:** criterios no verificables → verificables cuando hay evidencia documental.

## Objetivo

Reducir los criterios `not_verifiable` convirtiendo en verificables automáticamente los que dejan evidencia en el `.docx` o en el modelo de objetos de Word, y excluir del puntaje los que son inherentemente acciones de aplicación.

## Alcance implementado

- `list_bullet_char`: detección de viñeta personalizada en `numbering.xml` (portable y COM).
- `list_custom_number_format`: detección de formato de numeración no viñeta en `numbering.xml` (portable y COM).
- `list_restart`: detección de `w:startOverride` en `numbering.xml` (portable y COM).
- `revisions_reviewed`: detección de ausencia de cambios registrados (`tracked_change_count == 0` / `doc.Revisions.Count == 0`).
- Puntuación: `score_possible` excluye criterios `not_verifiable`.
- UI: muestra `N/A` / "verificación manual" cuando no hay criterios puntuables.

## Criterios que permanecen no verificables

`style_set_applied`, `share_electronically`, `compatibility_issues`, `navigate_location`, `image_background_removed`, `formatting_marks_visible`, `alternative_format`, `print_settings`, `format_painter_used` son acciones/configuración de aplicación o dependen del estado de Word en tiempo real; no dejan evidencia fiable en el `.docx`. Para automatizarlos se requiere un complemento de Word (COM/VSTO/Office.js) o rediseñar la tarea como resultado documental.

## Validación

- `pytest -q` → 72 passed.
