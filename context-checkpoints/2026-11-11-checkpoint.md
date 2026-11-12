# Checkpoint — 2026-11-11 (v0.4.2)

## Estado
- Rama: main
- Último commit: 20265db (v0.4.2) + cambios posteriores sin commit (docs, list_start_value COM, toc_custom, checklist).
- Tag publicado: v0.4.2 (push realizado).
- Tests: 68 passed.
- Build Linux: OK (`dist/VirMetrix/VirMetrix`).

## Cambios aplicados (sin push adicional)
- UI: módulo D en selector de evaluación.
- `table_sorted`: columna configurable (Precio) en portable y COM.
- `d-list-001/002/003`: corrección de `numPr` para Word COM.
- `d-collab-001`: detección de comentario resuelto en comentarios modernos (Office 2024).
- `paragraph_style`: alias localizados (Heading2/Encabezado2/Ttulo2).
- `d-list-003` (`list_start_value`): respaldo OOXML en COM.
- `d-refs-001` (`toc_custom`): verificación de niveles 1-2.
- Docs: CHANGELOG, SESSION, QA_RESULTS, WINDOWS_VALIDATION, EXERCISE_GENERATION_GUIDELINES actualizados.

## Pendiente en Windows
- Validar Word COM para ejercicios tocados (ver checklist en WINDOWS_VALIDATION.md).
- Registrar evidencia en QA_RESULTS.md.

## Notas
- No se realizó push adicional tras las últimas correcciones (según instrucción del usuario).
- Build Windows se generará por tag v0.4.2.