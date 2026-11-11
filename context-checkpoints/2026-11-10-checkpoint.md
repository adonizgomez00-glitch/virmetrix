# CHECKPOINT — 2026-11-10

## Proyecto
VirMetrix (simulador MOS Word Associate MO-100)

## Iteración / Fase
v0.4.0 publicada. Prácticamente cierra Evaluación D y limpieza A/B/C/D.

## Tarea actual
Ninguna en curso. Sesión cerrada por el usuario.

## Archivos clave
- src/virmetrix/ui/main_window.py (APP_VERSION = 0.4.0)
- src/virmetrix/infrastructure/evaluators/portable.py (6 criterios NOT_VERIFIABLE)
- src/virmetrix/infrastructure/evaluators/word_com.py (mismas semánticas)
- docs/EXERCISE_GENERATION_GUIDELINES.md (A/B/C/D, excepción page_margins)
- exercises/catalog.json (192 ejercicios; módulo D registrado)
- CHANGELOG.md (entrada v0.4.0)

## Estado del proyecto (snapshot)
- Tests: 68 pass, 0 fail.
- Catálogo: 192 ejercicios; A=89, B=33, C=50, D=20.
- Referencia + metadata válidos para los 20 ejercicios D nuevos; d-collab-001/002 con comentario base; d-list-001..003 con numPr/ilvl.
- Referencias triviales corregidas: b-refs-001, b-refs-004, text-para-011.
- Criterios triviales marcados not_verifiable (igual que alternativa_format desde ITER-021): alternative_format, formatting_marks_visible, print_settings, format_painter_used, image_background_removed, revisions_reviewed, list_restart.
- Excepción documentada: `page_margins.expected` = dict {top,bottom,left,right} (intencional, no es metadata anidada defectuosa).
- Git: origen/main en 6d81df2; tag v0.4.0 pusheado.
- CI Build Windows: ejecutado por el tag; release v0.4.0 adjuntará automáticamente `VirMetrix-windows.zip` cuando termine.

## Próximo paso al reanudar
1. Leer este checkpoint + `SESSION.md` + `CHANGELOG.md` (v0.4.0).
2. Verificar release v0.4.0 en GitHub: que el .exe quedó adjunto.
3. Validación en Windows (Word COM) de los ejercicios D (referencia ≠ máximo y documento correcto = máximo); registrar evidencia en `QA_RESULTS.md`.
4. Deuda conocida: text-para-011 usa `document_text_equals` (ignora formato); rediseñar criterio con un tipo que valide ausencia de formato en runs, o marcar el ejercicio como verificación manual.

## Skills activas
saqi-context-manager, saqi-adaptive-mode, saqi-vertical-slice, saqi-harness.

## Contexto restante
Sesión cerrada. Re-hidratar desde este checkpoint. Los últimos cambios están commiteados en origin/main (6d81df2) y tag v0.4.0. Si el usuario lo desea, commitear este checkpoint + SESSION.md/CONTEXT.md como docs de cierre (requiere confirmación, no se ha hecho).
