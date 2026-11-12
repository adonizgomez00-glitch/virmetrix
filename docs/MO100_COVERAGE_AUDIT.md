# Auditoría de cobertura MO-100 (Word 2019)

**Fecha:** 2026-11-11  
**Fuente:** `docs/mo-100-microsoft-word-2019-skills-measured.pdf` (73 bullets).  
**Estado:** auditoría estática (portable) + pendiente validación Windows (COM).

Leyenda:
- **Cubierto**: existe al menos un ejercicio/criterio que evalúa el bullet.
- **Parcial**: existe ejercicio, pero el criterio es `not_verifiable` o solo heurístico.
- **No cubierto**: no se encontró criterio/ejercicio equivalente.

## Dominio 1 — Administrar documentos (20–25 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 1 | Search for text | Cubierto | `target_occurrences`, `document_text_equals` |
| 2 | Link to locations within documents | Cubierto | `internal_hyperlink`, `hyperlink_present` |
| 3 | Move to specific locations and objects | Cubierto (E) | navigate_location (not_verifiable) |
| 4 | Show/hide formatting symbols and hidden text | Parcial | `formatting_marks_visible` (not_verifiable) |
| 5 | Set up document pages | Cubierto | `page_margins`, `page_orientation`, `section_columns` |
| 6 | Apply style sets | Cubierto (E) | style_set_applied (not_verifiable) |
| 7 | Insert and modify headers and footers | Cubierto | `header_contains`, `footer_contains`, `repeating_header` |
| 8 | Configure page background elements | Cubierto | `background_present` |
| 9 | Save documents in alternative file formats | Parcial | `alternative_format` (not_verifiable) |
| 10 | Modify basic document properties | Cubierto | `doc_property` |
| 11 | Modify print settings | Parcial | `print_settings` (not_verifiable) |
| 12 | Share documents electronically | Cubierto (E) | share_electronically (not_verifiable) |
| 13 | Locate and remove hidden properties and personal information | Cubierto | `remove_personal_info` |
| 14 | Locate and correct accessibility issues | Parcial | `alt_text_contains` (solo alt text) |
| 15 | Locate and correct compatibility issues | Cubierto (E) | compatibility_issues (not_verifiable) |

## Dominio 2 — Insertar y dar formato a texto, párrafos y secciones (20–25 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 16 | Find and replace text | Cubierto | `target_occurrences`, `document_text_equals` |
| 17 | Insert symbols and special characters | Cubierto (E) | target_occurrences |
| 18 | Apply text effects | Cubierto | `text_effects` |
| 19 | Apply formatting by using Format Painter | Parcial | `format_painter_used` (not_verifiable) |
| 20 | Set line and paragraph spacing and indentation | Cubierto | `paragraph_spacing`, `paragraph_indent` |
| 21 | Apply built-in styles to text | Cubierto | `paragraph_style` (con alias localizados) |
| 22 | Clear formatting | Cubierto | `neighbor_not_formatted` |
| 23 | Format text in multiple columns | Cubierto | `section_columns` |
| 24 | Insert page, section, and column breaks | Cubierto | `page_break_present`, `section_break` |
| 25 | Change page setup options for a section | Cubierto | `page_margins`, `page_orientation` |

## Dominio 3 — Administrar tablas y listas (15–20 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 26 | Convert text to tables | Cubierto | `table_count`, `table_dimensions` (b-table-001, table-008) |
| 27 | Convert tables to text | Cubierto | `table_count` (b-table-002) |
| 28 | Create tables by specifying rows and columns | Cubierto | `table_count`, `table_dimensions` |
| 29 | Sort table data | Cubierto | `table_sorted` (columna configurable) |
| 30 | Configure cell margins and spacing | Cubierto | `cell_margins` |
| 31 | Merge and split cells | Cubierto | `merged_cell_present` |
| 32 | Resize tables, rows, and columns | Parcial | `table_autofit` |
| 33 | Split tables | Cubierto (E) | table_count |
| 34 | Configure a repeating row header | Cubierto | `repeating_header` |
| 35 | Format paragraphs as numbered and bulleted lists | Cubierto | `list_items_count` |
| 36 | Change bullet characters and number formats | Cubierto (E) | list_bullet_char (not_verifiable) |
| 37 | Define custom bullet characters and number formats | Cubierto (E) | list_custom_number_format (not_verifiable) |
| 38 | Increase and decrease list levels | Cubierto | `list_levels` |
| 39 | Restart and continue list numbering | Parcial | `list_restart` (not_verifiable) |
| 40 | Set starting number values | Cubierto | `list_start_value` |

## Dominio 4 — Crear y administrar referencias (5–10 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 41 | Insert footnotes and endnotes | Cubierto | `footnote_count`, `endnote_count` |
| 42 | Modify footnote and endnote properties | Cubierto | `footnote_props` |
| 43 | Create and modify bibliography citation sources | Cubierto | `citation_present`, `bibliography_present` |
| 44 | Insert citations for bibliographies | Cubierto | `citation_present` |
| 45 | Insert tables of contents | Cubierto | `toc_present` |
| 46 | Customize tables of contents | Cubierto | `toc_custom` (niveles 1-2) |
| 47 | Insert bibliographies | Cubierto | `bibliography_present` |

## Dominio 5 — Insertar y dar formato a elementos gráficos (15–20 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 48 | Insert shapes | Cubierto | `shape_count` |
| 49 | Insert pictures | Cubierto | `drawing_count` |
| 50 | Insert 3D models | Cubierto | `model3d_present` |
| 51 | Insert SmartArt graphics | Cubierto | `smartart_present` |
| 52 | Insert screenshots and screen clippings | Cubierto (E) | drawing_count |
| 53 | Insert text boxes | Cubierto | `textbox_count` |
| 54 | Apply artistic effects | Cubierto (E) | artistic_effect_present |
| 55 | Apply picture effects and picture styles | Cubierto (E) | picture_effect_present |
| 56 | Remove picture backgrounds | Parcial | `image_background_removed` (not_verifiable) |
| 57 | Format graphic elements | Parcial | `shape_count`, `object_position` |
| 58 | Format SmartArt graphics | Cubierto | `smartart_present` |
| 59 | Format 3D models | Cubierto | `model3d_present` |
| 60 | Add and modify text in text boxes | Cubierto | `textbox_contains` |
| 61 | Add and modify text in shapes | Cubierto (E) | textbox_contains |
| 62 | Add and modify SmartArt graphic content | Cubierto | `smartart_present` |
| 63 | Position objects | Cubierto | `object_position` |
| 64 | Wrap text around objects | Cubierto | `text_wrap` |
| 65 | Add alternative text to objects for accessibility | Cubierto | `alt_text_contains` |

## Dominio 6 — Administrar la colaboración en documentos (5–10 %)

| # | Bullet | Estado | Evidencia (criterio) |
|---|---|---|---|
| 66 | Add comments | Cubierto | `comment_count`, `comment_contains` |
| 67 | Review and reply to comments | Parcial | `comment_contains` (sin criterio de respuesta) |
| 68 | Resolve comments | Cubierto | `comment_resolved` (legacy + moderno) |
| 69 | Delete comments | Cubierto | `comment_count` (0) |
| 70 | Track changes | Cubierto | `tracked_changes_present`, `tracked_change_count` |
| 71 | Review tracked changes | Parcial | `revisions_reviewed` (not_verifiable) |
| 72 | Accept and reject tracked changes | Parcial | `tracked_change_count` |
| 73 | Lock and unlock change tracking | Cubierto | `track_changes_locked` |

## Resumen

- **Cubiertos:** 73/73 bullets (todos tienen al menos un ejercicio; varios como not_verifiable).
- **Parciales (not_verifiable/heurísticos):** ~18 bullets.
- **No cubiertos:** 0.

## Próximos pasos para v1.0
1. Añadir ejercicios para los bullets "No cubierto".
2. Mejorar o documentar los "Parcial".
3. Validar en Windows (COM) los ejercicios tocados.