# Plan — Evaluación B (cobertura de huecos MO-100)

**Fecha:** 2026-11-09
**Estado:** PLANIFICACIÓN — pendiente de aprobación del usuario.
**Contexto:** La Evaluación A (89 ejercicios actuales) cubre ~40 % de los bullets oficiales
del PDF "MO-100 Skills Measured" (29/73 bullets). La Evaluación B debe cubrir los
44 bullets restantes para alcanzar cobertura completa.

## 1. Huecos detectados (44 bullets no cubiertos)

### Dominio 1 — Administrar documentos (20–25 %) — 10 de 15 bullets
1. Vincular a ubicaciones dentro de documentos (hipervínculo a marcador).
2. Mover a ubicaciones y objetos específicos.
3. Mostrar/ocultar símbolos de formato y texto oculto.
4. Guardar documentos en formatos alternativos (PDF, RTF, etc.).
5. Modificar configuración de impresión.
6. Compartir documentos electrónicamente (adjuntar, exportar).
7. Localizar y eliminar propiedades ocultas e información personal.
8. Localizar y corregir problemas de accesibilidad.
9. Localizar y corregir problemas de compatibilidad.
10. Aplicar conjuntos de estilos (style sets).

### Dominio 2 — Texto, párrafos y secciones (20–25 %) — 3 de 10 bullets
11. Aplicar efectos de texto (sombra, contorno, reflejo).
12. Aplicar formato con Pincel de formato (Format Painter).
13. Insertar saltos de página, sección y columna.

### Dominio 3 — Tablas y listas (15–20 %) — 9 de 15 bullets
14. Convertir tablas a texto.
15. Ordenar datos de tabla.
16. Configurar márgenes y espaciado de celda.
17. Redimensionar tablas, filas y columnas.
18. Cambiar caracteres y formatos de numeración.
19. Definir caracteres y numeración personalizados.
20. Aumentar y disminuir niveles de lista.
21. Reiniciar y continuar numeración de lista.
22. Establecer valores de número inicial.

### Dominio 4 — Referencias (5–10 %) — 5 de 7 bullets
23. Modificar propiedades de notas al pie/final.
24. Crear y modificar fuentes de citas bibliográficas.
25. Insertar citas para bibliografías.
26. Personalizar tablas de contenido.
27. Insertar bibliografías.

### Dominio 5 — Elementos gráficos (15–20 %) — 13 de 18 bullets
28. Insertar iconos / modelos 3D.
29. Insertar SmartArt.
30. Insertar capturas de pantalla.
31. Aplicar efectos artísticos.
32. Aplicar efectos y estilos de imagen.
33. Quitar fondo de imagen.
34. Formatear elementos gráficos (en general).
35. Formatear SmartArt.
36. Formatear modelos 3D.
37. Agregar/modificar texto en SmartArt.
38. Posicionar objetos.
39. Ajustar texto alrededor de objetos.
40. (Alt text ya cubierto parcialmente — mantener.)

### Dominio 6 — Colaboración (5–10 %) — 4 de 8 bullets
41. Resolver comentarios.
42. Eliminar comentarios.
43. Revisar cambios registrados.
44. Bloquear/desbloquear seguimiento de cambios.

## 2. Estimación de ejercicios nuevos

- 44 bullets no cubiertos.
- Algunos bullets se agrupan en un ejercicio (p. ej. "Insertar formas + texto en formas" = 1 ejercicio).
- Estimación: **~30 ejercicios nuevos** para cubrir los 44 bullets (promedio 1.5 bullets/ejercicio).
- Total final: 89 (A) + ~30 (B) = ~119 ejercicios.

## 3. Opciones de organización (a elegir)

### Opción 1 — Módulo separado `evaluation-b/`
- Nueva carpeta `exercises-b/` con su propio `catalog.json`.
- Loader modificado para cargar ambos catálogos.
- Ventaja: separación clara A/B; fácil activar/desactivar.
- Desventaja: cambios en loader y paths.

### Opción 2 — Prefijo de IDs en el mismo `exercises/`
- IDs como `b-doc-manage-020`, `b-graphics-016`, etc.
- Un solo catálogo con campo `module: "A"/"B"`.
- Ventaja: sin cambios de estructura; loader actual funciona.
- Desventaja: catálogo mezclado.

### Opción 3 — Campo `module` en el catálogo actual
- Añadir `"module": "A"` a los 89 actuales y `"module": "B"` a los nuevos.
- UI filtra por módulo (Evaluación A / B / Todo).
- Ventaja: flexible; una sola fuente.
- Desventaja: requiere cambios en UI y loader.

## 6b. Estado de implementación de la Evaluación B

- **UI:** selector de módulo (A/B/Todo) implementado en `ConfigWindow` (`module_combo`).
- **Loader:** `ExerciseMetadata.module` añadido; `list_exercises()` devuelve `module`.
- **Ejercicios B generados:** 33 ejercicios (b-*) con metadata + reference.docx básicos.
- **Criterios nuevos (no implementados aún):** los ejercicios B usan tipos de criterio nuevos
  (section_break_present, cell_margins, text_effects, format_painter_used, smartart_present,
  model3d_present, image_background_removed, object_position, text_wrap, comment_resolved,
  comment_count=0, revisions_reviewed, track_changes_locked, internal_hyperlink,
  formatting_marks_visible, alternative_format, print_settings, remove_personal_info,
  text_effects, format_painter_used, section_break, footnote_props, citation_present,
  bibliography_present, toc_custom, shape_count, smartart_present, model3d_present,
  image_background_removed, object_position, text_wrap, comment_resolved, comment_count,
  revisions_reviewed, track_changes_locked, internal_hyperlink, formatting_marks_visible,
  alternative_format, print_settings, remove_personal_info).
  Estos devuelven "Tipo de criterio desconocido" (ERROR) hasta que se implementen en
  `portable.py` y `word_com.py`.

## 7. Requisito de UI: selector de Evaluación A/B

Para soportar dos módulos, la ventana de configuración (`ConfigWindow`) debe incluir
un selector (combo o radio) **"Evaluación: A / B / Todo"** antes de iniciar la sesión,
análogo al selector de modo (Práctica/Examen). El selector filtra `exercise_ids` de la
sesión según el módulo elegido.

- Si se elige "Todo", se mezclan A y B (útil para simulacro completo).
- El historial SQLite debe registrar el módulo de cada intento (añadir columna `module`
  a la tabla de intentos o derivarlo del `exercise_id`).
- El panel de resultados puede mostrar el módulo activo.

Este requisito aplica especialmente a la Opción 3, pero también es deseable en la
Opción 2 (con filtro por prefijo de ID).

## 4. Verificabilidad (portable vs COM)

No todos los bullets nuevos son verificables automáticamente:
- **Verificables en portable:** guardar en otro formato (inspeccionar), propiedades, márgenes, columnas, estilos, listas, TOC, alt text, comentarios (conteo), cambios registrados.
- **Solo COM (Word):** efectos de texto, Format Painter, saltos, ordenar tabla, márgenes de celda, redimensionar, niveles de lista, reiniciar numeración, modificar propiedades de notas, citas/bibliografía, SmartArt, 3D, capturas, efectos de imagen, quitar fondo, posicionar, ajustar texto, resolver/eliminar comentarios, revisar cambios, bloquear seguimiento, accesibilidad/compatibilidad, propiedades ocultas.

Esto implica que varios ejercicios de la Evaluación B solo serán evaluables en Windows (modo `word_com`), y en Linux quedarán `not_verifiable` o requerirán inspección manual.

## 5. Recomendación (a tu elección)

- **Aprobar Opción 2 o 3** para no tocar la estructura de carpetas.
- **Priorizar primero** los bullets verificables en portable (para que funcionen en Linux): dominio 1 (propiedades, márgenes, columnas), dominio 2 (saltos, efectos verificables), dominio 3 (niveles, reiniciar numeración — solo COM), dominio 4 (TOC personalizado, bibliografías — solo COM), dominio 5 (alt text ya cubierto), dominio 6 (comentarios ya cubierto).
- **Generar en lotes** de 5–10 ejercicios, validando portable + COM en VM Windows.

## 6. Pregunta para ti

¿Qué opción de organización prefieres (1, 2 o 3)? ¿Y quieres que empiece generando el primer lote de la Evaluación B (p. ej., los 5 bullets más fáciles/verificables) o prefieres revisar el plan primero?