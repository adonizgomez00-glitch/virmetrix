# Guía para generar nuevos ejercicios (Evitar errores conocidos)

**Objetivo:** que los ejercicios (Evaluación A, B, C y D) sean **verificables de verdad**,
no triviales, y funcionen tanto en portable como en Word COM en Windows.

Aplica a cualquier ejercicio nuevo, regenerado o auditado. Complementa `docs/EVALUATION_B_PLAN.md`.

---

## 1. Principios generales

1. **Un ejercicio = una tarea observable** con un resultado verificable en el `.docx`
   (o, si es puramente visual/configuración de app, marcarlo como `not_verifiable`
   y no contar en la puntuación).
2. **La referencia (`reference.docx`) debe ser el punto de partida SIN la solución aplicada.**
   - Al evaluar la referencia en portable, el score debe ser **parcial** (ni 0 de todo
     por error, ni el máximo).
   - Si la referencia ya contiene la solución → "pass trivial" (inútil). Se detectó en
     39 ejercicios de A y 7 de B; reducir a 0.
3. **La instrucción debe coincidir exactamente con lo que el criterio verifica.**
   - Ej.: "Aplica sombra exterior" (no "sombra" genérica si solo exterior se detecta).
4. **Todo criterio nuevo debe implementarse en `portable.py` Y `word_com.py`**
   con la misma semántica, o marcarse `not_verifiable` en ambos.
5. **Probar en Windows (Word COM) antes de declarar validado.** La evidencia va en
   `QA_RESULTS.md`.

---

## 2. Checklist de generación (obligatorio)

### 2.1 Referencia (`reference.docx`)
- [ ] Abre con `python-docx` sin error.
- [ ] Pasa zip + parseo XML + conversión LibreOffice (sin corrupción).
- [ ] Contiene los **elementos base necesarios** para la tarea:
  - Si la tarea es "resolver comentario" → debe existir un comentario **no resuelto**
    (con `commentRangeStart/End` + `commentReference` en el documento y `comments.xml`).
  - Si la tarea es "eliminar comentario" → debe existir un comentario.
  - Si la tarea es "notas al pie" → debe existir al menos una nota.
  - Si la tarea es "TOC" → debe existir un campo `TOC`.
  - Etc. Sin el elemento base, el ejercicio es imposible de resolver.
- [ ] **No** contiene la solución preaplicada (verificar: score de la referencia ≠ máximo).
- [ ] `content_types` y `rels` incluyen todas las partes (comments, footnotes, etc.).
  - Bug histórico: refs-007..013 corruptos por falta de relación/content type de
    footnotes/endnotes.

### 2.2 Criterios (`criteria_specs`)
- [ ] Cada criterio tiene un `type` implementado en ambos evaluadores.
- [ ] El criterio verifica el **resultado**, no la presencia de una palabra clave cuando
  sea posible.
- [ ] Evitar criterios **triviales (siempre pass)**. Los detectados:
  - `formatting_marks_visible` (configuración de app, no del documento)
  - `alternative_format` (no verificable desde el archivo)
  - `print_settings` (no verificable desde el archivo)
  - `format_painter_used` (no verificable desde el archivo)
  - `list_restart`, `image_background_removed`, `revisions_reviewed` (heurísticos débiles)
  - Si la tarea es inherentemente no verificable → marcar `not_verifiable` y **excluir**
    del `score_possible` (o documentar como verificación manual).
- [ ] Verificar que la referencia **no** obtiene el máximo (test automático).
- [ ] Verificar que un documento "correcto" obtiene el máximo (test automático).

### 2.3 Implementación en evaluadores
- [ ] `portable.py`: implementar el `type` con lógica OOXML real.
- [ ] `word_com.py`: implementar el `type` con la **propiedad COM correcta**:
  - Comprobar el enum/valor real (ej.: `WdNoteNumberStyle`: romanos = 1 y 2, no 5–8).
  - Para sombras avanzadas (interior/exterior/perspectiva) usar `WordOpenXML`, no
    `Font.Shadow` (solo cubre sombra legada).
  - Para bibliografía: buscar en `doc.Fields` (código `BIBLIOGRAPHY`), no en
    `Content.Text`.
  - Para marcas de formato: es configuración de vista de la app → no detectable en
    la instancia de automatización → marcar como verificación manual.
  - Evitar leer propiedades de la ventana del usuario desde la instancia COM de la app
    (son instancias distintas).
- [ ] Ambos evaluadores deben coincidir en resultado para el mismo documento.

### 2.4 Pruebas
- [ ] Añadir test en `tests/unit/test_new_exercises.py` (o `test_catalog_exercises.py`)
  para el nuevo ejercicio: documento correcto → score completo.
- [ ] Añadir test adversarial si aplica (corrupción, traversal).
- [ ] Ejecutar `pytest -q` → todos pass.
- [ ] Validar en Windows: `pytest -q`, COM disponible, evaluación de referencia y
  documento correcto con scores esperados.

### 2.5 UI y distribución
- [ ] No marcar como "pendiente" si ya está validado.
- [ ] Actualizar `APP_VERSION` y crear tag `v*` para generar `.exe` y permitir
  actualización vía botón Update.

---

## 3. Errores históricos (A y B) y su causa

### Evaluación A
- **Referencias con solución preaplicada** (40 casos) → pass trivial. Fix: regenerar
  referencias sin formato aplicado.
- **Corrupción en Word** (refs-007..013): `footnotes.xml`/`endnotes.xml` sin relación
  ni content type. Fix: añadir relación y content type.
- **collab-012**: tracked changes corruptos. Fix: regrabar con LibreOffice.
- **Bugs COM** (`bugs_com.csv`): orientación/color/estilo Heading1/sangría/etc.
  detectados por usar la primera sección o comparar strings sin normalizar.

### Evaluación B
- **b-collab-001**: referencia sin comentario → imposible resolver. Fix: añadir
  comentario no resuelto.
- **b-doc-002**: COM leía `ShowAll` de la instancia de automatización → falso
  negativo. Fix: marcar como verificación manual.
- **b-text-001**: solo detectaba `Font.Shadow` (sombra legada). Fix: detección por
  XML para sombras avanzadas + instrucción "sombra exterior".
- **b-refs-005**: enum `WdNoteNumberStyle` incorrecto (5–8 en vez de 1–2). Fix:
  valores correctos + fallback XML.
- **b-refs-003**: buscaba "BIBLIOGRAPHY" en texto visible (está en código de campo).
  Fix: recorrer `doc.Fields` + fallback XML.
- **7 ejercicios triviales** (b-list-002, b-graphics-004, b-collab-003, b-doc-002/003/004,
  b-text-002): criterios siempre pass. Pendiente de endurecer.

---

## 4. Regla de oro

> Antes de dar por válido un ejercicio nuevo, ejecutar:
> `portable(reference) != max` y `portable(correcto) == max`,
> y en Windows: `COM(reference)` y `COM(correcto)` con los scores esperados.
> Registrar la evidencia en `QA_RESULTS.md`.

---

## 5. Lecciones adicionales (validación Windows)

- **Orientación de página (COM):** el enum `WdOrientation` es `0 = portrait`,
  `1 = landscape`. No asumir 1/2.
- **Actualización de la app:** el botón Update descarga la última release, muestra
  barra de progreso y pregunta si reiniciar al finalizar.
- **Resumen del módulo:** en modo examen muestra el último intento; en práctica
  el mejor. En examen no se puede navegar sin iniciar ni iniciar si está en pausa.
- **Encabezados/pies (COM):** recorrer todas las secciones y los 3 tipos de
  encabezado/pie (default, first, even), no solo la primera sección.
- **Márgenes/columnas (COM):** usar la última sección (principal), no la primera.
- **Propiedades de documento (COM):** normalizar nombres (Title/Subject/Author/...)
  y comparar sin distinguir mayúsculas/espacios.
- **Texto alternativo (COM):** fallback a `docPr/@descr` en el XML si
  `AlternativeText` está vacío.

---

## 6. Errores detectados en Evaluación C (y cómo evitarlos)

### 6.1 Metadata con formato anidado (C-10, C-31)
- **Problema:** algunos `criteria_specs` usaban `expected` como objeto anidado
  (`expected: { property, value }` o `{ row, col, text }`), mientras el evaluador
  y el resto de ejercicios esperan formato **plano** (`property` + `expected`,
  o `row`/`col`/`expected`).
- **Síntoma:** el criterio siempre falla aunque el documento sea correcto, porque
  `spec.get("property")`/`spec.get("expected")` devuelven `None` o un dict.
- **Fix:** normalizar la metadata a formato plano y validar que todos los
  `criteria_specs` usen la misma estructura que el evaluador espera.
- **Regla:** al generar ejercicios nuevos, copiar la estructura de un ejercicio
  ya validado (p. ej. `doc-manage-003` o `table-001`) y no inventar anidamientos.
- **Excepción (no confundir con el error anterior):** `page_margins` usa
  `expected` como dict `{top,bottom,left,right}` de forma **intencional**
  (cada clave se valida contra `w:pgMar`); `cell_margins` usa `expected`
  numérico (twips, p. ej. `57`). Otros criterios (`table_cell_text`, ...)
  usan claves planas `row`/`col`/`expected`. No rebajar estos casos a plano.

### 6.2 COM lee `None` en sangría/espaciado (C-48, C-49)
- **Problema:** `ParagraphFormat.LeftIndent` / `SpaceAfter` devuelven `None` en
  COM aunque el documento tenga el valor, por lo que el criterio falla.
- **Fix:** añadir respaldo por XML (`w:ind w:left="..."`, `w:spacing w:after="..."`)
  buscando el valor exacto en twips en el XML del párrafo/documento.
- **Regla:** para criterios de formato de párrafo, implementar siempre un fallback
  por XML además de la propiedad COM.

### 6.3 Etiquetas y unidades (UX)
- **Problema:** las instrucciones mostraban solo twips o usaban pulgadas, mientras
  Word en español trabaja en cm.
- **Fix:** mostrar twips + pt + cm en las etiquetas y convertir pulgadas a cm.
- **Regla:** las instrucciones deben usar la unidad que el usuario ve en Word
  (cm) y, si aplica, incluir la equivalencia en twips/pt.