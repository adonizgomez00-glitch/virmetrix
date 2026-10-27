# Vertical Slice 001 — `word-format-001`

**Estado:** especificación de aceptación; todavía no implementado.  
**Objetivo:** demostrar el recorrido completo desde seleccionar un ejercicio hasta recibir una evaluación verificable, tanto sin Word como con Word real en Windows.

## Alcance mínimo

1. Un ejercicio original en español sobre formato de texto.
2. Un DOCX inicial generado y versionado con el ejercicio.
3. UI mínima para mostrar instrucciones, elegir modo, abrir la copia de trabajo y evaluar.
4. Evaluador portable que inspecciona OOXML.
5. Adaptador COM de Word para Windows, aislado tras una interfaz común.
6. Resultado local por criterio con evidencia.
7. Pruebas automatizadas y documentación de límites.

No incluye aún temporizador de examen completo, login, nube, sincronización, editor de documentos propio ni un catálogo amplio.

## Tarea de práctica

El documento contiene un párrafo con la frase objetivo **“Informe de resultados”**. La instrucción pide aplicar negrita a esa frase y establecer el tamaño de fuente en 16 pt, sin cambiar el texto del párrafo. La plantilla de referencia, el texto objetivo y los criterios deben generarse de forma reproducible.

## Criterios de evaluación

- **C1 — Integridad del contenido:** la frase objetivo aparece exactamente una vez en el contenido esperado.
- **C2 — Negrita:** todo el texto objetivo tiene negrita efectiva.
- **C3 — Tamaño:** todo el texto objetivo tiene tamaño de fuente de 16 pt, considerando el formato directo y las reglas de herencia de estilo que el adaptador soporte.
- **C4 — No afectar texto vecino:** el texto inmediatamente anterior y posterior conserva las propiedades de formato de referencia definidas por la plantilla.
- **C5 — No alterar la plantilla:** el archivo original de referencia permanece idéntico; el usuario trabaja sobre una copia.

Para la primera versión, C2–C4 deben tener una implementación portable claramente delimitada. Si el evaluador no puede resolver una propiedad heredada con fiabilidad, devuelve `not_verifiable` con motivo, no `pass`.

## Contrato de estados

- `pass`: el criterio está satisfecho con evidencia.
- `fail`: la evidencia demuestra que no está satisfecho.
- `not_verifiable`: el adaptador no puede verificarlo con fiabilidad.
- `error`: la comprobación no pudo completarse.

No asignar puntos a `not_verifiable` como si estuviera aprobado. La regla de puntuación del ejercicio debe ser determinista, estar versionada y explicar cómo trata criterios no verificables.

## Criterios de aceptación funcional

### AC-01 — Inicio offline

**Dado** que el equipo no tiene conexión, **cuando** el usuario inicia la aplicación, **entonces** puede abrir el ejercicio incluido.

### AC-02 — Copia de trabajo

**Dado** el documento de referencia, **cuando** comienza el intento, **entonces** la aplicación crea/selecciona una copia de trabajo y no modifica el original.

### AC-03 — Modo portable

**Dado** un sistema Linux o Windows sin Word, **cuando** el usuario evalúa el DOCX, **entonces** el evaluador portable termina y muestra los estados C1–C5, la puntuación y la evidencia.

### AC-04 — Detección explícita de Word

**Dado** un sistema sin Word o sin COM, **cuando** se inicia la sesión, **entonces** el modo Word COM aparece como no disponible y la aplicación no se bloquea ni lo selecciona silenciosamente.

### AC-05 — Modo Word real

**Dado** Windows con Microsoft Word de escritorio compatible, **cuando** el usuario elige el modo Word COM, **entonces** el adaptador evalúa el documento mediante Word y registra `evaluator=word_com`.

### AC-06 — Corrección positiva

**Dado** un documento que satisface todos los criterios verificables, **cuando** se evalúa, **entonces** C1–C5 reciben los estados esperados y no hay falsos fallos.

### AC-07 — Corrección negativa

**Dado** un documento donde la frase objetivo no está en negrita o no está a 16 pt, **cuando** se evalúa, **entonces** el criterio correspondiente falla y se muestra una explicación concreta.

### AC-08 — Texto fragmentado

**Dado** que Word divide la frase objetivo en varios `runs`, **cuando** se evalúa, **entonces** la lógica identifica la frase a través de los fragmentos y comprueba todas las partes requeridas; no presupone un único run.

### AC-09 — Archivo inválido

**Dado** un DOCX corrupto o un ZIP no válido, **cuando** se evalúa, **entonces** la aplicación muestra un error controlado, no ejecuta contenido activo y permanece utilizable.

### AC-10 — Evidencia de resultados

**Dado** un intento finalizado, **cuando** se muestra el resultado, **entonces** quedan registrados el ID y versión del ejercicio, adaptador, fecha local, estados por criterio y versión de reglas.

### AC-11 — Reproducibilidad

**Dado** el mismo archivo y la misma versión del evaluador, **cuando** se ejecuta la prueba varias veces, **entonces** los resultados y puntuación son idénticos, excepto campos deliberadamente variables como la fecha.

### AC-12 — Regresión automatizada

Las pruebas unitarias cubren documento correcto, negrita ausente, tamaño incorrecto, texto objetivo ausente, texto fragmentado, formato heredado, DOCX corrupto y archivo excesivo. Las pruebas COM se ejecutan en un runner Windows configurado con Word; si no existe ese entorno, el resultado se registra como no ejecutado, nunca como aprobado.

## Pruebas adversariales

- ZIP con demasiadas entradas o tamaño descomprimido excesivo.
- XML inesperado y referencias externas.
- Ruta con traversal o nombre de archivo con caracteres Unicode.
- Documento de solo lectura o con permisos insuficientes.
- Word con cuadro de diálogo modal, proceso bloqueado o instancia ajena ya abierta.
- Cancelación durante la evaluación.
- Pulsar evaluar dos veces rápidamente.
- Documento que contiene varias coincidencias de la frase.
- Formato parcial aplicado solo a una parte de la frase.

## Evidencia de salida requerida

Actualizar `QA_RESULTS.md` con sistema operativo, versiones, comando de prueba, resultado por caso, limitaciones conocidas y defectos. No declarar el vertical slice terminado hasta que se cumplan los gates portable y Word COM en sus respectivos entornos.
