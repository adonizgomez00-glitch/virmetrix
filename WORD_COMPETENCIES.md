# Mapa de competencias — MOS Word Associate

**Estado:** esquema preliminar; no es una lista oficial definitiva.  
**Regla:** antes de publicar ejercicios como alineados a un examen específico, verificar los objetivos públicos vigentes de Microsoft para la versión elegida.

## Dominios propuestos

1. Administrar documentos y opciones.
2. Insertar y dar formato a texto, párrafos y secciones.
3. Administrar tablas y listas.
4. Crear y administrar referencias.
5. Insertar y dar formato a elementos gráficos.
6. Revisar, colaborar y proteger documentos, cuando corresponda a la versión objetivo.

Estos dominios sirven para organizar el contenido; su nombre, alcance y peso deben reconciliarse con los objetivos oficiales vigentes antes de construir el banco final.

## Modelo de ejercicio

Cada ejercicio debe incluir:

- `exercise_id`: identificador estable.
- `version`: versión del ejercicio.
- `title` y `language`.
- `competency_id` y referencia al objetivo público.
- `difficulty`.
- `estimated_minutes`.
- `evaluation_modes`: `docx_portable`, `word_com` o ambos.
- `starting_document`: ruta a la plantilla original.
- `instructions`: instrucciones originales para el estudiante.
- `criteria`: criterios de evaluación verificables.
- `hints`: pistas opcionales, no visibles en modo examen.
- `explanation`: explicación de la técnica al terminar.
- `rights`: origen y licencia de todos los recursos incluidos.

## Estados de criterio

- `pass`: evidencia suficiente de que se cumple el requisito.
- `fail`: evidencia suficiente de que no se cumple.
- `not_verifiable`: el adaptador no puede comprobar el criterio con fiabilidad.
- `error`: el evaluador no pudo completar la comprobación.

## Ejercicio piloto sugerido

**ID:** `word-format-001`  
**Tema:** formato de texto.

**Instrucción de ejemplo:** aplicar negrita a una frase indicada y cambiar el tamaño de fuente de un fragmento marcado, conservando el resto del texto.

**Criterios iniciales:**

- C1: el texto objetivo existe en el documento.
- C2: el fragmento objetivo tiene el formato de carácter requerido.
- C3: el texto fuera del objetivo conserva el formato de referencia cuando el criterio lo exige.

**Advertencia técnica:** los fragmentos de texto pueden dividirse en varios `runs` dentro de OOXML. El evaluador no debe asumir que una frase completa reside en un solo `run`. Los criterios exactos se definirán después de crear la plantilla y validar el comportamiento en Word real.

## Control de versión

Registrar la versión de los objetivos de examen consultados y la fecha de verificación. Si Microsoft actualiza los objetivos, realizar un análisis de impacto y versionar el mapa; no cambiar silenciosamente ejercicios publicados.

## Versión de referencia y pesos publicados

**Referencia seleccionada:** MOS Word Associate (Office 2019), examen MO-100. Fuente oficial: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/ (consultada el 2026-10-27).

| Dominio publicado | Peso |
|---|---:|
| Administrar documentos | 20–25 % |
| Insertar y dar formato a texto, párrafos y secciones | 20–25 % |
| Administrar tablas y listas | 15–20 % |
| Crear y administrar referencias | 5–10 % |
| Insertar y dar formato a elementos gráficos | 15–20 % |
| Administrar la colaboración en documentos | 5–10 % |

Estos dominios son la base inicial. Los subobjetivos y el diseño de ejercicios deben contrastarse con el archivo oficial “Exam MO-100 skills outline” antes de afirmar cobertura completa. Versionar el mapa y anotar fecha de revisión.

## Criterio para planificar cobertura

Para evitar que el banco inicial sobre-represente el formato de texto, planificar el catálogo de forma que todos los dominios tengan ejercicios. Los pesos orientan el volumen a medio plazo; no deben traducirse directamente en una puntuación oficial del examen.

## Corrección de referencia oficial

La referencia es **Microsoft Office Specialist: Word Associate (Office 2019), MO-100**, según Microsoft Learn: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

Los seis dominios publicados son: administrar documentos; insertar y dar formato a texto, párrafos y secciones; administrar tablas y listas; crear y administrar referencias; insertar y dar formato a elementos gráficos; y administrar la colaboración en documentos.

No conservar porcentajes procedentes de una versión de examen distinta sin verificar que aparecen en el esquema MO-100. Antes de fijar ponderaciones y subobjetivos, revisar el documento oficial “MO-100 Exam Skills Outline”.


## Pesos confirmados en el esquema oficial MO-100

Según el PDF oficial “MO-100: Microsoft Word (Office 2019) – Skills Measured”:

| Dominio | Peso |
|---|---:|
| Administrar documentos | 20–25 % |
| Insertar y dar formato a texto, párrafos y secciones | 20–25 % |
| Administrar tablas y listas | 15–20 % |
| Crear y administrar referencias | 5–10 % |
| Insertar y dar formato a elementos gráficos | 15–20 % |
| Administrar la colaboración en documentos | 5–10 % |

Fuente oficial: https://arch-center.azureedge.net/Learning/Credentials/mo-100-microsoft-word-2019-skills-measured.pdf

### Subobjetivos a representar en el catálogo

1. **Administrar documentos:** navegación, búsqueda, vínculos internos, formato de documento, encabezados y pies, fondos de página, formatos alternativos, propiedades, impresión, compartir electrónicamente, inspección de información personal, accesibilidad y compatibilidad.
2. **Texto, párrafos y secciones:** buscar y reemplazar, símbolos, efectos de texto, copiar formato, espaciado, sangrías, estilos integrados, borrar formato, columnas y saltos de página/sección/columna.
3. **Tablas y listas:** convertir texto y tablas, crear tablas, ordenar datos, márgenes de celda, combinar/dividir, ajustar tamaño, dividir tablas, repetir encabezados y administrar viñetas/numeración/niveles.
4. **Referencias:** notas al pie y al final, fuentes de citas, citas bibliográficas, tablas de contenido y bibliografías.
5. **Elementos gráficos:** formas, imágenes, modelos 3D, SmartArt, capturas, cuadros de texto, efectos/estilos, quitar fondos, posición, ajuste de texto y texto alternativo.
6. **Colaboración:** comentarios, respuestas, resolver/eliminar comentarios, control de cambios, revisar, aceptar/rechazar cambios y bloquear/desbloquear seguimiento.

La lista oficial advierte que los puntos son ejemplos de cómo se evalúa cada aptitud y no constituyen una lista exhaustiva. VirMetrix debe versionar esta matriz y usar tareas originales.
