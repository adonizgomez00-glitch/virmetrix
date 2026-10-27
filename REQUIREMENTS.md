# Requisitos del producto

**Estado:** borrador inicial.  
**Prioridad:** P0 = necesaria para el MVP; P1 = importante; P2 = posterior.

## Funcionales

### FR-001 — Inicio local (P0)

La aplicación debe iniciar sin conexión a Internet después de instalarse.

**Aceptación:** iniciar y abrir el catálogo con la red desactivada.

### FR-002 — Catálogo de ejercicios (P0)

Mostrar ejercicios agrupados por competencia, dificultad y modo de evaluación.

**Aceptación:** el usuario puede abrir un ejercicio y consultar instrucciones, archivo inicial y criterios generales.

### FR-003 — Sesión de práctica (P0)

Permitir iniciar, pausar cuando el modo lo permita, finalizar y reiniciar un intento.

**Aceptación:** el estado de cada intento queda definido y no se duplican resultados al pulsar finalizar repetidamente.

### FR-004 — Archivos de práctica (P0)

Proporcionar un documento de inicio y trabajar sobre una copia aislada.

**Aceptación:** el original de la plantilla no se modifica durante una sesión.

### FR-005 — Evaluación portable (P0)

Evaluar criterios compatibles inspeccionando DOCX sin Microsoft Word.

**Aceptación:** cada criterio devuelve estado y evidencia; los criterios no comprobables se marcan `not_verifiable`.

### FR-006 — Evaluación Word real (P0)

En Windows con Word de escritorio disponible, ejecutar la evaluación mediante el adaptador COM.

**Aceptación:** detectar disponibilidad, informar si Word no está disponible y evitar que la UI se bloquee durante la evaluación.

### FR-007 — Modo explícito de evaluación (P0)

Mostrar si el intento usa `word_com` o `docx_portable`.

**Aceptación:** el modo queda guardado con el resultado y no cambia silenciosamente a mitad del intento.

### FR-008 — Puntuación y retroalimentación (P0)

Mostrar puntuación, criterios aprobados/fallidos y explicación de los fallos.

**Aceptación:** la puntuación se calcula con reglas deterministas y versionadas.

### FR-009 — Historial local (P1)

Guardar fecha, ejercicio, duración, puntuación y modo de evaluación.

**Aceptación:** cerrar y reabrir la aplicación no elimina el historial.

### FR-010 — Simulacro cronometrado (P1)

Permitir una sesión con límite de tiempo y varias tareas.

**Aceptación:** el temporizador no reinicia accidentalmente al cambiar de pantalla.

### FR-011 — Accesibilidad e idioma (P1)

Interfaz inicial en español, navegación por teclado y contraste legible.

### FR-012 — Exportar resultados (P2)

Exportar un informe local en un formato abierto sin datos personales innecesarios.

## No funcionales

- **NFR-001 Portabilidad:** núcleo y evaluador portable ejecutables en Linux y Windows.
- **NFR-002 Compatibilidad:** integración COM solo en Windows con Word compatible instalado.
- **NFR-003 Honestidad:** indicar los límites de cada modo y nunca presentar el modo portable como equivalencia oficial.
- **NFR-004 Privacidad:** documentos y resultados permanecen locales por defecto.
- **NFR-005 Robustez:** un archivo DOCX corrupto o malicioso no debe cerrar inesperadamente toda la aplicación.
- **NFR-006 Mantenibilidad:** criterios, contenido y código de UI separados.
- **NFR-007 Pruebas:** reglas de puntuación y criterios críticos cubiertos por pruebas automatizadas.
- **NFR-008 Seguridad:** no ejecutar macros, scripts ni enlaces incrustados de los documentos de práctica.
- **NFR-009 Distribución:** generar un paquete instalable para Windows; la distribución Linux de desarrollo puede ser distinta.
- **NFR-010 Coste:** no requerir servicios pagados para el flujo esencial.

## Fuera de alcance inicial

- Replicar el algoritmo de puntuación de Microsoft.
- Certificar que un usuario aprobará el examen real.
- Sincronización de nube o analítica remota.
- Automatización de Word para macOS.

## Requisitos de distribución y conectividad (confirmados)

### FR-013 — Instalador Windows (P0)

Debe existir un instalador `.exe` para estudiantes.

**Aceptación:** en una máquina Windows de prueba limpia, el paquete instala y desinstala la aplicación según la documentación. La ausencia de Word no debe impedir iniciar el modo portable.

### FR-014 — Código fuente multiplataforma (P0)

El repositorio debe contener código fuente y pasos de desarrollo para Linux y Windows.

**Aceptación:** un desarrollador puede crear un entorno de desarrollo siguiendo instrucciones documentadas en ambos sistemas, y ejecutar las pruebas del núcleo/evaluador portable.

### NFR-011 — Offline completo (P0)

El flujo esencial del MVP debe funcionar sin Internet.

**Aceptación:** con la red desactivada, se puede iniciar la aplicación, consultar ejercicios incluidos, practicar, evaluar, ver retroalimentación e historial.

### NFR-012 — Sin servicios obligatorios (P0)

No requerir cuentas, telemetría, APIs ni backend remoto para las funciones del MVP.

**Aceptación:** inspección de dependencias y pruebas de red confirman que las acciones esenciales no necesitan servicios remotos.

## Objetivo de certificación seleccionado

Referencia: **MOS Word Associate (Office 2019), examen MO-100**, según la página oficial Microsoft Learn consultada el 2026-10-27. Los seis dominios y sus pesos publicados son:
- Administrar documentos: 20–25 %.
- Insertar y dar formato a texto, párrafos y secciones: 20–25 %.
- Administrar tablas y listas: 15–20 %.
- Crear y administrar referencias: 5–10 %.
- Insertar y dar formato a elementos gráficos: 15–20 %.
- Administrar la colaboración en documentos: 5–10 %.

Fuente: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

Antes de cerrar el mapa detallado de tareas, validar el esquema oficial descargable de aptitudes MO-100. El producto prepara para practicar; no es un producto oficial ni garantiza aprobar el examen.

## Restricciones técnicas adicionales

- La interfaz debe seguir respondiendo durante evaluaciones largas; ejecutar análisis de archivos y llamadas COM fuera del hilo principal de UI.
- El adaptador COM no debe cerrar instancias de Word ajenas a VirMetrix ni guardar cambios en documentos originales.
- El evaluador portable no debe usar Word, LibreOffice ni servicios remotos como dependencia obligatoria.
- Los resultados deben incluir el modo de evaluación y la versión del evaluador.
- Un resultado `not_verifiable` no se contabiliza como criterio aprobado. La política de puntuación debe declararse en el ejercicio.
