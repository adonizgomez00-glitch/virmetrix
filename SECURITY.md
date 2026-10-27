# Seguridad y privacidad

**Estado:** modelo de amenazas preliminar.

## Activos

- Documentos de práctica del estudiante.
- Plantillas originales.
- Historial de intentos y puntuaciones.
- Archivos de ejercicios y configuración.
- Integridad de la puntuación y del catálogo.

## Límites de confianza

- Documentos DOCX elegidos por el usuario.
- Paquetes de ejercicios descargados o importados.
- Word y el proceso COM.
- Archivos temporales y directorios de trabajo.
- Dependencias y paquete instalador.

## Amenazas y controles

### DOCX malicioso o corrupto

DOCX es un contenedor ZIP con XML y posibles relaciones/elementos activos. Tratarlo como entrada no confiable. Limitar tamaño comprimido y descomprimido, cantidad de entradas, profundidad y tiempo de análisis. No ejecutar macros, scripts ni contenido incrustado.

### Traversal y sobrescritura

Resolver y validar rutas dentro del directorio de trabajo autorizado. No confiar en nombres de archivo ni rutas contenidas en metadatos del ejercicio. Crear copias de trabajo; preservar la plantilla original.

### XML inseguro

Desactivar resolución de entidades externas y acceso de red en el parser. Usar versiones mantenidas de las bibliotecas y añadir pruebas adversariales.

### Automatización COM

No ejecutar Word con privilegios elevados. Abrir solo archivos de práctica controlados por la aplicación, manejar cuadros de diálogo y timeouts, y limpiar procesos temporales sin terminar instancias de Word ajenas al simulador.

### Privacidad

La aplicación debe ser local-first. No subir documentos ni telemetría por defecto. Evitar almacenar el contenido del documento en logs. Permitir borrar el historial local y explicar dónde se guarda.

### Cadena de suministro

Fijar y revisar dependencias, comprobar licencias y construir instaladores reproducibles cuando sea viable. No incluir secretos en repositorios ni en logs.

## No objetivos de seguridad

- No es un entorno de aislamiento para abrir documentos arbitrarios con macros habilitadas.
- No sustituye un antivirus ni la seguridad del sistema operativo.
- No debe prometer protección absoluta frente a todos los archivos maliciosos.

## Revisión antes de publicar

- [ ] Modelo de amenazas actualizado.
- [ ] Dependencias y licencias revisadas.
- [ ] Casos de archivo malformado probados.
- [ ] Rutas temporales validadas.
- [ ] Logs inspeccionados para detectar exposición de datos.
- [ ] Aviso de privacidad y ubicación de datos documentados.

## Controles adicionales para automatización y archivos

- Establecer límites de tamaño del archivo comprimido, tamaño total descomprimido y cantidad de entradas ZIP.
- Rechazar rutas internas inesperadas y evitar extracción de paquetes DOCX al sistema de archivos cuando no sea necesaria.
- Aplicar límites de tiempo al análisis y permitir cancelación controlada.
- En COM, registrar el PID o identificador de la instancia que la aplicación crea cuando sea posible; cerrar solo recursos propios.
- No desactivar controles de seguridad de Word globalmente ni cambiar opciones persistentes del usuario sin consentimiento.
