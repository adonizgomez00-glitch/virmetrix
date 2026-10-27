# VirMetrix — Plan de implementación técnica

**Estado:** plan de trabajo; no equivale a código implementado.

## Orden de implementación

### Iteración 1 — Núcleo y evaluador portable

- Crear paquete Python `virmetrix`.
- Definir modelos tipados de ejercicio, criterio y resultado.
- Crear un lector DOCX seguro con límites ZIP/XML.
- Implementar la evaluación de `word-format-001`.
- Añadir pruebas de documentos correctos, incorrectos, corruptos y texto dividido en varios runs.
- Ejecutar en Linux y Windows sin Word.

**Salida:** CLI o prueba de integración reproducible que evalúe un DOCX sin UI.

### Iteración 2 — Interfaz mínima

- PySide6 muestra un ejercicio local y sus instrucciones.
- El usuario selecciona o abre una copia de trabajo.
- La evaluación se ejecuta fuera del hilo de UI.
- Mostrar estado, puntuación, evidencia y adaptador utilizado.
- Guardar historial local en SQLite.

**Salida:** un ejercicio puede iniciarse, evaluarse y revisarse en modo portable.

### Iteración 3 — Adaptador Word COM

- Detectar Windows y disponibilidad de Word.
- Crear un adaptador aislado que utilice `pywin32`.
- Evitar cerrar procesos Word ajenos.
- Añadir timeout, cancelación, limpieza de recursos propios y mensajes de error.
- Probar con Word instalado en una máquina Windows real.

**Salida:** el mismo ejercicio puede evaluarse mediante Word real y registra `evaluator=word_com`.

### Iteración 4 — Empaquetado

- Probar PyInstaller para generar ejecutable.
- Elegir una herramienta de instalador si se requiere instalación/desinstalación convencional.
- Validar en Windows limpio sin Word y Windows con Word.
- Documentar instalación, desinstalación y directorios de datos.

### Iteración 5 — Gate SAQI

- Revisar requisitos y trazabilidad.
- Ejecutar pruebas automatizadas y QA adversarial.
- Corregir defectos y añadir pruebas de regresión.
- Actualizar `PROJECT_STATE.md`, `QA_RESULTS.md` y `SESSION.md`.
- No declarar release si quedan defectos críticos/altos abiertos.

## Regla para el agente de programación

Trabajar en una sola iteración a la vez. Antes de cambiar arquitectura, actualizar `DECISIONS.md`. No generar módulos grandes de golpe sin pruebas ejecutables. Al terminar cada tarea, informar archivos cambiados, comandos ejecutados, pruebas aprobadas/fallidas/no ejecutadas y riesgos restantes.
