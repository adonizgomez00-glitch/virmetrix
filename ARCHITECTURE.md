# Architecture — VirMetrix

**Estado:** propuesta inicial; sujeta a validación con un vertical slice.

## Objetivos arquitectónicos

- Separar la interfaz, la lógica de evaluación y los mecanismos específicos del sistema operativo.
- Permitir el modo portable sin Microsoft Word.
- Integrar Word real en Windows sin que el núcleo dependa de COM.
- Hacer que cada ejercicio y criterio sean versionables y testeables.
- Operar localmente y sin conexión después de instalar la aplicación y los paquetes de ejercicios.

## Arquitectura propuesta

```text
Desktop UI
   |
Application / Session Controller
   |--------------------------|
Exercise Catalog          Attempt / Progress Store
   |
Evaluation Orchestrator
   |-------------------------------|
Portable DOCX Evaluator        Word COM Adapter (Windows)
   |                               |
OOXML / DOCX package             Microsoft Word
   |
Criterion Results
   |
Scoring + Feedback + Report
```

## Stack candidato — no aprobado todavía

- Python 3.x.
- PySide6 para la interfaz de escritorio.
- `python-docx` para leer y crear documentos DOCX sencillos.
- `zipfile` + `lxml` u otra biblioteca OOXML evaluada para inspecciones más específicas.
- `pywin32` exclusivamente en Windows para automatización COM de Word.
- SQLite para historial local.
- `pytest` para pruebas unitarias e integración.
- PyInstaller u otra herramienta evaluada para empaquetar Windows.

La selección definitiva se hará tras comprobar compatibilidad, licencias, accesibilidad, tamaño del instalador y facilidad de desarrollo en Linux.

## Límites de los componentes

### Interfaz de escritorio

Muestra catálogo, instrucciones, temporizador, estado de sesión, resultados y accesibilidad. No debe contener reglas de puntuación directamente.

### Catálogo de ejercicios

Cada ejercicio tiene identificador estable, versión, idioma, objetivo de competencia, archivos iniciales, instrucciones, criterios y metadatos de dificultad. Los criterios deben ser declarativos cuando sea viable.

### Orquestador de evaluación

Selecciona el adaptador disponible y expone capacidades explícitas. No debe cambiar silenciosamente de evaluación Word a evaluación portable durante un intento.

### Evaluador portable DOCX

Lee el paquete OOXML sin iniciar Word. Comprueba solo criterios cuya representación en el documento sea fiable. Debe diferenciar `pass`, `fail`, `not_verifiable` y `error`.

### Adaptador Word COM

Se ejecuta solo en Windows con Microsoft Word de escritorio disponible. Debe tener timeout, manejo de procesos huérfanos, limpieza de documentos temporales y pruebas aisladas. No debe automatizar acciones destructivas fuera de los archivos de práctica.

### Persistencia local

Guarda intentos, tiempos, resultados por criterio y preferencias. En el MVP no se requiere cuenta ni sincronización en la nube.

## Contrato de resultado propuesto

```json
{
  "exercise_id": "word-format-001",
  "evaluator": "docx_portable",
  "status": "completed",
  "criteria": [
    {
      "criterion_id": "C1",
      "status": "pass",
      "evidence": "Se detectó el estilo de carácter esperado."
    }
  ],
  "score": {
    "earned": 1,
    "possible": 1
  }
}
```

Estados permitidos por criterio: `pass`, `fail`, `not_verifiable`, `error`. No convertir `not_verifiable` en `pass`.

## Principios de seguridad

- Operar sobre copias de documentos de práctica.
- No ejecutar macros ni contenido activo procedente de archivos de usuario.
- Validar rutas y limitar acceso al sistema de archivos.
- No abrir archivos con privilegios elevados.
- No recoger datos personales innecesarios.
- No cargar documentos del estudiante a servicios externos.
- Registrar errores sin volcar contenido completo de documentos privados.



## ADRs (Architecture Decision Records)

Ver `docs/adr/`:

- ADR-001: Evaluador portable basado en inspección OOXML.
- ADR-002: Interfaz común de adaptador de evaluación.
- ADR-003: Adaptador Word COM solo en Windows con importación bajo demanda.
- ADR-004: Estados de criterio y regla de puntuación.
- ADR-005: Trabajo sobre copia; plantilla original inmutable.
- ADR-006: Seguridad del parsing OOXML.

## Contratos de puertos (Domain Ports)

Ver `src/domain/ports/evaluation.py`:

| Puerto | Interface | Implementación (Adapter) |
|---|---|---|
| EvaluationAdapter | `is_available()`, `evaluate()` | `PortableDocxEvaluator`, `WordComAdapter` |
| ExerciseCatalog | `load()`, `list()` | `JsonExerciseCatalog` (a definir) |
| AttemptStore | `save()`, `list()` | Diferido a ITER-002 (SQLite) |

## Estructura de código propuesta (ITER-001)

```text
src/
├── domain/
│   ├── models/          # ExerciseMetadata, CriterionResult, EvaluationResult
│   └── ports/           # EvaluationAdapter, ExerciseCatalog
├── application/
│   └── services/        # EvaluationOrchestrator, ScoringService, WorkCopyManager
├── infrastructure/
│   ├── evaluators/      # PortableDocxEvaluator (OOXML), WordComAdapter (Windows)
│   ├── catalog/         # JsonExerciseCatalog
│   └── persistence/     # (ITER-002)
├── ui/                  # PySide6 views/controllers
└── tests/
    ├── unit/
    ├── integration/
    └── adversarial/
```

El núcleo no importa `win32com`/`pythoncom`; el adaptador COM importa bajo demanda.

## Decisiones abiertas

- UI: PySide6 frente a alternativas.
- Paquete y estrategia de actualizaciones.
- Modelo declarativo de criterios.
- Formato de distribución de ejercicios.
- Compatibilidad exacta de Word y Windows.

## Decisión de distribución y operación offline

El MVP debe operar localmente sin requerir conexión a Internet. No usar autenticación, telemetría, APIs remotas ni servicios en la nube para el flujo esencial. Catálogo, documentos de práctica, reglas de evaluación, historial y ayuda deben venir incluidos o poder importarse localmente.

### Distribución

- Windows: paquete instalable `.exe` para estudiantes.
- Desarrollo: código fuente y pasos reproducibles para Linux y Windows.
- Word COM: dependencia opcional de Microsoft Word de escritorio en Windows; su ausencia no impide usar el evaluador portable.
- No asumir que el instalador puede incluir Microsoft Word; la licencia de Word es independiente.

### Implicación offline

La verificación de versiones de examen y la actualización de ejercicios se realizan durante el desarrollo/publicación, no en tiempo de ejecución. Mostrar la versión/fecha de referencia en los metadatos locales del contenido.

## Stack seleccionado para el MVP

- **Python:** lenguaje principal.
- **PySide6:** interfaz de escritorio.
- **python-docx:** lectura y generación de DOCX para operaciones compatibles.
- **Evaluador OOXML:** inspección explícita de XML y relaciones del paquete para criterios de evaluación; no depender solo de `python-docx`.
- **pywin32:** automatización COM de Word, únicamente en Windows.
- **SQLite:** persistencia local.
- **pytest:** pruebas automatizadas.
- **PyInstaller:** candidato para empaquetar la aplicación Windows; verificar si la distribución elegida produce un instalador adecuado o si hace falta un empaquetador adicional.

El evaluador portable debe ser independiente de `pywin32`; importar componentes Windows solo dentro del adaptador y protegerlos mediante detección de plataforma/capacidad.

## Contratos técnicos que deben respetarse

- El núcleo no importa `win32com` ni `pythoncom` al importarse. El adaptador COM realiza importaciones bajo demanda y solo en Windows.
- La UI no implementa lógica de puntuación; invoca un servicio de aplicación y presenta resultados.
- El evaluador portable y el evaluador COM implementan una interfaz común, pero pueden devolver distintos estados de verificabilidad.
- El evaluador recibe una ruta de archivo de trabajo y metadatos de ejercicio; devuelve un resultado estructurado sin modificar el documento.
- La aplicación no debe matar procesos `WINWORD.EXE` de forma global, porque podrían pertenecer a sesiones ajenas del usuario.
- PyInstaller genera un ejecutable empaquetado, pero no necesariamente un instalador convencional. El instalador `.exe` se debe validar por separado y puede requerir una herramienta de empaquetado adicional.
