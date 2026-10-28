# VirMetrix

**Estado:** especificación inicial — no implementado  
**Nombre de trabajo:** VirMetrix  
**Objetivo:** crear una aplicación gratuita de escritorio para practicar tareas de Microsoft Word orientadas a la certificación MOS Word Associate.

## Principios

- Windows es la plataforma principal para ejecutar la experiencia completa con Microsoft Word de escritorio.
- Linux es una plataforma de desarrollo y pruebas del núcleo; se debe poder ejecutar y probar sin Microsoft Word instalado.
- La evaluación debe verificar resultados observables, no solo respuestas de opción múltiple.
- La aplicación debe distinguir entre evaluación con Word real y evaluación de un archivo `.docx` sin Word.
- No copiar preguntas, simulacros, interfaz, marcas gráficas ni materiales propietarios de GMetrix. Crear ejercicios originales basados en objetivos públicos de certificación.
- No afirmar afiliación, patrocinio ni certificación oficial de Microsoft.
- Mantener el proyecto gratuito; la licencia del código y la licencia de los materiales didácticos se decidirán antes de la primera publicación.

## Alcance del MVP

1. Aplicación de escritorio local.
2. Catálogo de ejercicios originales, organizados por objetivo.
3. Instrucciones paso a paso y modo de examen cronometrado.
4. Apertura/selección de archivos DOCX de práctica.
5. Evaluador con dos adaptadores:
   - **Word Windows:** automatización de Microsoft Word mediante COM, cuando esté instalado.
   - **DOCX portable:** inspección de la estructura OOXML del archivo sin requerir Word.
6. Resultado por criterio, explicación de errores y puntuación.
7. Historial local de intentos y progreso.
8. Funcionamiento sin conexión para el núcleo y el contenido instalado.

## Fuera del MVP

- Cuenta en la nube, clases virtuales, sincronización, pagos o publicidad.
- Reproducción exacta de la interfaz de GMetrix.
- Garantizar equivalencia con el motor de puntuación oficial de Microsoft.
- Evaluar Word para la Web automáticamente.
- Compatibilidad universal con todos los formatos y versiones de Word.

## Documentación viva SAQI

- `CONTEXT.md`: reglas operativas y contexto del proyecto.
- `PROJECT_STATE.md`: estado actual, decisiones tomadas y siguiente tarea.
- `ARCHITECTURE.md`: componentes, límites y contratos.
- `QA_RESULTS.md`: resultados comprobables de QA y defectos.
- `SESSION.md`: checkpoint de la sesión de trabajo.
- `REQUIREMENTS.md`: requisitos funcionales y no funcionales.
- `WORD_COMPETENCIES.md`: mapa de objetivos y ejercicios.
- `ROADMAP.md`: hitos del MVP.
- `TEST_STRATEGY.md`: estrategia de pruebas y QA adversarial.
- `SECURITY.md`: modelo de amenazas y controles.
- `DECISIONS.md`: registro de decisiones de arquitectura.

## Estado de verdad

Estos documentos son una propuesta de requisitos. No implican que el software, los adaptadores o las pruebas ya estén implementados. Los resultados solo se marcarán como aprobados cuando exista evidencia reproducible.

## Primera tarea recomendada

Completar la Fase 1 de SAQI: confirmar el alcance de la primera versión de Word, los sistemas operativos de Windows objetivo y la versión de examen MOS que se usará como referencia. Después, crear un vertical slice que evalúe una tarea DOCX de formato de texto con y sin Word instalado.

## Objetivo de certificación de referencia

Referencia inicial confirmada: **Microsoft Office Specialist: Word Associate (Office 2019), examen MO-100**. Microsoft Learn indica seis dominios: administrar documentos; insertar y dar formato a texto, párrafos y secciones; administrar tablas y listas; crear y administrar referencias; insertar y dar formato a elementos gráficos; y administrar la colaboración en documentos. Verificar el esquema oficial de aptitudes antes de fijar el contenido detallado de cada ejercicio.

## Decisiones del MVP confirmadas

- Objetivo de certificación: versión vigente de Word Associate, actualmente identificada como MO-100 para Office 2019 según la página oficial consultada el 27-10-2026.
- Distribución: instalador `.exe` para estudiantes y código fuente multiplataforma para desarrollo.
- Conectividad: el MVP debe funcionar 100 % offline en sus funciones esenciales; no depender de cuentas, servicios web, telemetría ni APIs remotas.
- Plataformas: Windows para la experiencia completa con Word real; Linux para desarrollo y evaluación portable sin Word.

## Stack técnico seleccionado

El stack inicial queda aprobado para el MVP:
- Python para lógica y aplicación.
- PySide6 para la interfaz de escritorio.
- `python-docx` para operaciones DOCX compatibles.
- Evaluador OOXML separado para criterios que requieren inspección más detallada del paquete DOCX.
- `pywin32` para automatización COM de Microsoft Word en Windows.
- SQLite para historial y datos locales.
- `pytest` para pruebas.
- PyInstaller como candidato inicial para generar el instalador/ejecutable Windows, sujeto a prueba de empaquetado.

Nombre del producto: **VirMetrix**. SAQI es la metodología de desarrollo, no el nombre de la aplicación.

## Corrección de objetivo de certificación

VirMetrix prepara para **Microsoft Office Specialist: Word Associate (Office 2019), examen MO-100**. No utilizar MO-110 como objetivo de este proyecto. Fuente oficial: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

La página oficial consultada indica 50 minutos, incluye español entre los idiomas y no muestra una fecha de retirada. Confirmar nuevamente antes de cada publicación.

## Empaquetado (ITER-005)

Generar ejecutable con PyInstaller:

Linux:
```bash
./scripts/build_linux.sh
```

Windows (PowerShell):
```powershell
.\scripts\build_windows.ps1
```

Salida: `dist/VirMetrix/` (ejecutable + recursos). Para un instalador `.exe` convencional, usar Inno Setup o NSIS sobre `dist/VirMetrix/`.

Notas:
- El bundle incluye `exercises/` automáticamente.
- El historial SQLite se guarda en `%USERPROFILE%\.virmetrix\history.sqlite` (Windows) o `~/.virmetrix/history.sqlite` (Linux).
- En Windows, el modo `word_com` requiere Microsoft Word de escritorio.
