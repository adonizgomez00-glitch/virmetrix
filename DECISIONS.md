# Registro de decisiones (ADR)

## ADR-001 — Aplicación de escritorio integrada con Word

- **Estado:** aceptada por el usuario.
- **Fecha:** 2026-10-27.
- **Decisión:** la experiencia principal será una aplicación de escritorio integrada con Microsoft Word real cuando esté disponible.
- **Motivo:** el objetivo es practicar tareas prácticas de Word y evaluar el documento producido.
- **Consecuencias:** el adaptador de Word debe aislarse del núcleo; la integración completa de Word se prueba en Windows.
- **Riesgos:** COM depende de Windows y de una instalación compatible de Word.

## ADR-002 — Modo de práctica sin Word instalado

- **Estado:** aceptada por el usuario.
- **Fecha:** 2026-10-27.
- **Decisión:** el producto también debe permitir practicar sin Microsoft Word instalado.
- **Implementación propuesta:** evaluar archivos DOCX mediante un adaptador portable basado en OOXML.
- **Consecuencia:** los resultados deben identificar el adaptador utilizado y sus limitaciones. Los criterios no verificables no pueden marcarse como aprobados.
- **Riesgo:** no todos los comportamientos visuales o interactivos de Word se pueden inferir con precisión desde DOCX.

## ADR-003 — Stack técnico

- **Estado:** pendiente.
- **Propuesta a validar:** Python, PySide6, `python-docx`, parser OOXML, `pywin32` para COM, SQLite y pytest.
- **Criterios de decisión:** soporte Windows/Linux, robustez, licencias, facilidad de empaquetado, accesibilidad, rendimiento y mantenibilidad.
- **Siguiente acción:** prototipo vertical antes de cerrar la decisión.

## ADR-004 — Versión objetivo del examen

- **Estado:** aceptada como referencia inicial.
- **Fecha:** 2026-10-27.
- **Decisión:** priorizar MOS Word Associate (Office 2019), examen MO-100, la versión vigente identificada en Microsoft Learn.
- **Fuente oficial:** https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/
- **Evidencia consultada:** la página enumera seis dominios de evaluación y sus rangos porcentuales.
- **Condición:** antes de fijar subobjetivos y anunciar cobertura, revisar el documento oficial “Exam MO-100 skills outline” y volver a comprobar el estado de la certificación antes de publicar.

## ADR-005 — Distribución

- **Estado:** aceptada.
- **Fecha:** 2026-10-27.
- **Decisión:** entregar un instalador `.exe` para estudiantes y mantener código fuente multiplataforma para desarrollo.
- **Consecuencia:** el empaquetado Windows se prueba por separado del flujo de desarrollo en Linux.

## ADR-006 — MVP 100 % offline

- **Estado:** aceptada.
- **Fecha:** 2026-10-27.
- **Decisión:** las funciones esenciales no deben requerir conexión, cuentas, telemetría, APIs remotas ni backend.
- **Consecuencia:** catálogo, ejercicios, evaluación, ayuda e historial deben ser locales.
- **Verificación:** pruebas con red desactivada y revisión de llamadas de red en el flujo esencial.

## ADR-007 — Nombre del producto

- **Estado:** aceptada.
- **Fecha:** 2026-10-27.
- **Decisión:** el nombre de la aplicación será **VirMetrix**.
- **Distinción:** SAQI es la metodología de desarrollo utilizada para construir y mantener el producto; no es el nombre de la aplicación.

## ADR-008 — Stack tecnológico del MVP

- **Estado:** aceptada como stack inicial.
- **Fecha:** 2026-10-27.
- **Decisión:** Python, PySide6, `python-docx`, evaluador OOXML, `pywin32` para Word COM en Windows, SQLite, pytest y PyInstaller como candidato de empaquetado.
- **Motivo:** permite un núcleo multiplataforma, integración específica con Word en Windows y evaluación de documentos sin Word instalado.
- **Condición:** validar versiones, licencias, compatibilidad, seguridad y empaquetado con el vertical slice antes de una versión de producción.

## ADR-009 — Corrección del examen objetivo

- **Estado:** aceptada.
- **Fecha:** 2026-10-27.
- **Decisión:** VirMetrix se orienta a **Microsoft Office Specialist: Word Associate (Office 2019), examen MO-100**.
- **Corrección:** reemplaza la decisión anterior que apuntaba a MO-110 para Microsoft 365 Apps.
- **Fuente oficial:** https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/
- **Evidencia:** Microsoft Learn identifica MO-100 como Microsoft Word (Office 2019), enumera seis dominios de habilidades, indica una duración de 50 minutos y lista español entre los idiomas disponibles. No se muestra fecha de retirada en la página consultada.
- **Acción pendiente:** descargar y revisar el “MO-100 Exam Skills Outline” antes de cerrar la matriz de cobertura.
