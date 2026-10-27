# Instrucciones iniciales para OpenCode — VirMetrix

Actúa como ingeniero de software responsable del proyecto **VirMetrix**. Lee primero, en este orden:

1. `CONTEXT.md`
2. `PROJECT_STATE.md`
3. `REQUIREMENTS.md`
4. `ARCHITECTURE.md`
5. `VERTICAL_SLICE_001.md`
6. `TEST_STRATEGY.md`
7. `SECURITY.md`
8. `DECISIONS.md`
9. `SESSION.md`

Si el repositorio SAQI contiene reglas de agente o Skills aplicables, inspecciónalas y síguelas. No inventes nombres de Skills ni afirmes haber ejecutado fases que no hayas completado.

## Objetivo de la primera tarea

Implementar únicamente la **Iteración 1 — Núcleo y evaluador portable** de `IMPLEMENTATION_PLAN.md`. No construyas todavía toda la interfaz, el adaptador COM, la base de datos ni el instalador.

## Stack

- Python
- PySide6 (para iteración posterior)
- python-docx cuando resulte apropiado
- parser OOXML explícito para criterios que lo requieran
- pywin32 solo en Windows y solo dentro del adaptador futuro
- SQLite para persistencia futura
- pytest
- PyInstaller como candidato de empaquetado futuro

## Reglas obligatorias

- Antes de programar, resume requisitos, plan, riesgos y pruebas.
- Inspecciona el repositorio existente y conserva convenciones útiles; no sobrescribas archivos sin entenderlos.
- Trabaja incrementalmente y ejecuta pruebas después de cada cambio relevante.
- No importes `win32com` desde el núcleo ni hagas que Linux requiera dependencias Windows.
- No marques `not_verifiable` como aprobado.
- Trata DOCX como entrada no confiable; impón límites ZIP/XML y no ejecutes macros ni contenido activo.
- No hagas llamadas de red; el MVP es offline.
- No afirmes que una prueba pasó si no la ejecutaste. Distingue `passed`, `failed` y `not run`.
- No cierres procesos Word globalmente ni modifiques documentos originales.
- No amplíes el alcance sin registrar una decisión en `DECISIONS.md`.
- Añade pruebas de regresión cuando corrijas un defecto.

## Entregables de la primera tarea

1. Estructura mínima del paquete Python.
2. Modelos de resultado y estados de criterio (`pass`, `fail`, `not_verifiable`, `error`).
3. Evaluador portable del ejercicio `word-format-001`.
4. Pruebas pytest con documentos sintéticos: correcto, negrita ausente, tamaño incorrecto, objetivo ausente, frase dividida en runs y DOCX inválido.
5. `README` técnico con comandos exactos para instalar dependencias y ejecutar pruebas en Linux.
6. Actualización de `PROJECT_STATE.md`, `QA_RESULTS.md` y `SESSION.md`.

## Criterio de finalización

La tarea termina solo cuando el evaluador puede ejecutarse localmente sin Word y las pruebas existentes se ejecutan realmente. Informa comandos y resultados exactos. Si no puedes probar Windows desde este entorno, documenta esa limitación y no simules resultados.

Primero presenta el plan y las dudas bloqueantes. Después implementa únicamente esta iteración.

## Versión de certificación obligatoria

El objetivo del producto es **Microsoft Office Specialist: Word Associate (Office 2019), examen MO-100**. No utilizar MO-110 ni objetivos específicos de Microsoft 365 Apps como referencia del proyecto. Consultar `WORD_COMPETENCIES.md` y la fuente oficial https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/ antes de crear o modificar ejercicios alineados al examen.
