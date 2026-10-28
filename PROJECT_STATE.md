# PROJECT_STATE

**Última actualización:** 2026-10-27 (ITER-014 Build)  
**Estado global:** ITER-001..013 cerrados; ITER-014 en progreso (abrir documento en Word). 41 ejercicios.

## Resumen

Se propone una aplicación de escritorio gratuita para practicar tareas de MOS Word Associate, con Windows como plataforma principal y Linux como entorno principal de desarrollo.

## Iteración actual

- **Iteración:** ITER-001 — Vertical slice `word-format-001`.
- **Fase SAQI actual:** 14 — Improve Skills (ITER-001 completado: propuestas en SKILL_CHANGELOG.md; pendiente aprobación gobernanza).
- **Modo:** estándar con MVP-First en Build y macro-fase QA final.
- **Alcance:** AC-01..AC-12 de `VERTICAL_SLICE_001.md`.
- **Siguiente acción:** aprobación del plan por PO/Tech Lead → Fase 2 (Architect).

## Decisiones confirmadas

- Arquitectura de escritorio integrada con Word real cuando esté disponible.
- Debe permitir practicar también sin Word instalado.
- La evaluación portable se basará en el archivo DOCX y tendrá límites documentados.

## Hechos verificados

- El repositorio SAQI describe un proceso de 14 fases, Skills versionadas, QA dual (convencional y adversarial) y documentos de contexto persistente.
- En esta fase no se ha creado ni probado código de producto.

## Hipótesis de diseño

- Una aplicación Python con UI de escritorio puede reducir fricción para integrar COM en Windows y permitir desarrollo en Linux.
- `python-docx` y un evaluador OOXML propio pueden cubrir una primera parte de los criterios portables.
- La automatización COM debe aislarse en un adaptador para que el núcleo funcione en Linux.

Estas hipótesis requieren un prototipo antes de convertirse en decisiones definitivas.

## Trabajo actual

- [x] Recoger decisiones iniciales del usuario.
- [x] Preparar documentos iniciales de requisitos y gobernanza.
- [x] Confirmar versión objetivo de MOS Word Associate (MO-100, Office 2019).
- [ ] Seleccionar y justificar stack.
- [ ] Definir un ejercicio vertical de prueba y criterios verificables.
- [ ] Crear prototipo del evaluador DOCX portable.
- [ ] Crear adaptador COM mínimo en Windows.
- [ ] Ejecutar QA funcional y adversarial.

## Bloqueos / riesgos

1. El evaluador portable no puede garantizar fidelidad visual o funcional a Word para todos los criterios.
2. Automatizar Word depende de una instalación compatible de Word en Windows.
3. La versión de los objetivos de examen puede cambiar; el mapa debe versionarse.
4. Los ejercicios deben ser originales y respetar las licencias de contenido y marca.

## Siguiente paso

Responder las preguntas pendientes de `SESSION.md` y después realizar un prototipo que verifique una tarea de formato de texto tanto en modo portable como en Word real.

## Decisiones adicionales confirmadas

- [x] Referencia de examen: MOS Word Associate para Office 2019, MO-100, según página oficial consultada el 2026-10-27.
- [x] Distribución: instalador Windows `.exe` y código fuente multiplataforma.
- [x] MVP 100 % offline.

## Verificación oficial del examen

La página oficial Microsoft Learn del examen MO-100 lista seis dominios con estos pesos:
- Administrar documentos: 20–25 %.
- Insertar y dar formato a texto, párrafos y secciones: 20–25 %.
- Administrar tablas y listas: 15–20 %.
- Crear y administrar referencias: 5–10 %.
- Insertar y dar formato a elementos gráficos: 15–20 %.
- Administrar la colaboración en documentos: 5–10 %.

Fuente: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

La versión objetivo está identificada, pero los subobjetivos detallados aún deben reconciliarse con el documento oficial “Exam MO-100 skills outline”. La compatibilidad de idioma del examen puede depender de la oferta vigente; no asumir que todas las traducciones están disponibles.

## Decisiones nuevas

- [x] Nombre del producto: VirMetrix.
- [x] Stack inicial: Python + PySide6 + python-docx + evaluador OOXML + pywin32 (Windows) + SQLite + pytest; PyInstaller como candidato de empaquetado.
- [ ] Validar el stack mediante el vertical slice y comprobar licencias y compatibilidad de dependencias.

## Corrección confirmada del examen

- Objetivo correcto: **Microsoft Office Specialist: Word Associate (Office 2019)**.
- Número de examen: **MO-100**.
- Se descarta MO-110 como referencia de este proyecto.
- Fuente oficial: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/
- La página oficial lista seis dominios de habilidades, indica 50 minutos y ofrece español como idioma de examen; no muestra fecha de retirada en la página consultada el 2026-10-27.
- Aún falta revisar el PDF/archivo descargable “MO-100 Exam Skills Outline” para completar la trazabilidad de subobjetivos y pesos antes de afirmar cobertura integral.


## Verificación del esquema de aptitudes MO-100 completada

Se abrió el PDF oficial “MO-100: Microsoft Word (Office 2019) – Skills Measured”. Se confirmaron seis dominios y sus ponderaciones. Tablas y listas pesa 15–20 % (no 20–25 %). La lista de subobjetivos está registrada en `WORD_COMPETENCIES.md`. Fuente: https://arch-center.azureedge.net/Learning/Credentials/mo-100-microsoft-word-2019-skills-measured.pdf
