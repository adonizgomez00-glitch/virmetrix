# CONTEXT — VirMetrix

## Identidad

- Nombre de trabajo: VirMetrix.
- Propósito: simulador gratuito de práctica para objetivos de Microsoft Office Specialist (MOS) Word Associate.
- Metodología de desarrollo: SAQI (Skill-based Agentic Quality Iteration), rama experimental del repositorio del usuario.
- Idioma inicial de la interfaz y documentación: español.
- Plataforma principal de ejecución: Windows.
- Plataforma de desarrollo: Linux y Windows.

## Decisiones confirmadas por el usuario

1. Arquitectura A: aplicación de escritorio integrada con Word real.
2. También debe ser posible practicar cuando Microsoft Word no está instalado.
3. El desarrollo se realizará principalmente en Linux, pero la experiencia completa debe priorizar Windows.

## Interpretación técnica

La aplicación tendrá una interfaz de escritorio y dos rutas de evaluación explícitas:

- `word_com`: integración con Word de escritorio en Windows, disponible solo cuando Word esté instalado y la automatización COM sea accesible.
- `docx_portable`: evaluación del paquete `.docx` mediante inspección OOXML, sin Word instalado.

No se debe presentar el evaluador portable como idéntico al comportamiento de Word. Los criterios que no puedan inferirse de manera fiable del DOCX deben marcarse como no verificables o requerir validación adicional, no aprobarse por suposición.

## Reglas de trabajo SAQI

- Seguir el ciclo: Plan → Architect → Skill Selection → Build → Test → QA adversarial → Fix → Verify → Learn → mejorar Skills.
- Mantener los cuatro documentos vivos requeridos por SAQI y `SESSION.md` como checkpoint.
- Distinguir hechos observados, decisiones, hipótesis y trabajo pendiente.
- Cada corrección de un defecto debe añadir o actualizar una prueba de regresión cuando sea viable.
- No declarar cobertura, pruebas, integración o funcionalidades que no se hayan ejecutado.
- La calidad se evalúa sobre el vertical slice integrado y después sobre el producto; no bloquear el avance por perfeccionar módulos aislados si no impiden el MVP.
- La fase de QA adversarial no se omite antes de una versión candidata.
- Actualizar `PROJECT_STATE.md`, `QA_RESULTS.md` y `SESSION.md` al finalizar cada iteración significativa.

## Restricciones de contenido y marca

- Crear ejercicios, documentos de práctica, criterios de puntuación y explicaciones originales.
- Usar los objetivos públicos de certificación como mapa; verificar la versión vigente de los objetivos antes de fijar el catálogo definitivo.
- No extraer ni redistribuir preguntas, simulacros o recursos propietarios de terceros.
- Indicar claramente que el proyecto es independiente y no está afiliado, aprobado ni respaldado por Microsoft.
- No usar logotipos o recursos de marca sin permiso.

## Supuestos pendientes de confirmar

- Versión exacta de MOS Word Associate que se tomará como referencia.
- Windows mínimo y versiones de Word que se soportarán primero.
- Licencia del código y licencia de los ejercicios.
- Stack de interfaz y empaquetado.
- Si LibreOffice será una ayuda opcional para inspección visual, nunca un sustituto implícito de Word.

## Decisiones confirmadas adicionales

4. Usar como referencia la certificación vigente **Microsoft Office Specialist: Word Associate (Office 2019)**, examen **MO-100**, sujeto a una comprobación del esquema oficial de aptitudes antes de publicar el mapa completo.
5. Entregar un instalador `.exe` para Windows y conservar código fuente ejecutable/desarrollable en Linux y Windows.
6. El MVP debe ser 100 % offline: sin login, telemetría, servicios de nube ni APIs remotas necesarias para practicar, evaluar o consultar el historial.

## Referencia oficial

Página oficial del examen MO-100: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

La página consultada enumera estos dominios y pesos:
- Manage documents: 20–25%.
- Insert and format text, paragraphs, and sections: 20–25%.
- Manage tables and lists: 15–20%.
- Create and manage references: 5–10%.
- Insert and format graphic elements: 15–20%.
- Manage document collaboration: 5–10%.

Los pesos y objetivos deben revisarse periódicamente contra el esquema oficial de aptitudes descargable. La página del examen indicaba que no había fecha de retirada al momento de consulta (27-10-2026); volver a comprobarlo antes de cada publicación.

## Identidad y stack aprobados

- Nombre del producto: **VirMetrix**.
- SAQI se utiliza como metodología de desarrollo; no forma parte del nombre comercial de la aplicación.
- Stack inicial aprobado: Python, PySide6, `python-docx`, evaluador OOXML, `pywin32` en Windows, SQLite, pytest y PyInstaller como candidato de empaquetado.
- El stack se considera decisión del MVP, pero cada componente sigue sujeto a validación técnica, de licencias y de empaquetado durante el vertical slice.

## Corrección de certificación — 2026-10-27

El objetivo del proyecto es **Microsoft Office Specialist: Word Associate (Office 2019)**, examen **MO-100**, no MO-110. Fuente oficial: https://learn.microsoft.com/en-us/credentials/certifications/mos-word-2019/

La página oficial consultada lista los dominios:
- Administrar documentos.
- Insertar y dar formato a texto, párrafos y secciones.
- Administrar tablas y listas.
- Crear y administrar referencias.
- Insertar y dar formato a elementos gráficos.
- Administrar la colaboración en documentos.

La página indica una duración de 50 minutos y ofrece el examen en español, entre otros idiomas. No muestra fecha de retirada para MO-100 en la página consultada. Verificar la página oficial y el archivo descargable “MO-100 Exam Skills Outline” antes de publicar una versión del catálogo.

## Actualización ITER-001 (2026-10-27)

### Patrones de éxito
- Patrón: Evaluador OOXML directo — lectura de `word/document.xml` con límites de seguridad; permite runs fragmentados y estados honestos.
- Patrón: Estados de criterio `pass/fail/not_verifiable/error`; `not_verifiable` no suma.
- Patrón: Adaptador COM con detección explícita y sin importación global de `pywin32`.

### Supuestos validados
- Stack Python + PySide6 + python-docx + lxml + pytest funciona en Linux para el núcleo y el evaluador portable.
- `python-docx` suficiente para generar documentos de referencia; la evaluación se basa en OOXML.

### Fixes recientes
- Ninguno (0 defectos registrados en ITER-001).
