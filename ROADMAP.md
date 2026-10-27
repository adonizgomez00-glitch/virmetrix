# Roadmap

## Fase 0 — Definición y gobernanza

- [x] Registrar decisiones iniciales.
- [x] Crear documentación viva inicial.
- [ ] Confirmar versión objetivo del examen y alcance del MVP.
- [ ] Definir licencia de código y licencia de contenido.
- [ ] Seleccionar Skills SAQI aplicables y comprobar sus versiones.

**Salida:** especificación acordada y riesgos priorizados.

## Fase 1 — Prueba de viabilidad técnica

- [ ] Crear un DOCX piloto original.
- [ ] Definir criterios deterministas para formato de texto.
- [ ] Implementar lectura OOXML mínima en Linux.
- [ ] Implementar un adaptador COM mínimo en Windows.
- [ ] Comparar resultados entre evaluadores y documentar diferencias.
- [ ] Confirmar que el modo portable nunca afirma verificar criterios que no puede comprobar.

**Gate:** la misma tarea de prueba tiene resultados reproducibles y el modo de evaluación se identifica correctamente.

## Fase 2 — Vertical slice de escritorio

- [ ] Ventana principal y selección de ejercicio.
- [ ] Instrucciones y apertura de plantilla/copia de trabajo.
- [ ] Ejecución del evaluador.
- [ ] Vista de puntuación por criterio.
- [ ] Persistencia local mínima.
- [ ] Pruebas unitarias, integración y flujo E2E básico.

**Gate:** un ejercicio se puede iniciar, completar, evaluar y revisar en Linux (modo portable) y Windows (modo portable y, si está disponible, Word COM).

## Fase 3 — MVP de práctica

- [ ] Catálogo pequeño de ejercicios originales por varias competencias.
- [ ] Pistas y explicaciones.
- [ ] Historial de intentos.
- [ ] Modo práctica y modo cronometrado.
- [ ] Manejo de errores y recuperación.
- [ ] Paquete de instalación Windows.

## Fase 4 — QA SAQI

- [ ] QA funcional integrado.
- [ ] QA adversarial: documentos corruptos, ZIP malformados, rutas maliciosas, archivos grandes, estados incompletos, resultados inconsistentes y doble envío.
- [ ] Pruebas de accesibilidad y usabilidad.
- [ ] Regresión completa.
- [ ] Revisión de licencias, privacidad y marca.
- [ ] Actualizar `QA_RESULTS.md` con evidencia reproducible.

**Gate:** ningún defecto crítico o alto abierto; limitaciones conocidas documentadas; los tests no ejecutados no se cuentan como aprobados.

## Fase 5 — Publicación inicial

- [ ] Licencias y avisos legales.
- [ ] Instrucciones de instalación y desinstalación.
- [ ] Política de soporte de Windows/Word.
- [ ] Paquete de ejercicios con derechos claros.
- [ ] Versión y notas de publicación.

## Fase 0 — Definición y gobernanza

- [x] Registrar decisiones iniciales.
- [x] Confirmar referencia: MOS Word Associate (Office 2019), MO-100.
- [x] Confirmar distribución: instalador `.exe` y código fuente multiplataforma.
- [x] Confirmar MVP 100 % offline.
- [ ] Descargar/revisar el esquema oficial de aptitudes MO-100 y registrar versión/fecha.
- [ ] Definir licencia de código y licencia de contenido.
- [ ] Seleccionar y justificar stack.

**Salida:** especificación acordada y objetivos de examen trazables a fuente oficial.
