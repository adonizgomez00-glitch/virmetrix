# ADR-006: Seguridad del parsing OOXML

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

El evaluador portable procesa archivos proporcionados por el usuario, incluyendo entradas hostiles (ZIP bomb, XML inesperado, traversal).

## Decisión

Aplicar límites de tamaño y número de entradas, validación de rutas dentro del directorio de trabajo, deshabilitación de entidades externas/contenido activo, y manejo controlado de errores (`error` por criterio o error global controlado). No ejecutar macros ni contenido activo.

## Consecuencias

- Mitigación de riesgos de seguridad (SECURITY.md, QA adversarial).
- Casos de prueba específicos (AC-09).

## Referencias

- `SECURITY.md`, `TEST_STRATEGY.md` (adversarial)
- `VERTICAL_SLICE_001.md` (AC-09)
