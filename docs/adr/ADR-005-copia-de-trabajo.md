# ADR-005: Trabajo sobre copia; plantilla original inmutable

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

El criterio C5 exige que el archivo de referencia permanezca idéntico; el usuario trabaja sobre una copia.

## Decisión

Al iniciar un intento, la aplicación crea/selecciona una copia de trabajo y evalúa únicamente esa copia. La plantilla de referencia se trata como inmutable y se verifica por hash (C5).

## Consecuencias

- C5 verificable por comparación de hash.
- Requiere gestión de rutas y permisos (adversarial: traversal, Unicode).

## Referencias

- `VERTICAL_SLICE_001.md` (AC-02, C5)
