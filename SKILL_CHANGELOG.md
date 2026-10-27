# SKILL_CHANGELOG — Propuestas de mejora (ITER-001)

**Fase SAQI:** 14 — Improve Skills  
**Estado:** Propuestas para aprobación de gobernanza (no aplicadas automáticamente).

## Lecciones de ITER-001

1. El parsing OOXML directo requiere reglas de seguridad explícitas (límites de tamaño/entradas, traversal, entidades).
2. `not_verifiable` es un estado necesario para no falsear puntuación ante herencia de estilos.
3. La detección explícita de capacidades (Word COM) evita selección silenciosa y bloqueos.
4. La reproducibilidad exige hashing de la plantilla de referencia (C5).

## Propuestas

### P1 — saqi-qa-harness v1.0.0 → v1.1.0 (MINOR)
- Añadir regla SHOULD: los evaluadores de archivos de usuario deben validar tamaño, número de entradas y rutas (anti-traversal) antes de parsear.
- Añadir anti-patrón: contar `not_verifiable` como `pass`.
- Justificación: DEF (ninguno) / lección ITER-001 (seguridad OOXML, estados honestos).

### P2 — saqi-lean-code v1.0.0 → v1.0.1 (PATCH)
- Clarificar que el núcleo no debe importar módulos específicos de plataforma (ej. `pywin32`) al importarse.
- Justificación: ADR-003.

### P3 — saqi-vertical-slice v1.0.0 → v1.0.1 (PATCH)
- Añadir plantilla de criterio `not_verifiable` con motivo obligatorio.
- Justificación: ADR-004.

## Aprobación
- [ ] Aprobado por gobernanza: ______  Fecha: ______
