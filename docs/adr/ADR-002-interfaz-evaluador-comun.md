# ADR-002: Interfaz común de adaptador de evaluación

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

El orquestador de evaluación debe soportar dos rutas (`docx_portable` y `word_com`) sin cambiar silenciosamente de una a otra durante un intento, y devolver resultados comparables.

## Decisión

Definir el puerto `EvaluationAdapter` con el método `evaluate(file_path, exercise_metadata) -> EvaluationResult`. Tanto el evaluador portable como el adaptador Word COM implementan este puerto. El orquestador selecciona el adaptador disponible y expone capacidades explícitas (`available`, `reason`).

## Consecuencias

- Contrato único para resultados y puntuación.
- Los adaptadores pueden diferir en verificabilidad (`not_verifiable`) sin romper el contrato.
- El modo elegido se registra en `evaluator` del resultado.

## Alternativas

- Adaptadores con APIs distintas: descartado por duplicación de lógica de orquestación.

## Referencias

- `ARCHITECTURE.md` (Contrato de resultado, Orquestador)
