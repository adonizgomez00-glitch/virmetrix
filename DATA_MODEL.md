# Modelo de datos inicial — VirMetrix

**Estado:** propuesta mínima para la primera implementación.

## Exercise

- `exercise_id`: identificador estable.
- `version`: versión del contenido.
- `title`: título mostrado.
- `language`: idioma.
- `competency_id`: objetivo o dominio de competencia.
- `instructions`: instrucciones para el estudiante.
- `starting_document`: ruta relativa controlada a la plantilla.
- `criteria`: lista ordenada de criterios.
- `content_version`: versión del mapa de competencias al que corresponde.

## Criterion

- `criterion_id`: identificador estable dentro del ejercicio.
- `description`: requisito verificable.
- `points`: peso entero no negativo.
- `evaluator_support`: adaptadores capaces de evaluar el criterio.
- `rule`: referencia a una regla de evaluación versionada.
- `not_verifiable_policy`: comportamiento explícito cuando no se puede verificar.

## CriterionResult

- `criterion_id`
- `status`: `pass | fail | not_verifiable | error`
- `points_earned`
- `points_possible`
- `evidence`: texto breve y no sensible.
- `reason_code`: código estable para pruebas y traducción.

## EvaluationResult

- `exercise_id`
- `exercise_version`
- `evaluator`: `docx_portable | word_com`
- `evaluator_version`
- `started_at` y `finished_at` en ISO 8601 local con zona horaria cuando esté disponible.
- `criteria`: lista de resultados.
- `score_earned` y `score_possible`
- `status`: `completed | cancelled | error`.

## Reglas de puntuación

- Cada ejercicio define su política y versión.
- Solo `pass` obtiene puntos.
- `fail` obtiene cero puntos para ese criterio.
- `not_verifiable` no obtiene puntos y se presenta de forma explícita; el informe advierte si el total es parcial.
- `error` invalida o deja incompleto el resultado según política del ejercicio; no se debe convertir en fallo silencioso.
- No llamar a la puntuación “puntuación oficial de Microsoft”.

## Persistencia

La primera iteración puede usar dataclasses/Pydantic (decidir tras revisar dependencias) y JSON de fixtures para pruebas. SQLite se añade cuando se implemente el historial; no hace falta introducirlo en la primera prueba del evaluador.
