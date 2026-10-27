# ITERATION_REPORT — ITER-001 (Vertical Slice word-format-001)

**Fecha:** 2026-10-27  
**Estado:** Candidato a release (pendiente sign-off humano)

## 1. Resumen

Se implementó el vertical slice `word-format-001`: evaluador portable de DOCX (OOXML) con 5 criterios (C1–C5), adaptador Word COM base (detección), UI mínima y pruebas automatizadas. 27+ pruebas pasan en Linux; Word COM no ejecutado (entorno Linux) y registrado como no ejecutado.

## 2. Alcance cumplido

- AC-01 Inicio offline: sí (sin red).
- AC-02 Copia de trabajo: sí.
- AC-03 Modo portable: sí (estados C1–C5, puntuación, evidencia).
- AC-04 Detección explícita de Word: sí (Linux → no disponible).
- AC-05 Modo Word real: no ejecutado (sin Windows/Word). Registrado como no ejecutado.
- AC-06 Corrección positiva: sí (documento correcto → 5/5).
- AC-07 Corrección negativa: sí (negrita/tamaño ausentes → fail).
- AC-08 Texto fragmentado: sí (pass).
- AC-09 Archivo inválido: sí (error controlado).
- AC-10 Evidencia de resultados: sí (ID, versión, adaptador, fecha, estados, versión de reglas).
- AC-11 Reproducibilidad: sí.
- AC-12 Regresión automatizada: sí (27+ pass; COM no ejecutado).

## 3. Métricas

- Tests totales: 27+ pass, 0 fail.
- Cobertura (excl. UI y rama COM de éxito): 92.36% (umbral 80%).
- Defectos abiertos: 0.
- `not_verifiable` tratado como no puntuable.

## 4. Defectos

Ninguno registrado en esta iteración.

## 5. Lecciones (para Fase 14)

- El parsing OOXML directo permite verificar runs fragmentados y estados honestos.
- La seguridad de ZIP/XML debe ser parte del diseño inicial (límites, traversal, entidades).
- `not_verifiable` es esencial para no falsear puntuación ante herencia de estilos.
- La UI requiere threading para no bloquear; diferido a siguiente iteración.

## 6. Trabajo diferido

- Persistencia SQLite del historial (ITER-002).
- Evaluación COM completa en Windows.
- Empaquetado PyInstaller / instalador.
- Temporizador de examen, catálogo amplio.
