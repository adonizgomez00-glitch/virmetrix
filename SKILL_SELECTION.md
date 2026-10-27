# SKILL_SELECTION — VirMetrix (ITER-001)

**Fase SAQI:** 3 — Skill Selection  
**Fecha:** 2026-10-27  
**Estado:** Borrador para aprobación del humano (Tech Lead)  
**Fuente de skills:** OpenCode global (`~/.config/opencode/skills/`), instaladas desde SAQI (rama experimental).

## 1. Matriz Requerimiento × Skill

| Requerimiento / Historia | Skills aplicables |
|---|---|
| H1 Metadatos ejercicio + generador DOCX | `saqi-spec-arch`, `saqi-vertical-slice` |
| H2 Evaluador portable OOXML (seguridad, parsing) | `saqi-qa-harness`, `saqi-lean-code`, `saqi-testing` |
| H3 Modelo de resultado + scoring | `saqi-spec-arch`, `saqi-lean-code` |
| H4 UI mínima (PySide6) | `saqi-vertical-slice`, `saqi-lean-code` |
| H5 Adaptador Word COM | `saqi-qa-harness`, `saqi-testing` |
| H6 Pruebas automatizadas + adversariales | `saqi-testing`, `saqi-qa-cycles`, `saqi-qa-harness` |
| H7 Documentación y métricas | `saqi-context-manager`, `saqi-metrics`, `saqi-token-optimizer` |
| Gestión de contexto / checkpoints | `saqi-context-manager`, `saqi-token-optimizer` |
| Modo adaptativo (estándar/riguroso) | `saqi-adaptive-mode` |
| Harness de validación (comandos, exit criteria) | `saqi-harness` |
| Aprendizaje → Skills (Fase 14) | `saqi-skills-learning` |

## 2. Skills seleccionadas por nivel

### Nivel A — Obligatorias (Core) — 100% incluidas

| Skill ID | Versión | Fuente | Notas |
|---|---|---|---|
| `saqi-lean-code` | 1.0.0 | saqi | Consolida A-coding-standards, A-project-architecture |
| `saqi-spec-arch` | 1.0.0 | saqi | Consolida A-sdd, A-project-architecture, effect |
| `saqi-qa-harness` | 1.0.0 | saqi | Consolida C-qa-breaker, A-testing, A-secure-coding |
| `saqi-context-manager` | 1.0.0 | saqi | Contexto persistente sin burocracia |
| `saqi-testing` | 1.0.0 | saqi | Unit/e2e/component/tui |
| `saqi-harness` | 1.0.0 | saqi | Comandos de validación, exit criteria |
| `saqi-metrics` | 1.0.0 | saqi | Métricas de iteración |
| `saqi-token-optimizer` | 1.0.0 | saqi | Presupuestos de tokens |
| `saqi-adaptive-mode` | 1.0.0 | saqi | Modo estándar/riguroso |
| `saqi-vertical-slice` | 1.0.0 | saqi | Historias pequeñas y ACs |
| `saqi-qa-cycles` | 1.0.0 | saqi | Ciclo rápido vs profundo |
| `saqi-skills-learning` | 1.0.0 | saqi | Aprendizaje → Skills (Fase 14) |

### Nivel B — Dominio / Tecnología

| Skill ID | Versión | Aplicabilidad | Notas |
|---|---|---|---|
| `effect` | (opencode) | No aplica directamente | Solo si se usa Effect TS; aquí stack Python. Inactivo para código de producto. |
| `rtl-aware-development` | (opencode) | Carga bajo demanda | Solo si se detecta necesidad RTL. |
| `app-e2e-tests`, `component-tests`, `tui-tests` | (opencode) | No aplica (stack Python/PySide6) | Inactivos para esta iteración. |

### Nivel C — Verificación / Debugging

Incluidas vía `saqi-qa-harness` y `saqi-testing`. `saqi-qa-cycles` define triggers de QA profundo.

### Nivel D — Soporte

`saqi-skills-learning` (Fase 14), `saqi-metrics`, `saqi-token-optimizer`.

## 3. Skills globales referenciadas (no instaladas)

El registry de SAQI referencia skills globales `A-sdd`, `A-testing`, `A-secure-coding`, `D-git-workflow` bajo `~/.config/opencode/skills/generic|security|git/`. **No están instaladas actualmente** (solo se instalaron las skills de SAQI). Las saqi-* consolidadas cubren su contenido para esta iteración. Si se requieren las versiones originales, instalarlas por separado.

## 4. Verificación de compatibilidad

- No se detectaron conflictos entre las saqi-* seleccionadas (consolidadas y coherentes por diseño).
- `effect`/`tui-tests`/`component-tests` no aplican al stack Python y se mantienen inactivos para evitar consumo de tokens.
- `rtl-aware-development` se carga solo bajo demanda.

## 5. Configuración del agente

- Skills activas (Nivel A): todas las 12 saqi-* listadas.
- Skills inactivas para esta iteración: `effect`, `app-e2e-tests`, `component-tests`, `tui-tests`, `opencode-dev`, `sample-skill`, `rtl-aware-development` (bajo demanda).
- Versiones registradas: 1.0.0 para todas las saqi-*.

## 6. Aprobación

- [ ] Aprobado por humano (Tech Lead): ______ Fecha: ______
- [ ] Conflictos resueltos: sí
- [ ] Agente configurado: sí
