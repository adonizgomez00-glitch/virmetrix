# ADR-003: Adaptador Word COM solo en Windows con importación bajo demanda

**Fecha:** 2026-10-27  
**Estado:** Accepted  
**Iteración:** ITER-001

## Contexto

`pywin32` no está disponible en Linux y no debe ser dependencia del núcleo. El núcleo debe importarse en Linux sin módulos Windows.

## Decisión

El adaptador COM realiza importaciones de `win32com`/`pythoncom` bajo demanda y solo en Windows. Expone `is_available()` que devuelve `False` con motivo si Word o COM no están accesibles. El núcleo no importa estos módulos al importarse.

## Consecuencias

- Núcleo portable y testeable en Linux.
- Detección explícita de Word (AC-04).
- Requiere timeouts, limpieza de documentos temporales y no terminar procesos `WINWORD.EXE` globales.

## Alternativas

- Importación condicional en módulo compartido: descartado por acoplamiento.

## Referencias

- `ARCHITECTURE.md` (Contratos técnicos)
- `VERTICAL_SLICE_001.md` (AC-04, AC-05, AC-09)
