# Checklist de accesibilidad — VirMetrix (Windows + Narrador)

## Preparación

1. Actualizar la app: `git pull`.
2. Activar Narrador: `Win + Ctrl + Enter`.
3. Ejecutar la app (desde la raíz del repo):
   ```powershell
   $env:PYTHONPATH="src"
   python -m virmetrix.ui.main_window
   ```
4. Usar solo teclado: `Tab`, `Shift+Tab`, `Enter`, flechas.

## Orden de foco esperado (ITER-021)

1. Modo de uso (Práctica / Examen)
2. Selector de ejercicio
3. Anterior
4. Siguiente
5. Modo de evaluación (Portable / Word COM)
6. Crear copia de trabajo
7. Abrir en Word
8. Evaluar ahora
9. Cancelar evaluación (solo visible mientras evalúa)
10. Iniciar
11. Pausar / Reanudar
12. Finalizar
13. Mantener encima
14. Acerca de
15. Resultados de la evaluación

## Verificación por control (marcar OK / FAIL)

- [ ] La ventana anuncia "VirMetrix".
- [ ] Los chips "Paso 1, crear copia", "Paso 2, editar en Word" y "Paso 3, evaluar"
      se leen y anuncian su estado (paso actual / completado / pendiente).
- [ ] "Modo de uso" se anuncia con su nombre y la opción seleccionada.
- [ ] "Selector de ejercicio" anuncia el ejercicio actual.
- [ ] "Instrucciones del ejercicio" se leen al cambiar de ejercicio.
- [ ] "Temporizador" anuncia el tiempo restante.
- [ ] "Modo de evaluación" anuncia si es Portable o Word COM y su disponibilidad.
- [ ] "Crear copia de trabajo" tiene nombre y tooltip legible.
- [ ] "Abrir en Word" abre el archivo y no bloquea la UI.
- [ ] "Evaluar ahora" anuncia el inicio y el final de la evaluación.
- [ ] "Cancelar evaluación" solo aparece durante la evaluación y se anuncia.
- [ ] Puntuación ("Puntuación" + "Detalle de puntuación") se lee al terminar.
- [ ] Resultados por criterio se leen completos (C1..Cn, estado y evidencia).
- [ ] Mensajes de error (QMessageBox) se leen automáticamente.
- [ ] Botones deshabilitados se anuncian como no disponibles.

## Anotar defectos

- Controles sin nombre o mal anunciados:
- Orden de foco incorrecto:
- Textos que no se leen (temporizador, puntuación, errores):
- Atajos faltantes (si aplica):

## Notas

- Si el temporizador no se anuncia al cambiar, considerar `QAccessible.updateAccessibility`
  o mostrar el tiempo también en un control enfocable.
- Si la puntuación no se anuncia al terminar, enfocar el panel de resultados
  automáticamente tras la evaluación.
- Revisar `setTabOrder` si el orden visual no coincide con el de foco.
