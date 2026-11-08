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

## Observaciones en captura (@captura2_2026-10-29_10-25.png)

Basado en lo reportado:

**Positivo**
- Ventana: marca "VirMetrix" — OK
- Encabezado: degradado visual — buen contraste
- Chips: visualizan 1 · Copia → 2 · Editar → 3 · Evaluar — claros
- Selector de ejercicio: la lista cabe en 460 px — OK
- Instrucciones: aparecen en tarjeta con fondo blanco — OK
- Temporizador: visible (20:00) y formato claro — bueno
- Modo de evaluación: Portátil seleccionado (predeterminado) — OK
- Crear copia: el botón funciona; ya no hay apertura automática de Word — OK
- Abrir en Word: abre el documento y la UI no se congela — OK
- Evaluar: se activa, barra de progreso visible, UI no bloqueada — OK
- Score: se muestra como 1/2 con colores (ámbar/rojo) — buen feedback visual
- Resultados: los criterios aparecen con iconos y texto (C1, C2) — OK
- Mantener encima: el checkbox funciona — se confirma en captura
- Acerca de: el botón está presente — OK
- Scroll: el panel es desplazable — OK

**Posible ajuste**
- Chip visual: desde "1 · Copia" (paso 1) hasta "2 · Editar" (paso 2) hasta "3 · Evaluar" (paso 3).
  Dado que el workflow actual es: copiar → editar (en Word) → evaluar, el paso 2 debería ser "Editar" (y el chip debe estar resaltado cuando se edita el documento).

## Acción recomendada

- Limpiar el chip de paso cuando no haya edición en Word (es decir, después de copiar, antes de abrir el documento) o durante la edición, pero esto es preferentemente parte de la lógica del flujo (pasar al paso 2 al abrir el documento en Word).
- Confirmar el orden de foco: Modo → Ejercicio → Anterior/Siguiente → Modo de evaluación → Crear copia → Abrir → Evaluar → Cancelar → Iniciar/Pausar/Finalizar → Mantener encima → Acerca de → Resultados.
- Validar que todos los controles tengan nombres accesibles (ya agregamos `setAccessibleName` en los chips de paso).
- Probar con Narrador y llenar la checklist (`ACCESSIBILITY_CHECKLIST.md`).

## Próximos pasos

1. Abrir la captura para revisar visualmente cualquier problema de contraste o texto pequeño.
2. Probar en Windows VM con Narrador y llenar la checklist.
3. Si algún control está oculto para Narrador, agregar `setAccessibleName` / `setAccessibleDescription` donde falte.
4. Verificar nuevamente el orden de tabulación (`setTabOrder`).
5. Después de accesibilidad aprobada, preparar el empaquetado Windows (.exe).
