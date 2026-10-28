# Checklist de accesibilidad — VirMetrix (Windows + Narrador)

## Preparación
1. Actualizar la app: `git pull` (hecho).
2. Activar Narrador: `Win + Ctrl + Enter`.
3. Ejecutar la app: `python main.py` (o el `.exe`).
4. Usar solo teclado: `Tab`, `Shift+Tab`, `Enter`, flechas.

## Orden de foco esperado
- Modo de uso (Práctica/Examen)
- Selector de ejercicio
- Instrucciones
- Temporizador
- Modo de evaluación (portable/Word COM)
- Crear copia de trabajo
- Ejercicio anterior
- Ejercicio siguiente
- Iniciar sesión
- Pausar o reanudar
- Evaluar
- Finalizar sesión
- Acerca de VirMetrix
- Resultados de la evaluación

## Verificación por control (marcar OK / FAIL)
- [ ] "Modo de uso" se anuncia con su nombre.
- [ ] "Selector de ejercicio" anuncia el ejercicio actual.
- [ ] "Instrucciones del ejercicio" se leen al cambiar de ejercicio.
- [ ] "Temporizador" anuncia el tiempo restante al cambiar.
- [ ] "Modo de evaluación" anuncia disponible/no disponible.
- [ ] "Crear copia de trabajo" tiene nombre y tooltip legible.
- [ ] "Abrir documento" abre el archivo y no bloquea la UI.
- [ ] "Evaluar" anuncia resultado/puntuación al terminar.
- [ ] "Resultados de la evaluación" se leen (incluye C1..Cn y puntuación).
- [ ] Mensajes de error (QMessageBox) se leen automáticamente.
- [ ] Botones deshabilitados se anuncian como no disponibles.

## Anotar defectos
- Controles sin nombre o mal anunciados:
- Orden de foco incorrecto:
- Textos que no se leen (temporizador, puntuación, errores):
- Atajos faltantes (si aplica):

## Notas
- Si el temporizador no se anuncia al cambiar, considerar usar `QAccessible.updateAccessibility` o un `QLabel` con `aria-live` equivalente.
- Revisar `setTabOrder` si el orden visual no coincide con el de foco.
