# Estrategia de pruebas

## Capas

1. **Unitarias:** normalización de criterios, reglas de puntuación, validación de metadatos, parsing de fragmentos DOCX.
2. **Integración:** catálogo + plantilla + evaluador + persistencia.
3. **Funcionales:** iniciar, completar, reiniciar y evaluar un ejercicio.
4. **E2E:** flujo completo en UI, primero con evaluador portable; luego con Word COM en Windows.
5. **Adversariales SAQI:** entradas hostiles, estados límite, fallos de proceso y resultados falsos positivos.
6. **Compatibilidad:** Linux sin Word, Windows sin Word y Windows con Word instalado.

## Matriz de entorno mínima

| Entorno | DOCX portable | Word COM |
|---|---:|---:|
| Linux sin Word | Obligatorio | No disponible |
| Windows sin Word | Obligatorio | No disponible |
| Windows con Word compatible | Obligatorio | Obligatorio |
| CI sin entorno gráfico | Núcleo y evaluador portable | Solo si existe runner Windows configurado |

## Casos mínimos

- DOCX válido con formato esperado.
- DOCX válido con formato incorrecto.
- Texto objetivo ausente.
- Texto dividido en varios runs.
- Estilos heredados desde párrafo o estilo de documento.
- DOCX corrupto o ZIP malformado.
- Archivo con tamaño excesivo.
- Ruta con espacios, Unicode y caracteres especiales.
- Archivo que no se puede leer por permisos.
- Word no instalado o COM no disponible.
- Word no responde o queda un proceso residual.
- Evaluación cancelada.
- Doble clic en finalizar.
- Criterio `not_verifiable` y su impacto en la puntuación.
- Plantilla original intacta después del intento.

## Reglas de evidencia

Cada ejecución debe registrar: ID de prueba, entorno, versión de aplicación, resultado esperado, resultado real, estado, evidencia y defecto relacionado.

No declarar “100 % aprobado” sin identificar el conjunto ejecutado. Separar pruebas no ejecutadas de las aprobadas. Guardar muestras de DOCX sintéticos y no incluir documentos personales reales.

## QA adversarial

- **Entrada:** corrupción, ZIP bombs mitigados por límites, XML inesperado, entidades externas y contenido activo.
- **Filesystem:** traversal de rutas, enlaces simbólicos, sobrescritura fuera del directorio de práctica.
- **Automatización:** COM bloqueado, documento modal, cierre inesperado de Word.
- **Estado:** evaluación concurrente, repetición, cierre durante guardado, resultados parciales.
- **Puntuación:** falsos positivos, redondeo, criterios duplicados y `not_verifiable` contado erróneamente como aprobado.
- **Privacidad:** logs con texto del documento, rutas personales o datos innecesarios.
- **Usabilidad:** pérdida de foco, lector de pantalla, navegación por teclado y mensajes de error ambiguos.

## Gate de release

- Las pruebas unitarias e integración requeridas pasan.
- El flujo E2E portable pasa en Linux y Windows.
- El flujo COM pasa en un entorno Windows de prueba con Word instalado.
- QA adversarial completado y defectos críticos/altos cerrados.
- Limitaciones y pruebas no ejecutadas documentadas.

## Pruebas específicas de arquitectura

- Importar el núcleo en Linux sin módulos Windows instalados.
- Importar la aplicación con Word ausente y comprobar que el modo portable sigue disponible.
- Confirmar que la UI no se congela mientras se evalúa un archivo.
- Confirmar que la automatización COM no termina instancias de Word preexistentes del usuario.
- Comprobar el paquete de distribución en una máquina Windows limpia.
- Ejecutar el flujo esencial con conectividad desactivada y verificar que no se intenta contactar servicios remotos.
