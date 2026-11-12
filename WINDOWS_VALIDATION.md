# Validación en Windows (Word COM)

## Checklist v0.4.2 (ejercicios tocados / con problemas recientes)

Probar en Word COM los siguientes ejercicios y registrar score de referencia y de documento correcto:

1. `d-text-para-001` — texto y párrafos
2. `d-text-para-002` — texto y párrafos
3. `d-table-001` — tablas
4. `d-table-002` — tablas
5. `d-table-003` — ordenar por “Precio” (debe pasar solo si se ordena por Precio)
6. `d-table-004` — márgenes de celda
7. `d-table-005` — ajuste/redimensionado de tabla
8. `d-list-001` — 3 niveles de lista (apertura COM)
9. `d-list-002` — reiniciar numeración
10. `d-list-003` — iniciar lista en 5
11. `d-refs-001` — TOC personalizada niveles 1-2
12. `d-graphics-001..005` — gráficos/alt text
13. `d-collab-001` — resolver comentario (Office 2024 moderno)
14. `d-collab-002` — revisar cambios
15. `d-doc-001` — propiedades del documento
16. `d-doc-002` — guardar/exportar PDF
17. `c-text-para-001` — estilo “Heading 2” / “Encabezado 2”
18. `text-para-004` — estilo “Heading 1” / “Encabezado 1”

Evidencia a registrar: score referencia, score documento correcto, captura y observaciones en `QA_RESULTS.md`.



## 1. Obtener el código

Opción A — Git (recomendado):
```bash
git init
git add .
git commit -m "VirMetrix ITER-002"
git remote add origin <URL>
git push -u origin main
```
En Windows:
```powershell
git clone <URL>
cd virmetrix
```

Opción B — Archivo comprimido:
- En Linux: `zip -r virmetrix.zip . -x "*.pyc" -x "__pycache__/*" -x ".pytest_cache/*"`
- Copiar `virmetrix.zip` a Windows y descomprimir.

Opción C — Carpeta compartida/USB: copiar todo el directorio del proyecto.

## 2. Preparar entorno en Windows

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
```

Requisitos:
- Windows 10/11.
- Microsoft Word de escritorio instalado (licencia propia).
- Python 3.11 o 3.12.

## 3. Ejecutar pruebas portables

```powershell
$env:PYTHONPATH="src"
pytest -q
```
Esperado: 33+ passed, 0 failed (las pruebas COM se saltan/no ejecutan en Linux, pero en Windows `test_word_com` valida disponibilidad).

## 4. Validar Word COM

```powershell
$env:PYTHONPATH="src"
python -c "from virmetrix.infrastructure.evaluators.word_com import WordComAdapter; print(WordComAdapter().is_available())"
```
Debe imprimir `(True, 'Word disponible (v...).')`.

Prueba de evaluación COM (manual):
```powershell
$env:PYTHONPATH="src"
python - <<'PY'
from pathlib import Path
from virmetrix.application.orchestrator import EvaluationOrchestrator
from virmetrix.infrastructure.catalog.loader import load_exercise
o = EvaluationOrchestrator()
ref = Path("exercises/word-format-001/reference.docx")
work = o.create_work_copy(str(ref), str(Path.home()/".virmetrix"/"work"))
ex = load_exercise("word-format-001")
res = o.evaluate("word_com", work, ex)
print(res.evaluator, res.score_earned, "/", res.score_possible)
for c in res.criteria:
    print(c.criterion_id, c.status.value, "-", c.evidence)
PY
```

## 5. Validar UI

```powershell
$env:PYTHONPATH="src"
python -m virmetrix.ui.main_window
```
- Crear copia de trabajo.
- Evaluar en modo `docx_portable` y `word_com`.
- Confirmar que no se bloquea la UI y que el historial se guarda.

## 6. Evidencia a registrar en QA_RESULTS.md

- SO y versión de Windows.
- Versión de Word.
- Versión de Python y dependencias.
- Comando ejecutado y resultado.
- Estados C1–C5 para documento correcto/incorrecto.
- Defectos encontrados (si los hay).

No declarar COM como aprobado sin esta evidencia.