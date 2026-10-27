# Validación en Windows (Word COM)

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