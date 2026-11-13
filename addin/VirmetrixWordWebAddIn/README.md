# Virmetrix Word Web Add-in (Office.js)

Para activarlo en Windows:
1. Ejecuta `python scripts/serve_addin.py` (o `python -m http.server 3000 --directory addin/VirmetrixWordWebAddIn`).
2. Abre Word y ve a Insertar > Mis complementos > Cargar complemento.
3. Selecciona `addin/VirmetrixWordWebAddIn/manifest.xml`.
4. El panel "Virmetrix" aparece; usa los botones para registrar acciones.

El manifest apunta a http://localhost:3000/index.html.
