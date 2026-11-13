"""Sirve el complemento web de Word (Office.js) en http://localhost:3000.

Uso (Windows/Linux):
    python scripts/serve_addin.py
Luego, en Word: Insertar > Mis complementos > Cargar complemento (manifest.xml).
"""
import http.server
import os
import socketserver

PORT = 3000
ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "addin", "VirmetrixWordWebAddIn")

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

if __name__ == "__main__":
    with socketserver.TCPServer(("localhost", PORT), Handler) as httpd:
        print(f"Sirviendo complemento en http://localhost:{PORT}/index.html")
        print(f"Raiz: {ROOT}")
        httpd.serve_forever()
