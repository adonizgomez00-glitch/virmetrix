# Releases / Binarios

Coloca aquí el instalador/ejecutable (`.exe` para Windows, o el binario Linux) que quieras
publicar en el repositorio.

- `dist/` y `build/` siguen ignorados por Git (son salidas temporales de PyInstaller).
- Esta carpeta `releases/` NO está ignorada, así que los archivos que pongas aquí sí se
  sincronizan con GitHub.
- Para generar el ejecutable, consulta `README.md` (sección Empaquetado) o el workflow
  `.github/workflows/build-windows.yml`.

> Recomendación: para binarios grandes (>50 MB) usa **GitHub Releases** en lugar de
> commitearlos al repo. El workflow adjunta automáticamente el `.exe` a cada release
> creado con un tag `v*`.