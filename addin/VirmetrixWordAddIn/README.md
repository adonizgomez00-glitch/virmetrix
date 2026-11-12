# Virmetrix Word Add-in (VSTO)

Requiere Windows, Visual Studio con carga de trabajo "Office/SharePoint development" y Microsoft Word.

## Build
1. Abrir el proyecto en Visual Studio.
2. Compilar en modo Release.
3. Instalar el .vsto/.msi resultante en Windows.

## Funcionamiento
Los botones de la cinta registran una propiedad personalizada del documento (`Virmetrix_<accion>`). El evaluador Word COM/Python lee esa propiedad y marca el criterio como pass.
EOF
echo done