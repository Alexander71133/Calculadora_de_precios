@echo off
:: Cambia al directorio donde está el script
cd /d "%~dp0"
:: Ejecuta usando el python del entorno virtual (.venv)
start "" ".\.venv\Scripts\pythonw.exe" "calculadora_rapida.pyw"
exit