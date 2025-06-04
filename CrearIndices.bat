@echo off
:: Establece la ruta de trabajo a la ubicación del archivo .bat
cd /d %~dp0

:: Ejecuta el comando con el archivo .exe
"Creador_de_indices.exe" CREAR_ARCHIVO

REM Cerrar la ventana de este BAT
exit