@echo off
rem Doble clic para abrir la interfaz grafica de cruce CONSULTA OSA vs Edimusica.
rem Usa el Python instalado en este equipo (no requiere terminal ni comandos).

set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe

if not exist "%PYTHON_EXE%" (
    echo No se encontro Python en "%PYTHON_EXE%".
    echo Instala Python 3 o ajusta esta ruta dentro del archivo .bat.
    pause
    exit /b 1
)

start "" "%PYTHON_EXE%" "%~dp0recon_osa_gui.py"
