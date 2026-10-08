@echo off
cd /d "%~dp0"
set "RETIE_PYTHON=%LocalAppData%\Programs\Python\Python314\python.exe"
if not exist "%RETIE_PYTHON%" set "RETIE_PYTHON=%LocalAppData%\Programs\Python\Python313\python.exe"
if not exist "%RETIE_PYTHON%" set "RETIE_PYTHON=python"
"%RETIE_PYTHON%" -c "import tkinter, openpyxl, PIL, googleapiclient, google_auth_oauthlib" >nul 2>&1
if errorlevel 1 (
    "%RETIE_PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo No se pudieron preparar las bibliotecas. Revisa Python y la conexion.
        pause
        exit /b 1
    )
)
"%RETIE_PYTHON%" aplicacion.py
if errorlevel 1 pause
