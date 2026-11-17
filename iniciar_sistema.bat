@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto DirectoryError

title Sistema de Organizacion de Horarios Universitarios (Servicio Comunitario)
set "VENV_PYTHON=.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    echo [INFO] No se encontro el entorno virtual. Creando .venv...
    where py >nul 2>&1
    if not errorlevel 1 (
        py -3 --version >nul 2>&1
        if not errorlevel 1 (
            py -3 -m venv .venv
            if errorlevel 1 goto VenvError
            goto VenvReady
        )
    )

    where python >nul 2>&1
    if errorlevel 1 goto PythonError
    python --version >nul 2>&1
    if errorlevel 1 goto PythonError
    python -m venv .venv
    if errorlevel 1 goto VenvError
)

:VenvReady
if not exist "%VENV_PYTHON%" goto VenvError
if not exist "requirements.txt" goto RequirementsError

"%VENV_PYTHON%" -c "import flet" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Instalando dependencias necesarias...
    "%VENV_PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 goto DependenciesError
)

echo Abriendo interfaz grafica del Sistema de Horarios Universitarios...
"%VENV_PYTHON%" SGHU_Universidad.py
if errorlevel 1 (
    echo.
    echo [ERROR] La aplicacion termino con un error.
)
pause
exit /b

:DirectoryError
echo [ERROR] No se pudo acceder a la carpeta del sistema.
goto Failed

:PythonError
echo [ERROR] No se encontro una instalacion funcional de Python 3.
echo Instale Python 3 desde https://www.python.org/downloads/ y habilite "Add Python to PATH".
goto Failed

:VenvError
echo [ERROR] No se pudo crear o encontrar el entorno virtual .venv.
echo Verifique que Python 3 incluya el modulo venv y que tenga permisos en esta carpeta.
goto Failed

:RequirementsError
echo [ERROR] No se encontro requirements.txt en la carpeta del sistema.
goto Failed

:DependenciesError
echo [ERROR] No se pudieron instalar las dependencias.
echo Verifique su conexion a Internet y vuelva a ejecutar este archivo.
goto Failed

:Failed
pause
exit /b 1
