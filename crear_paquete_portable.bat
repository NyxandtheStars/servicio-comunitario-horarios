@echo off
cd /d "%~dp0"
py -3 crear_paquete_portable.py
if errorlevel 1 python crear_paquete_portable.py
pause
