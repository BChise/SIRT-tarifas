@echo off
rem Prueba funcional - DIA 2 (acciones, modificaciones). Desde el dia siguiente al dia 1.
rem   prueba_dia2.bat Prueba4      corre el dia 2 de la carpeta Prueba4
rem   prueba_dia2.bat              corre el dia 2 con los Excel de pruebas_funcionales
cd /d "%~dp0"
call .venv\Scripts\activate.bat
if "%~1"=="" (python prueba_funcional.py --dia 2) else (python prueba_funcional.py --dia 2 --carpeta %1)
pause
