@echo off
rem Prueba funcional - DIA 1 (registro, empaquetados, LyC, promociones).
rem   prueba_dia1.bat P4 Prueba4   genera la serie P4 en la carpeta Prueba4 y corre el dia 1
rem   prueba_dia1.bat              corre el dia 1 con los Excel de pruebas_funcionales
cd /d "%~dp0"
call .venv\Scripts\activate.bat
if "%~1"=="" (
    python prueba_funcional.py --dia 1
) else if "%~2"=="" (
    python prueba_funcional.py --dia 1 --serie %1
) else (
    python prueba_funcional.py --dia 1 --serie %1 --carpeta %2
)
pause
