@echo off
REM ==========================================================================
REM  SIRT - prepara la consola para trabajar (doble clic o desde cmd)
REM  1. Entra a la carpeta del proyecto y activa el entorno .venv
REM  2. Activa UTF-8 (tildes y simbolos en la consola)
REM  3. Abre Edge controlable (puerto 9222) si no esta abierto, con opciones
REM     para que no se detenga aunque la ventana se minimice
REM  4. Deja la consola abierta lista para los comandos
REM ==========================================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo No existe el entorno .venv. Crealo con:
    echo     python -m venv .venv
    echo     .venv\Scripts\activate.bat
    echo     pip install -r requirements.txt
    pause
    exit /b 1
)
call ".venv\Scripts\activate.bat"
set PYTHONUTF8=1

netstat -ano | findstr ":9222" | findstr "LISTENING" >nul
if errorlevel 1 (
    echo Abriendo Edge en modo controlable ^(puerto 9222^)...
    echo Si Edge ya estaba abierto de forma normal, cierralo por completo y vuelve a ejecutar este archivo.
    start "" msedge --remote-debugging-port=9222 --user-data-dir=C:\perfiles\sirt --disable-backgrounding-occluded-windows --disable-renderer-backgrounding --disable-background-timer-throttling
) else (
    echo Edge ya esta abierto en modo controlable ^(puerto 9222^).
)

echo.
echo Inicia sesion en SIRT y deja la pantalla en "Tarifas registradas". Comandos:
echo   python registro_excel.py --solo-listar
echo   python empaquetados_excel.py --solo-listar
echo   python promociones_excel.py --solo-listar
echo   python licitaciones_excel.py --solo-listar
echo   python acciones_excel.py --solo-listar
echo   python modificaciones_excel.py --solo-listar
echo Los Excel estan en la carpeta automatizacion. Cierralos antes de correr un comando.
echo.
cmd /k
