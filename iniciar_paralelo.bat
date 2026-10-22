@echo off
REM ==========================================================================
REM  SIRT - SEGUNDA consola en paralelo (otro Edge, otro perfil, puerto 9223)
REM  Usala para correr un comando a la vez con otro (p. ej. registro mientras
REM  corren las acciones). Inicia sesion en SIRT tambien en este segundo Edge.
REM  No corras en paralelo cosas que usen las mismas tarifas ni el mismo Excel.
REM ==========================================================================
cd /d "%~dp0"
call ".venv\Scripts\activate.bat"
set PYTHONUTF8=1
set SIRT_PUERTO=9223
start "" msedge --remote-debugging-port=9223 --user-data-dir=C:\perfiles\sirt2 --disable-backgrounding-occluded-windows --disable-renderer-backgrounding --disable-background-timer-throttling
echo.
echo Consola en PARALELO: los comandos de esta ventana usan el Edge del puerto 9223.
echo Inicia sesion en SIRT en el Edge que se abrio y deja la pantalla en "Tarifas registradas".
echo.
cmd /k
