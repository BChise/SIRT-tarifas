# 1. Guía de instalación (Windows)

Tiempo aproximado: 15 minutos. Se hace **una sola vez por computadora**.

## Requisitos

| Qué | Versión probada | Notas |
|---|---|---|
| Windows | 10 / 11 | |
| Python | 3.13.2 | Cualquier 3.10 o superior debería funcionar. Al instalarlo, marcar **«Add python.exe to PATH»**. |
| Microsoft Edge | el que trae Windows | Los scripts manejan Edge; **no** se descarga otro navegador. |
| Git | cualquiera | Solo para obtener el proyecto (también puede llegar en un .zip). |
| Acceso a SIRT QA | usuario del operador | El inicio de sesión (usuario, contraseña, reCAPTCHA y código por correo) lo hace siempre la persona. |

> **Consola:** usar **cmd** (Símbolo del sistema). En equipos de la empresa PowerShell suele estar bloqueado para
> ejecutar scripts (`Activate.ps1` no se puede cargar). Todo lo de esta guía es para cmd.

## Paso 1 — Obtener el proyecto

```
cd /d D:\
mkdir proyecto_sirt
cd proyecto_sirt
git clone <URL del repositorio privado> sirt_tarifas
cd sirt_tarifas
```

(Si llegó como .zip, descomprimirlo en `D:\proyecto_sirt\sirt_tarifas`.) Se puede usar otra carpeta; en esta guía se
usa `D:\proyecto_sirt\sirt_tarifas`.

## Paso 2 — Crear el entorno de Python e instalar dependencias

Desde la carpeta del proyecto:

```
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

La línea del cmd debe empezar con `(.venv)`. `requirements.txt` instala:

- `playwright` 1.64.0 — maneja Edge.
- `openpyxl` 3.1.5 — lee y escribe los Excel.

No hace falta `playwright install`: el proyecto usa el Edge que ya está instalado.

## Paso 3 — Primer arranque

1. **Cerrar Edge por completo** (también los íconos de la bandeja del sistema). Si queda abierto «normal», no se puede
   abrir en modo controlable.
2. Doble clic en **`iniciar.bat`** (o escribirlo en el cmd). Hace todo:
   - entra a la carpeta y activa `.venv`;
   - activa UTF-8 (tildes en la consola);
   - si no está abierto ya, abre Edge en modo controlable (puerto 9222, perfil `C:\perfiles\sirt`) con las opciones
     para que no se frene cuando la ventana queda tapada;
   - deja el cmd listo para los comandos.
3. En ese Edge, **iniciar sesión en SIRT** y dejar la pantalla en **«Tarifas registradas»**.

Desde la segunda vez, el perfil `C:\perfiles\sirt` recuerda los datos del navegador; el inicio de sesión de SIRT
puede pedirse de nuevo cuando la sesión vence.

## Paso 4 — Comprobar que todo funciona (sin tocar SIRT)

```
python registro_excel.py --solo-listar
```

Debe mostrar las hojas de servicio del Excel y cuántas tarifas registraría. `--solo-listar` **nunca** toca SIRT ni el
Excel.

Luego, una prueba real mínima: ver el [manual de usuario](02_manual_usuario.md) («Probar con 1 fila»).

## Problemas típicos de instalación

| Síntoma | Solución |
|---|---|
| `python` no se reconoce | Reinstalar Python marcando «Add python.exe to PATH», o usar `py` en lugar de `python`. |
| `Activate.ps1 no se puede cargar` | Se está usando PowerShell. Abrir **cmd** y usar `.venv\Scripts\activate.bat`. |
| `iniciar.bat` dice que Edge está abierto pero no responde | Había un Edge «normal». Cerrarlo del todo (Administrador de tareas → Microsoft Edge → Finalizar) y volver a abrir `iniciar.bat`. |
| `No se pudo conectar con Edge (puerto 9222)` | Edge no está abierto en modo controlable. Abrirlo con `iniciar.bat`. |
| `pip install` falla por proxy | Pedir a TI el proxy de la empresa y usar `pip install --proxy http://usuario:clave@proxy:puerto -r requirements.txt`. |
