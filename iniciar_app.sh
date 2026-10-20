#!/usr/bin/env bash
# Arranca la app web de la herramienta SIRT.
#
# Uso (desde cualquier carpeta):
#     bash iniciar_app.sh            # instala lo que falte, abre Edge si no está y arranca la app
#     bash iniciar_app.sh --sin-edge # no intenta abrir Edge
#     PUERTO=8080 bash iniciar_app.sh
#
# Hace, en orden:
#   1. Se ubica en la carpeta del proyecto.
#   2. Activa el entorno virtual register_web_env (si existe y no está activo).
#   3. Instala fastapi, uvicorn y openpyxl si faltan.
#   4. Si Edge no responde en el puerto 9222, lo abre con el perfil de SIRT
#      (inicia sesión en SIRT en esa ventana si hace falta).
#   5. Arranca la app en http://127.0.0.1:${PUERTO:-8000} y abre el navegador.

set -e
cd "$(dirname "$0")"
PROYECTO="$(pwd)"
PUERTO="${PUERTO:-8000}"

echo "== Herramienta SIRT =="
echo "Proyecto: $PROYECTO"

# 2. Entorno virtual
if [ -z "$VIRTUAL_ENV" ]; then
    for ENTORNO in "$PROYECTO/../../register_web_env" "$PROYECTO/../register_web_env" "$PROYECTO/register_web_env"; do
        if [ -f "$ENTORNO/bin/activate" ]; then
            # shellcheck disable=SC1091
            source "$ENTORNO/bin/activate"
            echo "Entorno virtual: $ENTORNO"
            break
        fi
    done
fi
[ -z "$VIRTUAL_ENV" ] && echo "AVISO: no se encontró register_web_env; se usa el Python del sistema."
PYTHON="${PYTHON:-python}"

# 3. Dependencias de la app
if ! "$PYTHON" -c "import fastapi, uvicorn, openpyxl" 2>/dev/null; then
    echo "Instalando fastapi, uvicorn y openpyxl…"
    "$PYTHON" -m pip install --quiet fastapi uvicorn openpyxl
fi

# 4. Edge en modo controlable
if [ "$1" != "--sin-edge" ]; then
    if curl -s --max-time 2 http://localhost:9222/json/version >/dev/null 2>&1; then
        echo "Edge ya está abierto en el puerto 9222."
    elif command -v microsoft-edge >/dev/null 2>&1; then
        echo "Abriendo Edge (puerto 9222, perfil ~/perfiles/sirt)…"
        microsoft-edge --remote-debugging-port=9222 --user-data-dir="$HOME/perfiles/sirt" \
            "https://sirt-qa.be1business.com/extranet/consulta-tarifa" >/dev/null 2>&1 &
        sleep 3
        echo ">> Inicia sesión en SIRT en esa ventana si te lo pide."
    else
        echo "AVISO: no se encontró 'microsoft-edge'. Ábrelo a mano con --remote-debugging-port=9222."
    fi
fi

# 5. App
echo
"$PYTHON" -m app.servidor
