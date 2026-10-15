#!/usr/bin/env bash
# Ejecuta ejecutar.py para todos los servicios, uno tras otro, y muestra un
# resumen al final. Si un servicio falla, sigue con el siguiente.
#
# Uso (desde cualquier carpeta, con el entorno virtual activado):
#     bash probar_todos.sh                  # lote MUESTRA, 1 caso por servicio, con diagnóstico
#     bash probar_todos.sh MUESTRA 1        # lo mismo, explícito
#     bash probar_todos.sh PRUEBA2 2        # otro lote, 2 casos por servicio
#     bash probar_todos.sh MUESTRA 1 --trace   # opciones extra para ejecutar.py
#
# Requisitos: Edge abierto en modo controlable, sesión iniciada en SIRT y
# pantalla en "Tarifas registradas" (igual que ejecutar.py).
#
# La salida completa queda también en resultados/probar_todos_<lote>_<fecha>.log

LOTE="${1:-MUESTRA}"
LIMITE="${2:-1}"
shift 2 2>/dev/null
EXTRA="${*:---diagnostico}"     # sin opciones extra: --diagnostico
PYTHON="${PYTHON:-python}"

SERVICIOS=(
    telefonia_fija
    telefonia_fija_prepago
    otras_tarifas_telefonia_fija
    paquetes_telefonia_fija
    internet_fijo
    paquetes_internet_fijo
    tv_paga_postpago
    tv_paga_prepago
    paquetes_tv_paga
    movil_postpago
    movil_control
    movil_paquete
    movil_otros
    arrendamiento_circuitos
    otros_servicios
)

cd "$(dirname "$0")" || exit 1
mkdir -p resultados
LOG="resultados/probar_todos_${LOTE}_$(date +%Y%m%d-%H%M).log"

echo "Lote: $LOTE | Casos por servicio: $LIMITE | Opciones: $EXTRA"
echo "Servicios: ${#SERVICIOS[@]} | Registro: $LOG"
echo

declare -a RESUMEN

for servicio in "${SERVICIOS[@]}"; do
    echo "######################################################################" | tee -a "$LOG"
    echo "# $servicio" | tee -a "$LOG"
    echo "######################################################################" | tee -a "$LOG"

    # Salida en vivo en pantalla y, a la vez, en el log y en un archivo temporal
    # para armar el resumen. PYTHONUNBUFFERED: que se vea al momento (incluida
    # la pregunta de confirmación de ejecutar.py si se usa sin límite).
    temporal=$(mktemp)
    PYTHONUNBUFFERED=1 "$PYTHON" ejecutar.py "$servicio" --limite "$LIMITE" --lote "$LOTE" $EXTRA 2>&1 \
        | tee -a "$LOG" "$temporal"
    codigo=${PIPESTATUS[0]}
    salida=$(cat "$temporal")
    rm -f "$temporal"

    confirmadas=$(echo "$salida" | grep -oP "Confirmadas en esta ejecución:\s*\K\d+" | tail -1)
    errores=$(echo "$salida" | grep -oP "Errores en esta ejecución:\s*\K\d+" | tail -1)

    if [ "$codigo" -ne 0 ] && [ -z "$confirmadas" ]; then
        estado="FALLÓ (el programa se detuvo, ver log)"
    elif echo "$salida" | grep -q "No hay casos pendientes"; then
        estado="sin pendientes (ya creado en este lote)"
    elif [ "${errores:-0}" -gt 0 ]; then
        estado="ERROR en ${errores} caso(s), confirmadas: ${confirmadas:-0}"
    else
        estado="OK, confirmadas: ${confirmadas:-0}"
    fi

    RESUMEN+=("$(printf '%-30s %s' "$servicio" "$estado")")
    echo | tee -a "$LOG"
done

{
    echo "======================================================================"
    echo "RESUMEN (lote $LOTE)"
    echo "======================================================================"
    printf '%s\n' "${RESUMEN[@]}"
    echo
    echo "Detalle de cada servicio: resultados/<CODIGO>/resultados_${LOTE}.csv"
    echo "Registro completo: $LOG"
} | tee -a "$LOG"
