"""
Servicio: Paquetes adicionales de telefonía fija.

Fuente: grabaciones/paquetes_telefonia_fija/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2:
- Valor del servicio y Duración (días, entero); ambos opcionales aquí.
- Tipo de línea (NUEVO; múltiple: Postpago / Prepago; opcional).
- 2 tablas de tarifas por países: fijo LDI y móvil LDI (opcionales).
- Condiciones generales y Restricciones técnicas.

No tiene bloques de minutos (a diferencia de Telefonía fija).
Este servicio tiene "Tarifa Instalación" en el paso 1 (se llena con el valor por defecto).
"""

from comun import (
    agregar_tarifa_paises,
    completar_textos_finales,
    elegir_varias_opciones,
)


SERVICIO = "Paquetes adicionales de telefonía fija"   # texto exacto
CODIGO = "PTF"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "tipo_linea": {
        "opciones": ["Postpago", "Prepago"],
        "multiple": True,   # en pantalla se ven las dos juntas: "Postpago, Prepago"
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "30",            # opcional; hasta 4 decimales
    "duracion_dias": "30",             # opcional; entero

    # Tablas de tarifas por países (opcionales). Lista vacía = no se agrega fila.
    "ldi_fijo": [
        {"tipo": "Bloque", "tarifa": "25", "paises": ["ALEMANIA"]},
    ],
    "ldi_movil": [
        {"tipo": "Bloque", "tarifa": "124", "paises": ["ALEMANIA"]},
    ],

    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    if v["valor_servicio"]:
        page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    if v["duracion_dias"]:
        page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])

    elegir_varias_opciones(page, "Tipo de línea", valores["tipo_linea"])

    # Botones "Agregar": 0 = fijo LDI, 1 = móvil LDI.
    for fila in v["ldi_fijo"]:
        agregar_tarifa_paises(page, 0, fila)

    for fila in v["ldi_movil"]:
        agregar_tarifa_paises(page, 1, fila)

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
