"""
Servicio: Telefonía fija prepago.

Fuente: grabaciones/telefonia_fija_prepago/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2 (igual que Telefonía fija postpago, salvo lo marcado):
- Valor del servicio, Duración (NUEVO: días, entero; reemplaza a Periodicidad),
  Tipo de acceso, Tecnologia.
- 5 bloques de minutos ("condicional triple": Es ilimitado / minutos + tarifa).
- 2 tablas de tarifas por países: fijo LDI y móvil LDI.
- Condiciones generales y Restricciones técnicas.

Este servicio SÍ tiene "Tarifa Instalación" en el paso 1 (registraTarifaInstalacion
= true en el JSON); se llena con el valor por defecto de comun.PASO1_POR_DEFECTO.
"""

from comun import (
    agregar_tarifa_paises,
    completar_condicional_triple,
    completar_textos_finales,
    elegir_valor,
    elegir_varias_opciones,
)


SERVICIO = "Telefonía fija prepago"   # texto exacto de "Seleccionar servicio"
CODIGO = "TFP"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "tipo_acceso": {
        "opciones": ["Alta nueva", "Migración", "Portabilidad"],   # "Alta nueva" con n minúscula
        "multiple": True,
    },
    "tecnologia": {
        "opciones": ["HFC", "Tecnología 1", "Tecnología 2"],
        "multiple": False,   # mismo criterio que postpago: selección única
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "30",            # hasta 4 decimales
    "duracion_dias": "30",             # entero

    # Bloques de minutos. Modos: "omitir", "ilimitado" o "limitado".
    "minutos": {
        "Minutos todo destino nacional": {"modo": "ilimitado"},
        "Minutos red fija on net":       {"modo": "limitado", "minutos": "100", "tarifa": "0.10"},
        "Minutos red fija off net":      {"modo": "ilimitado"},
        "Minutos móvil on net":          {"modo": "ilimitado"},
        "Minutos móvil off net":         {"modo": "limitado", "minutos": "100", "tarifa": "0.10"},
    },

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

    # Los nombres accesibles incluyen la unidad ("Valor del servicio Soles",
    # "Duración Días"); la búsqueda por nombre parcial los encuentra igual.
    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])

    elegir_varias_opciones(page, "Tipo de acceso", valores["tipo_acceso"])
    elegir_valor(page, "Tecnologia", valores["tecnologia"])   # "Tecnologia" sin tilde

    for etiqueta, config in v["minutos"].items():
        completar_condicional_triple(page, etiqueta, config)

    # Botones "Agregar": 0 = fijo LDI, 1 = móvil LDI.
    for fila in v["ldi_fijo"]:
        agregar_tarifa_paises(page, 0, fila)

    for fila in v["ldi_movil"]:
        agregar_tarifa_paises(page, 1, fila)

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
