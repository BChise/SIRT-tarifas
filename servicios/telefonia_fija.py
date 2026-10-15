"""
Servicio: Telefonía fija postpago.

Paso 2:
- Valor del servicio, Periodicidad, Tipo de acceso, Tecnologia.
- 5 bloques de minutos ("condicional triple": Es ilimitado / minutos + tarifa).
- 2 tablas de tarifas por países: fijo LDI y móvil LDI.
- Condiciones generales y Restricciones técnicas.

Este servicio no tiene "Tarifa Instalación" (el campo queda bloqueado en el paso 1).
"""

from comun import (
    agregar_tarifa_paises,
    completar_condicional_triple,
    completar_textos_finales,
    elegir_opcion,
    elegir_valor,
    elegir_varias_opciones,
)


SERVICIO = "Telefonía fija postpago"   # texto exacto de "Seleccionar servicio"
CODIGO = "TF"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según el JSON del servicio.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "periodicidad": {
        "opciones": ["Diario", "Mensual", "Trimestral", "Semestral", "Anual"],
        "multiple": False,
    },
    "tipo_acceso": {
        # Ojo: aquí es "Alta nueva" (n minúscula), distinto de Internet fijo.
        "opciones": ["Alta nueva", "Migración", "Portabilidad"],
        "multiple": True,
    },
    "tecnologia": {
        "opciones": ["HFC", "Tecnología 1", "Tecnología 2"],
        "multiple": False,   # confirmado: selección única
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "30",            # hasta 4 decimales

    # Bloques de minutos. Modos: "omitir", "ilimitado" o "limitado".
    # minutos: entero; tarifa: hasta 4 decimales.
    "minutos": {
        "Minutos todo destino nacional": {"modo": "ilimitado"},
        "Minutos red fija on net":       {"modo": "limitado", "minutos": "100", "tarifa": "0.10"},
        "Minutos red fija off net":      {"modo": "ilimitado"},
        "Minutos móvil on net":          {"modo": "ilimitado"},
        "Minutos móvil off net":         {"modo": "limitado", "minutos": "100", "tarifa": "0.10"},
    },

    # Tablas de tarifas por países (opcionales). Lista vacía = no se agrega fila.
    # "Tipo" solo tiene la opción "Bloque".
    # paises: nombres en mayúsculas, tal como aparecen en la lista.
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

    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])
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
