"""
Servicio: Internet fijo postpago.

Solo contiene lo propio de este servicio:
- DIMENSIONES: las listas que se combinan para generar los casos.
- VALORES: los datos fijos del paso 2.
- completar_paso_2(): cómo se llena el paso 2.
"""

from comun import (
    agregar_velocidad,
    completar_textos_finales,
    elegir_opcion,
    elegir_varias_opciones,
)


SERVICIO = "Internet fijo postpago"   # texto exacto de "Seleccionar servicio"
CODIGO = "IF"                         # se usa en nombres de tarifa y carpetas

# Opcional: cambia valores del paso 1 solo para este servicio.
# PASO1 = {"tarifa_instalacion": "50"}


# --- Listas que se combinan --------------------------------------------------
# "alcance" es obligatoria en todos los servicios (pertenece al paso 1).
# Una lista con una sola opción = campo fijo.

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
        "opciones": ["Alta Nueva", "Migración"],
        "multiple": True,
    },
    "tecnologia_if": {
        "opciones": ["HFC"],
        # Todas: ["FTTH", "4G", "4.5G", "5G", "Satelital", "VHDL", "HFC"]
        "multiple": False,
    },
    "internet_ilimitado": {
        "opciones": ["Si"],
        # Ambas (sin tilde, como en la lista): ["Si", "No"]
        "multiple": False,
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "100",          # hasta 4 decimales
    "cantidad_datos_gb": "100",       # solo cuando Internet Ilimitado = "No"
    "repetidores": "1",               # opcional; "" para dejarlo vacío
    "velocidad_navegacion": {
        "tecnologia": "Única",
        "descarga_contratada": "200",
        "descarga_garantizada": "200",
        "carga_contratada": "200",
        "carga_garantizada": "200",
    },
    "agregar_degradacion": True,      # la tabla es opcional
    "velocidad_degradacion": {
        "tecnologia": "Única",
        "descarga_contratada": "30",
        "descarga_garantizada": "30",
        "carga_contratada": "30",
        "carga_garantizada": "30",
    },
    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    """valores: los valores del caso (las claves de DIMENSIONES)."""
    v = VALORES

    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])
    elegir_varias_opciones(page, "Tipo de Acceso (IF)", valores["tipo_acceso"])
    elegir_opcion(page, "Tecnología (IF)", valores["tecnologia_if"])

    agregar_velocidad(page, 0, v["velocidad_navegacion"])

    elegir_opcion(page, "Internet Ilimitado", valores["internet_ilimitado"])

    if valores["internet_ilimitado"] == "No":
        page.get_by_role("spinbutton", name="Cantidad de datos").fill(v["cantidad_datos_gb"])

    if v["agregar_degradacion"]:
        agregar_velocidad(page, 1, v["velocidad_degradacion"])

    if v["repetidores"]:
        page.get_by_role("spinbutton", name="Repetidores y/o equipos").fill(v["repetidores"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
