"""
Servicio: Otras tarifas de telefonía fija.

Fuente: grabaciones/otras_tarifas_telefonia_fija/ (paso2.json, grabacion_paso2.py, captura).

Paso 2:
- Valor del servicio (opcional), Periodicidad (opcional), Duración (días, entero).
- Tipo de línea (múltiple: Postpago / Prepago).
- Tarifa por minuto (NUEVO; Soles, hasta 4 decimales).
- Condiciones generales y Restricciones técnicas.

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from comun import (
    completar_textos_finales,
    elegir_opcion,
    elegir_varias_opciones,
)


SERVICIO = "Otras tarifas de telefonía fija"   # texto exacto
CODIGO = "OTF"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "periodicidad": {
        "opciones": ["Diario", "Mensual", "Trimestral", "Semestral", "Anual"],
        "multiple": False,
    },
    "tipo_linea": {
        "opciones": ["Postpago", "Prepago"],
        "multiple": True,   # en pantalla se ven las dos juntas: "Postpago, Prepago"
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "30",            # opcional; hasta 4 decimales
    "duracion_dias": "30",             # entero
    "tarifa_por_minuto": "0.10",       # hasta 4 decimales

    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    if v["valor_servicio"]:
        page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])
    elegir_varias_opciones(page, "Tipo de línea", valores["tipo_linea"])
    page.get_by_role("textbox", name="Tarifa por minuto").fill(v["tarifa_por_minuto"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
