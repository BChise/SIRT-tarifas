"""
Servicio: Arrendamiento de circuitos.

Fuente: grabaciones/arrendamiento_circuitos/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2:
- Tipo de moneda (sección "Datos generales", como en Otros servicios).
- Valor del servicio, Periodo de contrato (meses, entero), Periodicidad.
- Tecnología de transmisión (NUEVO: texto libre, máx. 40 caracteres).
- Velocidad de transmisión y Vel. transmisión garant. (Mbps, hasta 4 decimales).
- Condiciones generales + Otras condiciones tarifa (opcional).
- Restricciones técnicas + Otras restricciones (opcional).

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from comun import (
    completar_textos_finales,
    elegir_opcion,
)


SERVICIO = "Arrendamiento de circuitos"   # texto exacto
CODIGO = "AC"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "moneda": {
        "opciones": ["Soles", "Dolares", "Euro"],   # "Dolares" sin tilde
        "multiple": False,
    },
    "periodicidad": {
        "opciones": ["Diario", "Mensual", "Trimestral", "Semestral", "Anual"],
        "multiple": False,
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "50",                       # hasta 4 decimales
    "periodo_contrato_meses": "12",               # entero
    "tecnologia_transmision": "Fibra optica",     # texto libre, máx. 40 caracteres
    "velocidad_transmision": "100",               # Mbps, hasta 4 decimales
    "velocidad_garantizada": "80",                # Mbps, hasta 4 decimales

    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "otras_condiciones": "Otras condiciones de prueba automatizada",            # opcional, máx. 1000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
    "otras_restricciones": "Otras restricciones de prueba automatizada",        # opcional, máx. 1000
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    elegir_opcion(page, "Tipo de moneda", valores["moneda"])
    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    page.get_by_role("spinbutton", name="Periodo de contrato").fill(v["periodo_contrato_meses"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])

    page.get_by_role("textbox", name="Tecnología de transmisión").fill(v["tecnologia_transmision"])
    page.get_by_role("textbox", name="Velocidad de transmisión").fill(v["velocidad_transmision"])
    page.get_by_role("textbox", name="Vel. transmisión garant.").fill(v["velocidad_garantizada"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])

    if v["otras_condiciones"]:
        page.get_by_role("textbox", name="Otras condiciones tarifa").fill(v["otras_condiciones"])
    if v["otras_restricciones"]:
        page.get_by_role("textbox", name="Otras restricciones").fill(v["otras_restricciones"])
