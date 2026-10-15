"""
Servicio: Otros servicios.

Fuente: grabaciones/otros_servicios/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2:
- Tipo de moneda (NUEVO; sección "Datos generales", al inicio del paso 2).
- Valor del servicio, Periodicidad.
- Condiciones generales + Otras condiciones tarifa (opcional, NUEVO).
- Restricciones técnicas + Otras restricciones (opcional, NUEVO).

Nota: la etiqueta "Soles" de Valor del servicio no cambia al elegir otra moneda.
Este servicio tiene "Tarifa Instalación" en el paso 1 (se llena con el valor por defecto).
"""

from comun import (
    completar_textos_finales,
    elegir_opcion,
)


SERVICIO = "Otros servicios"   # texto exacto de "Seleccionar servicio"
CODIGO = "OS"


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
    "valor_servicio": "50",            # hasta 4 decimales

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
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])

    if v["otras_condiciones"]:
        page.get_by_role("textbox", name="Otras condiciones tarifa").fill(v["otras_condiciones"])
    if v["otras_restricciones"]:
        page.get_by_role("textbox", name="Otras restricciones").fill(v["otras_restricciones"])
