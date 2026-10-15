"""
Servicio: Paquetes adicionales y otros servicios de internet fijo.

Fuente: grabaciones/paquetes_internet_fijo/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2 (parecido a Internet fijo postpago, más simple):
- Valor del servicio, Duración (días, entero), Tecnología (IF) (opcional).
- Velocidades de navegación: una sola tabla (ventana con Tecnología "Única"
  + 4 velocidades), opcional; no hay tabla de degradación.
- Cantidad de datos (GB) y Repetidores y/o equipos similares (opcionales).
- Condiciones generales y Restricciones técnicas.

En la grabación se abrió una 2.ª ventana de velocidades y se cerró sin guardar:
no forma parte del flujo.
Este servicio tiene "Tarifa Instalación" en el paso 1 (se llena con el valor por defecto).
"""

from comun import (
    agregar_velocidad,
    completar_textos_finales,
    elegir_opcion,
)


SERVICIO = "Paquetes adicionales y otros servicios de internet fijo"   # texto exacto
CODIGO = "PIF"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "tecnologia_if": {
        "opciones": ["FTTH", "4G", "4.5G", "5G", "Satelital", "VHDL", "HFC"],
        "multiple": False,   # "Categorica Excluyente" en el JSON
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "50",            # hasta 4 decimales
    "duracion_dias": "30",             # entero
    "velocidad_navegacion": {          # None = no agregar fila (la tabla es opcional)
        "tecnologia": "Única",
        "descarga_contratada": "200",
        "descarga_garantizada": "200",
        "carga_contratada": "200",
        "carga_garantizada": "200",
    },
    "cantidad_datos_gb": "100",        # opcional; "" para dejarlo vacío
    "repetidores": "1",                # opcional; "" para dejarlo vacío

    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])
    elegir_opcion(page, "Tecnología (IF)", valores["tecnologia_if"])

    if v["velocidad_navegacion"]:
        agregar_velocidad(page, 0, v["velocidad_navegacion"])

    if v["cantidad_datos_gb"]:
        page.get_by_role("spinbutton", name="Cantidad de datos").fill(v["cantidad_datos_gb"])
    if v["repetidores"]:
        page.get_by_role("spinbutton", name="Repetidores y/o equipos").fill(v["repetidores"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
