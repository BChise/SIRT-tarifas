"""
Servicio: Paquetes adicionales y otras tarifas de televisión de paga.

Fuente: grabaciones/paquetes_tv_paga/ (paso2.json, grabacion_paso2.py, capturas).

Paso 2 (parecido a Televisión de paga Prepago):
- Valor del servicio, Duración (días, entero).
- Tipo de Contratación (NUEVO; múltiple: Postpago / Prepago; opcional).
- 3 campos "Carga y descarga" (Canales analógicos / SD / HD; aquí opcionales),
  cargados con el mismo Excel que TV de paga.
- Decodificadores o similares (opcional).
- Condiciones generales y Restricciones técnicas.

El Excel se toma de grabaciones/tv_paga_postpago/ (no hay copia en esta carpeta).
Este servicio tiene "Tarifa Instalación" en el paso 1 (se llena con el valor por defecto).
"""

from pathlib import Path

from comun import (
    cargar_archivo,
    completar_textos_finales,
    elegir_varias_opciones,
)


SERVICIO = "Paquetes adicionales y otras tarifas de televisión de paga"   # texto exacto
CODIGO = "PTV"

CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
ARCHIVO_CANALES = CARPETA_PROYECTO / "grabaciones" / "tv_paga_postpago" / "Canales analógicos_1.xlsx"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "tipo_contratacion": {
        "opciones": ["Postpago", "Prepago"],
        "multiple": True,   # en pantalla se ven las dos juntas: "Postpago, Prepago"
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "50",            # hasta 4 decimales
    "duracion_dias": "30",             # entero
    "decodificadores": "2",            # opcional; "" para dejarlo vacío

    # Campos de carga (opcionales aquí): {} para no cargar ninguno.
    "archivos": {
        "Canales analógicos":                    ARCHIVO_CANALES,
        "Canales digitales definición estándar": ARCHIVO_CANALES,
        "Canales digitales alta definición":     ARCHIVO_CANALES,
    },

    "condiciones_generales": "Condiciones generales de prueba automatizada",    # máx. 2000
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",  # máx. 500
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])
    elegir_varias_opciones(page, "Tipo de Contratación", valores["tipo_contratacion"])

    for etiqueta, archivo in v["archivos"].items():
        cargar_archivo(page, etiqueta, archivo)

    if v["decodificadores"]:
        page.get_by_role("spinbutton", name="Decodificadores o similares").fill(v["decodificadores"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
