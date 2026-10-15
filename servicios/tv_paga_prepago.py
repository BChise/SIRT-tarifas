"""
Servicio: Televisión de paga Prepago.

Fuente: grabaciones/tv_paga_prepago/ (paso2.json, grabacion_paso2.py, captura, Excel).

Paso 2 (igual que Televisión de paga Postpago, salvo lo marcado):
- Valor del servicio, Duración (días, entero; reemplaza a Periodicidad),
  Tipo de Acceso (TV), Tecnologia.
- 3 campos "Carga y descarga" (Canales analógicos / SD / HD) con el mismo Excel.
- Decodificadores o similares (entero).
- Condiciones generales y Restricciones técnicas.

Mismo patrón que Telefonía fija: el prepago cambia Periodicidad (lista)
por Duración (número de días).
Este servicio tiene "Tarifa Instalación" en el paso 1 (se llena con el valor por defecto).
"""

from pathlib import Path

from comun import (
    cargar_archivo,
    completar_textos_finales,
    elegir_valor,
    elegir_varias_opciones,
)


SERVICIO = "Televisión de paga Prepago"   # texto exacto ("Prepago" con P mayúscula)
CODIGO = "TVPR"

CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
ARCHIVO_CANALES = CARPETA_PROYECTO / "grabaciones" / "tv_paga_prepago" / "Canales analógicos_1.xlsx"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    "tipo_acceso": {
        "opciones": ["Alta Nueva", "Migración"],   # "Alta Nueva" con N mayúscula
        "multiple": True,
    },
    "tecnologia": {
        "opciones": ["Satelital", "DOCSIS"],
        "multiple": False,   # selección única (confirmado en pantalla en TV postpago)
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

VALORES = {
    "valor_servicio": "50",            # hasta 4 decimales
    "duracion_dias": "30",             # entero
    "decodificadores": "2",            # entero

    # Campos de carga: el mismo archivo para los tres.
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
    elegir_varias_opciones(page, "Tipo de Acceso (TV)", valores["tipo_acceso"])
    elegir_valor(page, "Tecnologia", valores["tecnologia"])   # "Tecnologia" sin tilde

    for etiqueta, archivo in v["archivos"].items():
        cargar_archivo(page, etiqueta, archivo)

    page.get_by_role("spinbutton", name="Decodificadores o similares").fill(v["decodificadores"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
