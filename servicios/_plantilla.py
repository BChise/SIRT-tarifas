"""
PLANTILLA para un servicio nuevo.

1. Copia este archivo con el nombre del servicio, por ejemplo:
       servicios/telefonia_fija.py
2. Completa SERVICIO, CODIGO, DIMENSIONES y VALORES.
3. Escribe completar_paso_2() a partir de la grabación del paso 2.
4. Prueba:
       python ejecutar.py telefonia_fija --solo-calcular
       python ejecutar.py telefonia_fija --limite 1

Los archivos que empiezan con "_" no aparecen como servicios.
"""

from comun import (
    agregar_velocidad,          # tabla de velocidades, si el servicio la tiene
    completar_textos_finales,   # Condiciones generales + Restricciones técnicas
    elegir_opcion,              # lista de selección única
    elegir_varias_opciones,     # lista de selección múltiple
    llenar_si_editable,         # campo que a veces está bloqueado
)


SERVICIO = "Nombre exacto del servicio"   # texto exacto de "Seleccionar servicio"
CODIGO = "XX"                             # corto, sin espacios: TF, TV, SM...

# Opcional: cambia valores del paso 1 solo para este servicio.
# PASO1 = {"tarifa_instalacion": "50"}


DIMENSIONES = {
    # Obligatoria en todos los servicios (paso 1).
    "alcance": {
        "opciones": ["Residencial", "Comercial"],
        "multiple": True,
    },
    # Ejemplos:
    # "periodicidad": {"opciones": ["Mensual", "Anual"], "multiple": False},
    # "tipo_acceso":  {"opciones": ["Alta Nueva", "Migración"], "multiple": True},
}


VALORES = {
    "valor_servicio": "100",
    "condiciones_generales": "Condiciones generales de prueba automatizada",
    "restricciones_tecnicas": "Restricciones técnicas de prueba automatizada",
}


def completar_paso_2(page, valores: dict) -> None:
    """valores: los valores del caso (las claves de DIMENSIONES)."""
    v = VALORES

    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    # elegir_opcion(page, "Periodicidad", valores["periodicidad"])
    # elegir_varias_opciones(page, "Tipo de acceso", valores["tipo_acceso"])

    completar_textos_finales(page, v["condiciones_generales"], v["restricciones_tecnicas"])
