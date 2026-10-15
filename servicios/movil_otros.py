"""
Servicio: Servicio móvil otros.

Fuente: grabaciones/movil_otros/ (paso2.json, grabacion_paso2.py, 6 capturas).

Es el formulario de móvil COMPLETO: tiene los campos de postpago
(Periodicidad, Tipo de acceso) y los de paquete (Duración, Tipo de línea móvil)
a la vez, y casi todo es opcional. Se llenan todos, como en la grabación.

Solo se escribe el inicio; desde "Componentes del plan" se reutiliza
movil_postpago.completar_resto_movil (igual en los 4 servicios móviles).

La grabación confirma la regla de móvil: tethering y traslada datos en 14 con
tope 12 dieron error al guardar; con 12 pasó.

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from comun import elegir_opcion, elegir_varias_opciones, elegir_valor
from servicios import movil_postpago


SERVICIO = "Servicio móvil otros"   # texto exacto
CODIGO = "SMO"


# --- Listas que se combinan --------------------------------------------------
# Combinando las cuatro listas saldrían 3 x 5 x 7 x 7 = 735 casos. Por ahora se
# combina solo Tipo de línea móvil (lo propio de este servicio, como en paquete)
# y las otras dos quedan fijas; para combinarlas, poner sus opciones completas.

DIMENSIONES = {
    "alcance": movil_postpago.DIMENSIONES["alcance"],
    "periodicidad": {
        "opciones": ["Mensual"],
        # Todas: ["Diario", "Mensual", "Trimestral", "Semestral", "Anual"]
        "multiple": False,
    },
    "tipo_acceso": {
        "opciones": ["Alta nueva", "Migración", "Portabilidad"],
        "multiple": True,
        "combinar": False,   # siempre las tres juntas (como en la grabación)
    },
    "tipo_linea_movil": {
        "opciones": ["Postpago", "Control", "Prepago"],
        "multiple": True,
    },
    "componentes": movil_postpago.DIMENSIONES["componentes"],
    "internet_ilimitado": movil_postpago.DIMENSIONES["internet_ilimitado"],
}


# --- Valores fijos del paso 2 ------------------------------------------------
# Los mismos de postpago, más la duración.

VALORES = {
    **movil_postpago.VALORES,
    "duracion_dias": "30",             # entero
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    # 1. Datos del plan: los de postpago y los de paquete
    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])
    elegir_valor(page, "Tipo de acceso", valores["tipo_acceso"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])
    elegir_varias_opciones(page, "Tipo de línea móvil", valores["tipo_linea_movil"])

    # 2. Todo lo demás, igual que postpago
    movil_postpago.completar_resto_movil(page, valores, v)
