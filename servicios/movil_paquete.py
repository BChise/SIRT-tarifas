"""
Servicio: Servicio público móvil paquete.

Fuente: grabaciones/movil_paquete/ (paso2.json, grabacion_paso2.py, 5 capturas).

El paso 2 es igual al de Servicio público móvil postpago, salvo el inicio:
- Duración (días, entero) en lugar de Periodicidad.
- Tipo de línea móvil (múltiple: Postpago / Control / Prepago) en lugar de
  Tipo de acceso.
- Varios campos que en postpago son obligatorios aquí son opcionales
  (velocidades, degradación, tarifas por minuto); se llenan igual.

Por eso solo se escribe el inicio y el resto se reutiliza de movil_postpago.py
(completar_resto_movil): una corrección allí vale para los tres servicios móviles.

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from comun import elegir_varias_opciones
from servicios import movil_postpago


SERVICIO = "Servicio público móvil paquete"   # texto exacto
CODIGO = "SMQ"


# --- Listas que se combinan --------------------------------------------------
# Textos exactos según paso2.json.

DIMENSIONES = {
    "alcance": movil_postpago.DIMENSIONES["alcance"],
    "tipo_linea_movil": {
        "opciones": ["Postpago", "Control", "Prepago"],
        "multiple": True,   # en pantalla se ven varias juntas: "Postpago, Control"
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

    # 1. Datos del paquete (lo propio de este servicio)
    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    page.get_by_role("spinbutton", name="Duración").fill(v["duracion_dias"])
    elegir_varias_opciones(page, "Tipo de línea móvil", valores["tipo_linea_movil"])

    # 2. Todo lo demás, igual que postpago
    movil_postpago.completar_resto_movil(page, valores, v)
