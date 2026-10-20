"""
Servicio: Servicio público móvil postpago.

Fuente: grabaciones/movil_postpago/ (paso2.json, grabacion_paso2.py, 13 capturas).

Es el formulario más grande hasta ahora (~45 campos). Orden en pantalla:
1. Valor del servicio, Periodicidad, Tipo de acceso, Componentes del plan.
2. Datos para tethering y Traslada datos (GB).
3. Tablas "Bolsa de gigas" (NUEVO): Bolsa de datos para apps, Datos para Apps 1-4.
4. Apps sin consumo de datos (múltiple).
5. Tabla "Tecnologías y velocidades".
6. Internet Ilimitado, Datos tope descarga, 4 velocidades con degradación.
7. Bloques "condicional triple" (minutos / SMS) y tarifas por minuto.
8. 9 tablas de tarifas por países (LDI y roaming) + datos roaming.
9. 10 textos de condiciones y 3 de restricciones.

Con tantas tablas, los botones "Agregar" se ubican por el TÍTULO de su sección
(no por su posición), con comun.boton_agregar.

REGLA DE SIRT (vista en las grabaciones): "Datos para tethering", "Traslada datos"
y "Datos Roaming incluidos en los datos del plan" no pueden ser mayores que
"Datos tope descarga". En las grabaciones dieron error al guardar: Traslada
datos = 24 con tope 12 (postpago) y Datos Roaming incluidos = 14 con tope 12 (paquete).

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from comun import (
    agregar_bolsa_datos,
    agregar_tarifa_paises,
    agregar_velocidad,
    completar_condicional_triple,
    elegir_opcion,
    elegir_valor,
    elegir_varias_opciones,
)


SERVICIO = "Servicio público móvil postpago"   # texto exacto
CODIGO = "SMP"


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
    "tipo_acceso": {
        "opciones": ["Alta nueva", "Migración", "Portabilidad"],   # "Alta nueva" con n minúscula
        "multiple": True,
    },
    "componentes": {
        "opciones": ["Voz y SMS", "Datos"],
        "multiple": True,
        # Por ahora siempre los dos (como en la grabación). Con uno solo, el
        # formulario podría ocultar campos de datos o de voz: falta verificarlo.
        "combinar": False,
    },
    "internet_ilimitado": {
        "opciones": ["Si"],
        # Ambas (sin tilde, como en la lista): ["Si", "No"]. Igual que Internet fijo.
        "multiple": False,
    },
}


# --- Valores fijos del paso 2 ------------------------------------------------

TABLA_PAIS = {"tipo": "Bloque", "tarifa": "1.50", "paises": ["ALEMANIA"]}

VALORES = {
    "valor_servicio": "50",            # hasta 4 decimales

    # Tope de datos y los dos campos que dependen de él (deben ser <= tope).
    "datos_tope_descarga_gb": "20",
    "datos_tethering_gb": "5",
    "traslada_datos_gb": "5",

    # Tablas "Bolsa de gigas" (opcionales). Lista vacía = no se agrega fila.
    "bolsas": {
        "Bolsa de datos para apps": [
            {"tipo": "Bolsa", "gigas": "10", "plataformas": ["Whatsapp", "Messenger"]},
        ],
        "Datos para Apps 1": [
            {"tipo": "Bolsa", "gigas": "5", "plataformas": ["Otros"]},
        ],
        "Datos para Apps 2": [],
        "Datos para Apps 3": [],
        "Datos para Apps 4": [],
    },

    "apps_sin_consumo": ("Otros", "Whatsapp"),   # () para dejarlo vacío (opcional)

    # Tabla "Tecnologías y velocidades" (obligatoria).
    "velocidades": [
        {
            "tecnologia": "5G",
            "descarga_contratada": "50",
            "descarga_garantizada": "40",
            "carga_contratada": "20",
            "carga_garantizada": "15",
        },
        # 19/10: una fila por tecnología (antes solo 5G). Si SIRT nombra distinto las
        # opciones de «Tecnología» (p. ej. «4.5G», «LTE»), ajustar aquí.
        {
            "tecnologia": "4G",
            "descarga_contratada": "30",
            "descarga_garantizada": "20",
            "carga_contratada": "10",
            "carga_garantizada": "8",
        },
        {
            "tecnologia": "3G",
            "descarga_contratada": "5",
            "descarga_garantizada": "3",
            "carga_contratada": "2",
            "carga_garantizada": "1",
        },
    ],

    # Velocidades con degradación (obligatorias; Mbps, hasta 4 decimales).
    "degradacion": {
        "Velocidad descarga contratada con degradación": "2",
        "Velocidad carga contratada con degradación": "2",
        "Velocidad descarga garantizada con degradación": "1",
        "Velocidad carga garantizada con degradación": "1",
    },

    # Bloques "condicional triple". Modos: "omitir", "ilimitado" o "limitado".
    # En los de SMS, "minutos" es la cantidad de mensajes.
    "condicionales": {
        "Minutos móvil nacional":       {"modo": "ilimitado"},
        "Cantidad SMS nacional":        {"modo": "limitado", "minutos": "100", "tarifa": "0.10"},
        "Cantidad SMS internacional":   {"modo": "limitado", "minutos": "10", "tarifa": "0.50"},
        "Minutos móvil on net":         {"modo": "ilimitado"},
        "Minutos móviles off net":      {"modo": "ilimitado"},
        "Minutos móvil red privada":    {"modo": "ilimitado"},
        "Minutos LDI dentro del plan":  {"modo": "ilimitado"},
        "Min. Roaming dentro del plan": {"modo": "ilimitado"},
        "SMS Roaming dentro del plan":  {"modo": "ilimitado"},
    },

    "tarifa_mb_adicional": "0.05",
    "tarifa_movil_a_fijo": "0.10",
    "tarifa_movil_todo_destino": "0.10",

    # Tablas de tarifas por países (opcionales), por título de sección.
    "tablas_paises": {
        "Tarifa por minuto LDI":          [TABLA_PAIS],
        "Tarifa minuto roaming":          [TABLA_PAIS],
        "Tarifa SMS roaming":             [TABLA_PAIS],
        "Tarifa MB roaming":              [TABLA_PAIS],
        "Minutos LDI países":             [TABLA_PAIS],
        "Datos Roaming del plan países":  [TABLA_PAIS],
        "Datos Roaming adicional países": [TABLA_PAIS],
        "Minutos Roaming países":         [TABLA_PAIS],
        "SMS Roaming países":             [TABLA_PAIS],
    },

    "datos_roaming_incluidos_gb": "2",      # opcional; "" para dejarlo vacío
    "datos_roaming_adicionales_gb": "",     # opcional; vacío como en la grabación

    # Textos (nombre exacto del campo → texto). Obligatorios salvo "Otras ...".
    "textos": {
        "Condiciones generales": "Condiciones generales de prueba automatizada",            # 2000
        "Condiciones cantidad datos": "Condiciones de datos de prueba automatizada",        # 1000
        "Condiciones bolsa de datos para apps": "Condiciones de bolsa de prueba",           # 1000
        "Condiciones apps zero rating": "Condiciones zero rating de prueba",                # 500
        "Condiciones datos tethering": "Condiciones tethering de prueba",                   # 500
        "Condiciones traslada datos": "Condiciones traslada datos de prueba",              # 500
        "Condiciones velocidades": "Condiciones velocidades de prueba",                    # 500
        "Condiciones Roaming internacional": "Condiciones roaming de prueba",              # 500
        "Condiciones LDI": "Condiciones LDI de prueba",                                     # 500
        "Otras condiciones tarifa": "Otras condiciones de prueba automatizada",            # 1000, opcional
        "Restricciones técnicas": "Restricciones técnicas de prueba automatizada",         # 500
        "Restricciones geográficas": "Restricciones geográficas de prueba automatizada",   # 1000
        "Otras restricciones": "Otras restricciones de prueba automatizada",               # 1000, opcional
    },
}


# --- Paso 2 ------------------------------------------------------------------

def completar_paso_2(page, valores: dict) -> None:
    v = VALORES

    # 1. Datos del plan (lo único que cambia entre los servicios móviles)
    page.get_by_role("textbox", name="Valor del servicio").fill(v["valor_servicio"])
    elegir_opcion(page, "Periodicidad", valores["periodicidad"])
    elegir_varias_opciones(page, "Tipo de acceso", valores["tipo_acceso"])

    completar_resto_movil(page, valores, v)


def completar_resto_movil(page, valores: dict, v: dict) -> None:
    """
    Todo el paso 2 de móvil desde "Componentes del plan" hasta el final.
    Es igual en postpago, control y paquete; cada servicio llena antes sus
    primeros campos y luego llama a esta función.
    valores: los del caso (DIMENSIONES); v: los VALORES fijos del servicio.
    """
    elegir_valor(page, "Componentes del plan", valores["componentes"])

    # 2. Internet y tope de datos primero: tethering y traslada dependen del tope.
    elegir_opcion(page, "Internet Ilimitado", valores["internet_ilimitado"])
    page.get_by_role("spinbutton", name="Datos tope descarga").fill(v["datos_tope_descarga_gb"])

    page.get_by_role("spinbutton", name="Datos para tethering").fill(v["datos_tethering_gb"])
    page.get_by_role("spinbutton", name="Traslada datos").fill(v["traslada_datos_gb"])

    # 3. Bolsas de gigas
    for seccion, filas in v["bolsas"].items():
        for fila in filas:
            agregar_bolsa_datos(page, seccion, fila)

    # 4. Apps sin consumo de datos
    if v["apps_sin_consumo"]:
        elegir_varias_opciones(page, "Apps sin consumo de datos", v["apps_sin_consumo"])

    # 5. Tecnologías y velocidades
    for fila in v["velocidades"]:
        agregar_velocidad(page, "Tecnologías y velocidades", fila)

    # 6. Velocidades con degradación
    for etiqueta, valor in v["degradacion"].items():
        page.get_by_role("textbox", name=etiqueta).fill(valor)

    # 7. Minutos / SMS y tarifas
    for etiqueta, config in v["condicionales"].items():
        completar_condicional_triple(page, etiqueta, config)

    page.get_by_role("textbox", name="Tarifa por MB adicional").fill(v["tarifa_mb_adicional"])
    page.get_by_role("textbox", name="Tarifa minuto móvil a fijo").fill(v["tarifa_movil_a_fijo"])
    page.get_by_role("textbox", name="Tarifa minuto móvil todo destino").fill(v["tarifa_movil_todo_destino"])

    # 8. Tablas por países y datos roaming
    for seccion, filas in v["tablas_paises"].items():
        for fila in filas:
            agregar_tarifa_paises(page, seccion, fila)

    if v["datos_roaming_incluidos_gb"]:
        page.get_by_role("spinbutton", name="Datos Roaming incluidos en").fill(v["datos_roaming_incluidos_gb"])
    if v["datos_roaming_adicionales_gb"]:
        page.get_by_role("spinbutton", name="Datos Roaming adicionales a").fill(v["datos_roaming_adicionales_gb"])

    # 9. Condiciones y restricciones
    for etiqueta, texto in v["textos"].items():
        if texto:
            page.get_by_role("textbox", name=etiqueta).fill(texto)
