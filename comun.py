"""
Funciones comunes a todos los servicios de "Tarifa establecida / Monoproducto".

Aquí vive todo lo que es igual para cualquier servicio:
- conexión al Edge controlable y navegación,
- listas desplegables, fechas y campos opcionales,
- paso 1 completo,
- bloques reutilizables del paso 2 (tabla de velocidades, minutos, LDI,
  carga de archivos),
- confirmación,
- combinaciones, plan, resultados y reanudación.

Si una corrección sirve para todos los servicios, se hace aquí una sola vez.
"""

import csv
import re
from datetime import date
from itertools import combinations, product
from pathlib import Path


# =============================================================================
# CONFIGURACIÓN GENERAL
# =============================================================================

CDP_URL = "http://localhost:9222"
DOMINIO_SIRT = "sirt"
TIMEOUT = 30_000

TIPO_TARIFA = "Tarifa establecida"
CONDICION = "Monoproducto"

# Valores por defecto del paso 1. Un servicio puede cambiarlos con su propio PASO1.
PASO1_POR_DEFECTO = {
    "tarifa_instalacion": "10",          # solo si el campo está habilitado
    "otros_datos": "Otros datos de prueba automatizada",
    "fecha_inicio_vigencia": None,       # None = hoy
    "fecha_inicio_comercializacion": None,
}

DIRECTORIO_RESULTADOS = Path("resultados")

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


# =============================================================================
# CONEXIÓN Y NAVEGACIÓN
# =============================================================================

def obtener_pagina_sirt(context):
    for page in context.pages:
        if DOMINIO_SIRT in page.url:
            return page

    raise RuntimeError("No se encontró una pestaña abierta de SIRT.")


def abrir_nueva_tarifa(page) -> None:
    page.get_by_role("button", name="Nueva Tarifa").click()
    page.get_by_role("combobox", name="Tipo de Tarifa").wait_for(state="visible")


def volver_al_listado(page, url_listado: str) -> None:
    """Sale de donde esté y vuelve a 'Tarifas registradas'."""
    page.keyboard.press("Escape")
    page.goto(url_listado, wait_until="domcontentloaded")
    page.get_by_role("button", name="Nueva Tarifa").wait_for(state="visible")


# =============================================================================
# LISTAS DESPLEGABLES Y CAMPOS
# =============================================================================

def opcion_en_overlay(page, texto: str):
    return page.locator(".cdk-overlay-container").get_by_role(
        "option", name=texto, exact=True
    )


def elegir_opcion(page, nombre_lista: str, opcion: str) -> None:
    """Lista de selección única."""
    page.get_by_role("combobox", name=nombre_lista).click()
    opcion_en_overlay(page, opcion).click()


def elegir_varias_opciones(page, nombre_lista: str, opciones) -> None:
    """Lista de selección múltiple (se cierra con Escape)."""
    page.get_by_role("combobox", name=nombre_lista).click()

    for opcion in opciones:
        opcion_en_overlay(page, opcion).click()

    page.keyboard.press("Escape")


def elegir_valor(page, nombre_lista: str, valor) -> None:
    """Elige según el valor del caso: texto = selección única, tupla = múltiple."""
    if isinstance(valor, tuple):
        elegir_varias_opciones(page, nombre_lista, valor)
    else:
        elegir_opcion(page, nombre_lista, valor)


def leer_opciones(page, nombre_lista: str) -> list[str]:
    page.get_by_role("combobox", name=nombre_lista).click()

    opciones = page.locator(".cdk-overlay-container").get_by_role("option")
    opciones.first.wait_for(state="visible")
    textos = [texto.strip() for texto in opciones.all_inner_texts() if texto.strip()]

    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    return textos


def llenar_si_editable(campo, valor: str) -> bool:
    """Llena el campo solo si existe y está habilitado. Devuelve si lo llenó."""
    if not valor or campo.count() == 0 or not campo.is_editable():
        return False

    campo.fill(valor)
    return True


def elegir_fecha(page, etiqueta: str, fecha: date) -> None:
    texto_boton = f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"

    page.locator("mat-form-field").filter(has_text=etiqueta).get_by_label(
        "Open calendar"
    ).click()

    boton = page.get_by_role("button", name=texto_boton, exact=True)

    if boton.count() > 0:
        boton.click()
        return

    # La fecha no está en el mes que muestra el calendario: se escribe directo.
    page.keyboard.press("Escape")
    campo = page.get_by_role("textbox", name=etiqueta)
    campo.fill(fecha.strftime("%d/%m/%Y"))
    campo.press("Tab")


# =============================================================================
# PASO 1 (igual para todos los servicios)
# =============================================================================

def completar_paso_1(page, nombre: str, servicio: str, alcances, paso1: dict) -> bool:
    """Llena el paso 1. Devuelve True si 'Tarifa Instalación' estaba habilitada."""
    hoy = date.today()

    elegir_opcion(page, "Tipo de Tarifa", TIPO_TARIFA)
    elegir_opcion(page, "Condición", CONDICION)
    elegir_opcion(page, "Seleccionar servicio", servicio)

    page.get_by_role("textbox", name="Nombre de la tarifa").fill(nombre)

    elegir_varias_opciones(page, "Alcance de la tarifa", alcances)

    elegir_fecha(page, "Inicio de vigencia", paso1["fecha_inicio_vigencia"] or hoy)
    elegir_fecha(page, "Inicio comercialización", paso1["fecha_inicio_comercializacion"] or hoy)

    instalacion = llenar_si_editable(
        page.get_by_role("spinbutton", name="Tarifa Instalación"),
        paso1["tarifa_instalacion"],
    )

    page.get_by_role("textbox", name="Otros datos generales").fill(paso1["otros_datos"])

    return instalacion


def guardar_paso_1(page) -> str:
    """Guarda el paso 1 y devuelve el número de registro (borrador) creado."""
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.wait_for_url(re.compile(r"step=2"))

    coincidencia = re.search(r"nueva-tarifa/(\d+)", page.url)
    return coincidencia.group(1) if coincidencia else ""


# =============================================================================
# BLOQUES REUTILIZABLES DEL PASO 2
# =============================================================================

def boton_agregar(page, referencia):
    """
    Botón "Agregar" de una tabla.

    referencia:
    - int: posición del botón en la página (0 = el primero). Sirve cuando hay
      pocas tablas (Internet fijo, Telefonía fija).
    - str: título exacto de la sección (por ejemplo, "Tarifa SMS roaming").
      Preferible cuando hay muchas tablas (Móvil): no depende del orden.
    """
    if isinstance(referencia, int):
        return page.get_by_role("button", name="Agregar").nth(referencia)

    # El "Agregar" de cada sección es el primer botón "Agregar" que SIGUE al
    # título en la página (así aparece en el mapa de Móvil, en las 15 tablas).
    # No se busca por "contenedor de la sección": en algunas secciones el
    # contenedor más cercano abarca otras tablas y se pulsaba el botón de otra.
    titulo = f"//*[normalize-space(text())='{referencia}']"

    if page.locator(f"xpath={titulo}").count() == 0:
        raise RuntimeError(f"No se encontró la sección '{referencia}'.")

    return page.locator(
        f"xpath=({titulo}/following::button[contains(normalize-space(.), 'Agregar')])[1]"
    )


def agregar_velocidad(page, indice_boton, fila: dict) -> None:
    """
    Tabla de velocidades (ventana emergente con Tecnología + 4 velocidades).
    indice_boton: posición del botón "Agregar" (int) o título de la sección (str).

    fila = {"tecnologia", "descarga_contratada", "descarga_garantizada",
            "carga_contratada", "carga_garantizada"}
    """
    boton_agregar(page, indice_boton).click()

    # La ventana no está marcada como "dialog": se detecta por su primer campo.
    # Mientras está abierta, el resto del formulario queda oculto para estas
    # búsquedas, por eso los campos se buscan en toda la página.
    descarga_contratada = page.get_by_role("spinbutton", name="Velocidad de descarga contratada")
    descarga_contratada.wait_for(state="visible")

    # "Tecnología" de la ventana, sin confundirla con "Tecnología (XX)" del formulario.
    page.get_by_role("combobox", name=re.compile(r"^Tecnología(?!\s*\()")).click()
    opcion_en_overlay(page, fila["tecnologia"]).click()

    descarga_contratada.fill(fila["descarga_contratada"])
    page.get_by_role("spinbutton", name="Velocidad de descarga garantizada").fill(fila["descarga_garantizada"])
    page.get_by_role("spinbutton", name="Velocidad de carga contratada").fill(fila["carga_contratada"])
    page.get_by_role("spinbutton", name="Velocidad de carga garantizada").fill(fila["carga_garantizada"])

    page.get_by_role("button", name="Guardar", exact=True).click()
    descarga_contratada.wait_for(state="hidden")


def campos_condicional(page, etiqueta: str):
    """
    Campos de un "condicional triple" (por ejemplo, "Minutos todo destino nacional"):
    título + casilla "Es ilimitado" + cantidad + tarifa adicional.

    Devuelve (casilla, cantidad, tarifa): la casilla y los dos campos que
    SIGUEN al título en la página. No depende de cómo esté anidada la fila
    (en Móvil, "Minutos móviles off net" está armada distinto y la búsqueda
    por "contenedor de la fila" tomaba la sección entera, con 9 casillas).
    """
    titulo = f"//*[normalize-space(text())='{etiqueta}']"

    if page.locator(f"xpath={titulo}").count() == 0:
        raise RuntimeError(f"No se encontró la fila '{etiqueta}'.")

    casilla = page.locator(f"xpath=({titulo}/following::input[@type='checkbox'])[1]")
    cantidad = page.locator(f"xpath=({titulo}/following::input[not(@type='checkbox')])[1]")
    tarifa = page.locator(f"xpath=({titulo}/following::input[not(@type='checkbox')])[2]")

    return casilla, cantidad, tarifa


def completar_condicional_triple(page, etiqueta: str, config: dict) -> None:
    """
    config = {"modo": "omitir"}                               → no se toca (campo opcional)
    config = {"modo": "ilimitado"}                            → marca "Es ilimitado"
    config = {"modo": "limitado", "minutos": "100", "tarifa": "0.10"}
                                                              → llena los dos campos
    """
    modo = config.get("modo", "omitir")

    if modo == "omitir":
        return

    casilla, cantidad, tarifa = campos_condicional(page, etiqueta)

    if modo == "ilimitado":
        casilla.check()
        return

    if modo == "limitado":
        if casilla.is_checked():
            casilla.uncheck()

        cantidad.fill(config["minutos"])
        tarifa.fill(config["tarifa"])
        return

    raise ValueError(f"Modo desconocido para '{etiqueta}': {modo}")


def agregar_tarifa_paises(page, indice_boton, fila: dict) -> None:
    """
    Tabla de tarifas por países (ventana con Tipo + Tarifa + Países),
    por ejemplo "Tarifa por minuto a fijo LDI".
    indice_boton: posición del botón "Agregar" (int) o título de la sección (str).

    fila = {"tipo": "texto exacto" o None (= primera opción),
            "tarifa": "1.50",
            "paises": ["ALEMANIA", ...]}
    """
    boton_agregar(page, indice_boton).click()

    # Igual que la ventana de velocidades: se detecta por uno de sus campos.
    campo_tarifa = page.get_by_role("spinbutton", name="Tarifa", exact=True)
    campo_tarifa.wait_for(state="visible")

    # Con la ventana abierta, el resto del formulario queda oculto para esta
    # búsqueda, así que "Tipo" no se confunde con "Tipo de acceso".
    page.get_by_role("combobox", name="Tipo").click()

    if fila.get("tipo"):
        opcion_en_overlay(page, fila["tipo"]).click()
    else:
        page.locator(".cdk-overlay-container").get_by_role("option").first.click()

    campo_tarifa.fill(fila["tarifa"])

    # "Países" admite escribir: se escribe el nombre para filtrar la lista larga
    # y luego se elige la opción. En la grabación, la lista se cierra después
    # de cada elección.
    for pais in fila["paises"]:
        campo_paises = page.get_by_role("combobox", name="Países")
        campo_paises.click()

        if campo_paises.is_editable():
            campo_paises.fill(pais)

        opcion_en_overlay(page, pais).click()

    page.get_by_role("button", name="Guardar", exact=True).click()
    campo_tarifa.wait_for(state="hidden")


def agregar_bolsa_datos(page, seccion: str, fila: dict) -> None:
    """
    Tabla "Bolsa de gigas para plataformas" (Móvil: "Bolsa de datos para apps",
    "Datos para Apps 1"...): ventana con Tipo + Gigas + Plataformas (múltiple).
    seccion: título de la sección (str) o posición del botón "Agregar" (int).

    fila = {"tipo": "Bolsa" o None (= primera opción),
            "gigas": "10",
            "plataformas": ["Whatsapp", "Messenger", ...]}
    """
    boton_agregar(page, seccion).click()

    # Como las otras ventanas: se detecta por uno de sus campos.
    campo_gigas = page.get_by_role("spinbutton", name="Gigas")
    campo_gigas.wait_for(state="visible")

    page.get_by_role("combobox", name="Tipo").click()
    if fila.get("tipo"):
        opcion_en_overlay(page, fila["tipo"]).click()
    else:
        page.locator(".cdk-overlay-container").get_by_role("option").first.click()

    campo_gigas.fill(fila["gigas"])

    # Plataformas es de selección múltiple. Se cierra la lista con su fondo
    # transparente (como en la grabación) para no cerrar la ventana con Escape.
    page.get_by_role("combobox", name="Plataformas").click()
    for plataforma in fila["plataformas"]:
        opcion_en_overlay(page, plataforma).click()

    fondo = page.locator(".cdk-overlay-transparent-backdrop")
    if fondo.count() > 0:
        fondo.last.click()
    else:
        page.keyboard.press("Escape")

    page.get_by_role("button", name="Guardar", exact=True).click()
    campo_gigas.wait_for(state="hidden")


def fila_con_carga(page, etiqueta: str):
    """
    Fila de un campo "Carga y descarga" (por ejemplo, "Canales analógicos"):
    título + casilla calculada + botón Descargar + botón Cargar.

    Busca el título por su texto exacto y sube hasta el primer contenedor que
    tenga un campo de archivo o, si aún no existe, al menos dos botones.
    """
    fila = page.locator(
        f"xpath=//*[normalize-space(text())='{etiqueta}']"
        "/ancestor::*[.//input[@type='file'] or count(.//button) >= 2][1]"
    )

    if fila.count() == 0:
        raise RuntimeError(f"No se encontró la fila de carga '{etiqueta}'.")

    return fila.first


def cargar_archivo(page, etiqueta: str, ruta: Path) -> None:
    """
    Adjunta un archivo en un campo "Carga y descarga" y espera a que SIRT
    muestre su nombre en la fila (señal de que la carga terminó).

    No se usa el selector del grabador (input[name="nameInputRamdom..."]):
    cambia en cada carga de la página.
    """
    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo a cargar: {ruta.resolve()}")

    fila = fila_con_carga(page, etiqueta)
    campo_archivo = fila.locator("input[type='file']")

    if campo_archivo.count() > 0:
        # El campo de archivo suele estar oculto; Playwright lo llena igual,
        # sin abrir la ventana del sistema para elegir archivos.
        campo_archivo.first.set_input_files(str(ruta))
    else:
        # Respaldo: pulsar "Cargar" (2.º botón de la fila) y entregar el archivo
        # a la ventana de selección que abriría el navegador.
        with page.expect_file_chooser() as seleccion:
            fila.locator("button").nth(1).click()
        seleccion.value.set_files(str(ruta))

    fila.get_by_text(ruta.name, exact=True).wait_for(state="visible")


def completar_textos_finales(page, condiciones: str, restricciones: str) -> None:
    """Características adicionales y restricciones (comunes a los servicios vistos)."""
    page.get_by_role("textbox", name="Condiciones generales").fill(condiciones)
    page.get_by_role("textbox", name="Restricciones técnicas").fill(restricciones)


def guardar_paso_2(page) -> None:
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa", exact=True).wait_for(state="visible")


# =============================================================================
# CONFIRMACIÓN
# =============================================================================

def confirmar_tarifa(page) -> None:
    page.get_by_role("button", name="Confirmar Tarifa", exact=True).click()

    boton_confirmar = page.get_by_role("button", name="Confirmar tarifa", exact=True)
    boton_confirmar.wait_for(state="visible")
    boton_confirmar.click()

    boton_aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    boton_aceptar.wait_for(state="visible")
    boton_aceptar.click()
    boton_aceptar.wait_for(state="hidden")


# =============================================================================
# COMBINACIONES
# =============================================================================

def variantes(dimension: dict) -> list:
    """
    dimension = {"opciones": [...], "multiple": bool, "combinar": bool}

    - Selección única: cada opción es una variante.
    - Selección múltiple con combinar=True: todas las combinaciones no vacías.
    - Selección múltiple con combinar=False: todas las opciones juntas (1 variante).
    """
    opciones = dimension["opciones"]

    if not dimension.get("multiple"):
        return list(opciones)

    if not dimension.get("combinar", True):
        return [tuple(opciones)]

    resultado = []

    for cantidad in range(1, len(opciones) + 1):
        resultado.extend(combinations(opciones, cantidad))

    return resultado


def texto_valor(valor) -> str:
    return " + ".join(valor) if isinstance(valor, tuple) else str(valor)


def generar_casos(dimensiones: dict, codigo: str, lote: str) -> list[dict]:
    nombres = list(dimensiones)
    listas = [variantes(dimensiones[nombre]) for nombre in nombres]

    casos = []

    for correlativo, valores in enumerate(product(*listas), start=1):
        caso = {
            "correlativo": correlativo,
            "nombre": f"TE-{codigo}-{lote}-{correlativo:03d}",
            "valores": dict(zip(nombres, valores)),
        }
        caso["clave"] = " | ".join(texto_valor(valor) for valor in valores)
        casos.append(caso)

    return casos


def mostrar_calculo(dimensiones: dict, casos: list[dict]) -> None:
    print("CÁLCULO DE CASOS")
    print("----------------")

    cantidades = []

    for nombre, dimension in dimensiones.items():
        cantidad = len(variantes(dimension))
        cantidades.append(str(cantidad))
        tipo = "múltiple" if dimension.get("multiple") else "única"
        print(f"{nombre:<22} {cantidad:>3} variantes ({tipo})")

    print(f"Total: {' x '.join(cantidades)} = {len(casos)} casos")


# =============================================================================
# PLAN, RESULTADOS Y REANUDACIÓN
# =============================================================================

CAMPOS_RESULTADO = [
    "correlativo", "nombre", "clave", "id_registro",
    "resultado", "detalle", "captura", "fecha_hora",
]


def rutas_lote(codigo: str, lote: str) -> dict[str, Path]:
    base = DIRECTORIO_RESULTADOS / codigo
    return {
        "base": base,
        "plan": base / f"plan_{lote}.csv",
        "resultados": base / f"resultados_{lote}.csv",
        "capturas": base / f"capturas_{lote}",
    }


def guardar_plan(ruta: Path, dimensiones: dict, casos: list[dict]) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)

    with ruta.open("w", newline="", encoding="utf-8-sig") as salida:
        escritor = csv.writer(salida, delimiter=";")
        escritor.writerow(["correlativo", "nombre", *dimensiones.keys()])

        for caso in casos:
            escritor.writerow(
                [caso["correlativo"], caso["nombre"],
                 *(texto_valor(valor) for valor in caso["valores"].values())]
            )


def leer_claves_exitosas(ruta: Path) -> set[str]:
    if not ruta.exists():
        return set()

    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        return {
            fila["clave"]
            for fila in csv.DictReader(entrada, delimiter=";")
            if fila.get("resultado") == "EXITOSO"
        }


def registrar_resultado(ruta: Path, fila: dict) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    nuevo = not ruta.exists()

    with ruta.open("a", newline="", encoding="utf-8-sig") as salida:
        escritor = csv.DictWriter(salida, fieldnames=CAMPOS_RESULTADO, delimiter=";")

        if nuevo:
            escritor.writeheader()

        escritor.writerow(fila)
