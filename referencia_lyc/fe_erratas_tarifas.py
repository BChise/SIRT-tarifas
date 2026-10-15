import csv
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import BrowserContext, Locator, Page, sync_playwright


CDP_URL = "http://localhost:9222"

ARCHIVO_ENTRADA = Path("tarifas_ejecucion_511.csv")
ARCHIVO_RESULTADOS = Path("resultados_fe_erratas.csv")
DIRECTORIO_CAPTURAS = Path("capturas_fe_erratas")

TIMEOUT = 60_000


def obtener_pagina_sirt(context: BrowserContext) -> Page:
    for page in context.pages:
        if "sirt.pe.be1business.com" in page.url:
            return page

    raise RuntimeError("No se encontró una pestaña abierta de SIRT.")


def obtener_cantidad_servicios(nombre: str) -> int:
    coincidencia = re.search(r"LyC-(\d{2})S-", nombre, re.IGNORECASE)

    if not coincidencia:
        raise ValueError(f"No se pudo obtener la cantidad de servicios: {nombre}")

    return int(coincidencia.group(1))


def leer_tarifas_objetivo() -> list[dict[str, str]]:
    if not ARCHIVO_ENTRADA.exists():
        raise FileNotFoundError(f"No se encontró: {ARCHIVO_ENTRADA.resolve()}")

    with ARCHIVO_ENTRADA.open("r", newline="", encoding="utf-8-sig") as archivo:
        filas = list(csv.DictReader(archivo, delimiter=";"))

    tarifas = []

    for fila in filas:
        codigo = fila.get("codigo_sirt", "").strip()
        nombre = fila.get("nombre", "").strip()

        if codigo and nombre:
            tarifas.append(
                {
                    "codigo_sirt": codigo,
                    "nombre": nombre,
                }
            )

    return tarifas


def leer_codigos_procesados():
    if not ARCHIVO_RESULTADOS.exists():
        return set()

    procesados = set()

    with ARCHIVO_RESULTADOS.open("r", newline="", encoding="utf-8-sig") as archivo:
        filas = csv.DictReader(archivo, delimiter=";")

        for fila in filas:
            if fila.get("resultado") != "EXITOSO":
                continue

            codigo = fila.get("codigo_sirt", "").strip()

            if codigo:
                procesados.add(codigo)

    return procesados


def registrar_resultado(
    codigo: str,
    version_procesada: str,
    nombre: str,
    cantidad_servicios: int,
    marca: str,
    resultado: str,
    detalle: str,
) -> None:
    archivo_nuevo = not ARCHIVO_RESULTADOS.exists()

    campos = [
        "codigo_sirt",
        "version_procesada",
        "nombre",
        "cantidad_servicios",
        "marca",
        "resultado",
        "detalle",
    ]

    with ARCHIVO_RESULTADOS.open("a", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=campos, delimiter=";")

        if archivo_nuevo:
            escritor.writeheader()

        escritor.writerow(
            {
                "codigo_sirt": codigo,
                "version_procesada": version_procesada,
                "nombre": nombre,
                "cantidad_servicios": cantidad_servicios,
                "marca": marca,
                "resultado": resultado,
                "detalle": detalle,
            }
        )


def abrir_tarifas_registradas(page: Page) -> None:
    campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")

    if campo_codigo.count() > 0 and campo_codigo.is_visible():
        return

    enlace = page.get_by_role("link", name="Tarifas registradas")

    if enlace.count() > 0 and enlace.is_visible():
        enlace.click()

    boton_filtro = page.get_by_role("button", name="Filtro avanzado")

    if boton_filtro.count() > 0:
        boton_filtro.wait_for(state="visible")


def abrir_filtro_avanzado(page: Page) -> None:
    campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")

    if campo_codigo.count() > 0 and campo_codigo.is_visible():
        return

    page.get_by_role("button", name="Filtro avanzado").click()
    campo_codigo.wait_for(state="visible")


def buscar_tarifa(page: Page, codigo: str) -> None:
    abrir_filtro_avanzado(page)

    campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")
    campo_codigo.fill(codigo)

    page.get_by_role("button", name="Buscar", exact=True).click()

    page.wait_for_function(
        """
        codigo => {
            const filas = Array.from(document.querySelectorAll("tbody tr"));

            return filas.some(fila => {
                const celdas = Array.from(fila.querySelectorAll("td"));

                return celdas.some(
                    celda => celda.innerText.trim() === codigo
                );
            });
        }
        """,
        arg=codigo,
        timeout=TIMEOUT,
    )

    page.wait_for_timeout(500)


def convertir_version(version: str) -> tuple[int, ...]:
    numeros = re.findall(r"\d+", version)

    if not numeros:
        return (0,)

    return tuple(int(numero) for numero in numeros)


def obtener_fila_version_mas_alta(
    page: Page,
    codigo: str,
) -> tuple[Locator, str, str]:
    filas = page.locator("tbody tr")
    coincidencias = []

    for indice in range(filas.count()):
        fila = filas.nth(indice)
        celdas = fila.locator("td")

        if celdas.count() < 7:
            continue

        codigo_fila = celdas.nth(2).inner_text().strip()
        version_fila = celdas.nth(3).inner_text().strip()
        nombre_fila = celdas.nth(6).inner_text().strip()

        if codigo_fila == codigo:
            coincidencias.append(
                {
                    "fila": fila,
                    "version": version_fila,
                    "nombre": nombre_fila,
                }
            )

    if not coincidencias:
        raise RuntimeError(
            f"No se encontró el código {codigo} después de aplicar el filtro."
        )

    coincidencias.sort(
        key=lambda registro: convertir_version(registro["version"]),
        reverse=True,
    )

    seleccionada = coincidencias[0]

    return (
        seleccionada["fila"],
        seleccionada["version"],
        seleccionada["nombre"],
    )


def agregar_marca(campo: Locator, marca: str) -> str:
    valor_actual = campo.input_value().strip()
    valor_nuevo = f"{valor_actual}_{marca}"

    campo.fill(valor_nuevo)

    return valor_nuevo


def completar_datos_generales(page: Page, marca: str) -> None:
    campo_descripcion = page.get_by_role("textbox", name="Descripción del Servicio")
    campo_atributos = page.get_by_role("textbox", name="Atributos de la tarifa")
    campo_ambito = page.get_by_role("textbox", name="Ambito geográfico")

    campo_descripcion.wait_for(state="visible")
    campo_atributos.wait_for(state="visible")
    campo_ambito.wait_for(state="visible")

    agregar_marca(campo_descripcion, marca)
    agregar_marca(campo_atributos, marca)
    agregar_marca(campo_ambito, marca)

    campo_ambito.press("Tab")
    page.wait_for_timeout(500)

    boton_continuar = page.locator("button:visible").filter(
        has_text="Continuar con el paso"
    ).first

    boton_continuar.wait_for(state="visible", timeout=TIMEOUT)

    page.wait_for_function(
        """
        boton => !boton.disabled && boton.getAttribute("aria-disabled") !== "true"
        """,
        arg=boton_continuar.element_handle(),
        timeout=TIMEOUT,
    )

    boton_continuar.scroll_into_view_if_needed()
    boton_continuar.click()

    page.get_by_role("textbox", name="Detalle del servicio").wait_for(
        state="visible",
        timeout=TIMEOUT,
    )


def omitir_detalles_servicios(page: Page) -> None:
    campo_detalle = page.get_by_role("textbox", name="Detalle del servicio")
    campo_detalle.wait_for(state="visible", timeout=TIMEOUT)

    campo_detalle.press("Tab")
    page.wait_for_timeout(500)

    boton_continuar = page.locator("button:visible").filter(has_text="Continuar con el paso").last
    boton_continuar.wait_for(state="visible", timeout=TIMEOUT)
    boton_continuar.scroll_into_view_if_needed()
    boton_continuar.click()

    page.get_by_role("button", name="Guardar cambios", exact=True).wait_for(
        state="visible",
        timeout=TIMEOUT,
    )


def guardar_fe_erratas(page: Page) -> None:
    page.get_by_role("button", name="Guardar cambios").click()

    page.get_by_role("button", name="Aceptar").wait_for(state="visible", timeout=TIMEOUT)
    page.get_by_role("button", name="Aceptar").click()

    page.wait_for_timeout(1_500)

    page.get_by_role("button", name="Aceptar").wait_for(state="visible", timeout=TIMEOUT)
    page.get_by_role("button", name="Aceptar").click()

    page.get_by_role("textbox", name="Ingresa código tarifa").wait_for(
        state="visible",
        timeout=TIMEOUT,
    )

    page.wait_for_timeout(500)


def aplicar_fe_erratas(
    page: Page,
    codigo: str,
    marca: str,
) -> tuple[str, str]:
    buscar_tarifa(page, codigo)

    fila, version_actual, nombre_actual = obtener_fila_version_mas_alta(
        page,
        codigo,
    )

    fila.get_by_label("Acciones").click()

    page.get_by_role(
        "menuitem",
        name="Fe de Erratas",
        exact=True,
    ).click()

    page.get_by_role(
        "textbox",
        name="Descripción del Servicio",
    ).wait_for(state="visible")

    completar_datos_generales(page, marca)
    omitir_detalles_servicios(page)
    guardar_fe_erratas(page)

    return version_actual, nombre_actual


def recuperar_pagina(page: Page) -> None:
    try:
        for _ in range(3):
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)

        campo_codigo = page.get_by_role(
            "textbox",
            name="Ingresa código tarifa",
        )

        if campo_codigo.count() > 0 and campo_codigo.is_visible():
            return

        enlace = page.get_by_role(
            "link",
            name="Tarifas registradas",
        )

        if enlace.count() > 0 and enlace.is_visible():
            enlace.click()

        abrir_filtro_avanzado(page)

    except Exception as error:
        print(f"No se pudo recuperar la pantalla: {error}")


def ejecutar() -> None:
    DIRECTORIO_CAPTURAS.mkdir(exist_ok=True)

    tarifas = leer_tarifas_objetivo()
    codigos_procesados = leer_codigos_procesados()

    exitosas = 0
    errores = 0
    omitidas = 0

    print(f"Tarifas objetivo: {len(tarifas)}")
    print(f"Procesadas anteriormente: {len(codigos_procesados)}")

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = obtener_pagina_sirt(context)

        page.set_default_timeout(TIMEOUT)
        page.bring_to_front()

        abrir_tarifas_registradas(page)
        abrir_filtro_avanzado(page)

        for correlativo, tarifa in enumerate(tarifas, start=1):
            codigo = tarifa["codigo_sirt"]
            nombre_original = tarifa["nombre"]

            if codigo in codigos_procesados:
                omitidas += 1
                continue

            cantidad_servicios = obtener_cantidad_servicios(
                nombre_original
            )

            marca = (
                "FE_"
                f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
            )

            print(f"\n[{correlativo}/{len(tarifas)}] {codigo}")
            print(f"Nombre base: {nombre_original}")
            print(f"Servicios: {cantidad_servicios}")
            print(f"Marca: {marca}")

            try:
                version_actual, nombre_actual = aplicar_fe_erratas(
                    page,
                    codigo,
                    marca,
                )

                registrar_resultado(
                    codigo=codigo,
                    version_procesada=version_actual,
                    nombre=nombre_actual,
                    cantidad_servicios=cantidad_servicios,
                    marca=marca,
                    resultado="EXITOSO",
                    detalle="",
                )

                codigos_procesados.add(codigo)
                exitosas += 1

                print(
                    "Resultado: Fe de Erratas exitosa "
                    f"sobre versión {version_actual}"
                )

            except Exception as error:
                errores += 1

                captura = (
                    DIRECTORIO_CAPTURAS
                    / f"error_{codigo}.png"
                )

                page.screenshot(
                    path=str(captura),
                    full_page=True,
                )

                registrar_resultado(
                    codigo=codigo,
                    version_procesada="",
                    nombre=nombre_original,
                    cantidad_servicios=cantidad_servicios,
                    marca=marca,
                    resultado="ERROR",
                    detalle=str(error).replace("\n", " "),
                )

                print("Resultado: error")
                print(f"Detalle: {error}")
                print(f"Captura: {captura}")

                recuperar_pagina(page)

    pendientes = len(tarifas) - len(codigos_procesados)

    print("\nPROCESO FINALIZADO")
    print("------------------")
    print(f"Fe de Erratas exitosas: {exitosas}")
    print(f"Omitidas por ejecución anterior: {omitidas}")
    print(f"Errores en esta ejecución: {errores}")
    print(f"Pendientes: {pendientes}")
    print(f"Resultados: {ARCHIVO_RESULTADOS.resolve()}")


if __name__ == "__main__":
    ejecutar()