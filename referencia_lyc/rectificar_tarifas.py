import csv
import re
from pathlib import Path

from playwright.sync_api import BrowserContext, Locator, Page, sync_playwright

from datetime import datetime

CDP_URL = "http://localhost:9222"

ARCHIVO_ENTRADA = Path("tarifas_ejecucion_511.csv")
ARCHIVO_RESULTADOS = Path("resultados_rectificacion.csv")
DIRECTORIO_CAPTURAS = Path("capturas_rectificacion")

SUFIJO_RECTIFICACION = "_RECT"
TIMEOUT = 60_000


def obtener_pagina_sirt(context: BrowserContext) -> Page:
    for page in context.pages:
        if "sirt.pe.be1business.com" in page.url:
            return page

    raise RuntimeError("No se encontró una pestaña abierta de SIRT.")


def obtener_cantidad_servicios(nombre: str) -> int:
    coincidencia = re.search(r"LyC-(\d{2})S-", nombre, re.IGNORECASE)

    if not coincidencia:
        raise ValueError(f"No se pudo obtener la cantidad de servicios del nombre: {nombre}")

    return int(coincidencia.group(1))


def leer_tarifas_objetivo() -> list[dict[str, str]]:
    if not ARCHIVO_ENTRADA.exists():
        raise FileNotFoundError(f"No se encontró: {ARCHIVO_ENTRADA.resolve()}")

    with ARCHIVO_ENTRADA.open("r", newline="", encoding="utf-8-sig") as archivo:
        filas = list(csv.DictReader(archivo, delimiter=";"))

    tarifas = []

    for fila in filas:
        codigo = fila.get("codigo_sirt", "").strip()
        version = fila.get("version", "").strip()
        nombre = fila.get("nombre", "").strip()

        if codigo and version and nombre:
            tarifas.append(
                {
                    "codigo_sirt": codigo,
                    "version": version,
                    "nombre": nombre,
                }
            )

    return tarifas


def leer_tarifas_procesadas() -> set[tuple[str, str]]:
    if not ARCHIVO_RESULTADOS.exists():
        return set()

    procesadas = set()

    with ARCHIVO_RESULTADOS.open("r", newline="", encoding="utf-8-sig") as archivo:
        filas = csv.DictReader(archivo, delimiter=";")

        for fila in filas:
            if fila.get("resultado") != "EXITOSO":
                continue

            codigo = fila.get("codigo_sirt", "").strip()
            version = fila.get("version_original", "").strip()

            if codigo and version:
                procesadas.add((codigo, version))

    return procesadas


def registrar_resultado(
    codigo: str,
    version_original: str,
    nombre: str,
    cantidad_servicios: int,
    resultado: str,
    detalle: str,
) -> None:
    archivo_nuevo = not ARCHIVO_RESULTADOS.exists()

    campos = [
        "codigo_sirt",
        "version_original",
        "nombre",
        "cantidad_servicios",
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
                "version_original": version_original,
                "nombre": nombre,
                "cantidad_servicios": cantidad_servicios,
                "resultado": resultado,
                "detalle": detalle,
            }
        )


def abrir_tarifas_registradas(page: Page) -> None:
    campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")

    if campo_codigo.count() > 0 and campo_codigo.is_visible():
        return

    enlace_tarifas = page.get_by_role("link", name="Tarifas registradas")

    if enlace_tarifas.count() > 0:
        enlace_tarifas.click()

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
                return celdas.some(celda => celda.innerText.trim() === codigo);
            });
        }
        """,
        arg=codigo,
        timeout=TIMEOUT,
    )

    page.wait_for_timeout(500)


def obtener_fila_tarifa(page: Page, codigo: str, version: str) -> Locator:
    filas = page.locator("tbody tr")

    for indice in range(filas.count()):
        fila = filas.nth(indice)
        celdas = fila.locator("td")

        if celdas.count() < 7:
            continue

        codigo_fila = celdas.nth(2).inner_text().strip()
        version_fila = celdas.nth(3).inner_text().strip()

        if codigo_fila == codigo and version_fila == version:
            return fila

    raise RuntimeError(f"No se encontró {codigo}, versión {version}, después de aplicar el filtro.")


def agregar_marca_rectificacion(campo: Locator, marca: str) -> str:
    valor_actual = campo.input_value().strip()
    valor_nuevo = f"{valor_actual}_{marca}"
    campo.fill(valor_nuevo)
    return valor_nuevo


def seleccionar_ambito_publico(page: Page) -> None:
    opcion_publico = page.get_by_role("option", name="Público", exact=True)

    selector_ambito = page.locator("mat-select").filter(has=page.locator('[formcontrolname="ambito"]'))

    if selector_ambito.count() > 0:
        selector_ambito.first.click()
    else:
        page.locator(".mat-mdc-select-arrow > svg").first.click()

    opcion_publico.wait_for(state="visible")
    opcion_publico.click()


def completar_datos_generales(page: Page, marca: str) -> None:
    seleccionar_ambito_publico(page)

    campo_descripcion = page.get_by_role("textbox", name="Descripción del Servicio")
    campo_atributos = page.get_by_role("textbox", name="Atributos de la tarifa")

    campo_descripcion.wait_for(state="visible")
    campo_atributos.wait_for(state="visible")

    agregar_marca_rectificacion(campo_descripcion, marca)
    agregar_marca_rectificacion(campo_atributos, marca)

    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("textbox", name="Detalle del servicio").wait_for(state="visible")


def omitir_detalles_servicios(page: Page) -> None:
    boton_continuar = page.get_by_role("button", name="Continuar con el paso")
    boton_continuar.wait_for(state="visible")
    boton_continuar.click()

    page.get_by_role("button", name="Guardar cambios").wait_for(state="visible")


def guardar_rectificacion(page: Page) -> None:
    boton_guardar = page.get_by_role("button", name="Guardar cambios", exact=True)
    boton_guardar.wait_for(state="visible")
    boton_guardar.click()

    primer_aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    primer_aceptar.wait_for(state="visible", timeout=TIMEOUT)
    primer_aceptar.click()

    page.wait_for_timeout(1_000)

    segundo_aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    segundo_aceptar.wait_for(state="visible", timeout=TIMEOUT)
    segundo_aceptar.click()

    campo_filtro = page.get_by_role("textbox", name="Ingresa código tarifa")
    campo_filtro.wait_for(state="visible", timeout=TIMEOUT)

    page.wait_for_timeout(500)


def rectificar_tarifa(page: Page, codigo: str, version: str) -> None:
    buscar_tarifa(page, codigo)

    fila = obtener_fila_tarifa(page, codigo, version)

    fila.get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Rectificar").click()

    page.get_by_role("textbox", name="Descripción del Servicio").wait_for(state="visible")

    marca = f"RECT_{datetime.now().strftime('%Y%m%d%H%M%S')}"

    completar_datos_generales(page, marca)
    omitir_detalles_servicios(page)
    guardar_rectificacion(page)



def recuperar_pagina(page: Page) -> None:
    try:
        for _ in range(3):
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)

        campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")

        if campo_codigo.count() > 0 and campo_codigo.is_visible():
            return

        enlace_tarifas = page.get_by_role("link", name="Tarifas registradas")

        if enlace_tarifas.count() > 0 and enlace_tarifas.is_visible():
            enlace_tarifas.click()

        abrir_filtro_avanzado(page)

    except Exception as error:
        print(f"No se pudo recuperar la pantalla: {error}")


def ejecutar() -> None:
    DIRECTORIO_CAPTURAS.mkdir(exist_ok=True)

    tarifas = leer_tarifas_objetivo()
    procesadas = leer_tarifas_procesadas()

    exitosas = 0
    errores = 0
    omitidas = 0

    print(f"Tarifas objetivo: {len(tarifas)}")
    print(f"Procesadas anteriormente: {len(procesadas)}")

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
            version = tarifa["version"]
            nombre = tarifa["nombre"]
            clave = (codigo, version)

            if clave in procesadas:
                omitidas += 1
                continue

            cantidad_servicios = obtener_cantidad_servicios(nombre)

            print(f"\n[{correlativo}/{len(tarifas)}] {codigo}, versión {version}")
            print(f"Nombre: {nombre}")
            print(f"Servicios: {cantidad_servicios}")

            try:
                rectificar_tarifa(page, codigo, version)

                registrar_resultado(
                    codigo=codigo,
                    version_original=version,
                    nombre=nombre,
                    cantidad_servicios=cantidad_servicios,
                    resultado="EXITOSO",
                    detalle="",
                )

                procesadas.add(clave)
                exitosas += 1

                print("Resultado: rectificación exitosa")

            except Exception as error:
                errores += 1
                captura = DIRECTORIO_CAPTURAS / f"error_{codigo}_v{version}.png"

                page.screenshot(path=str(captura), full_page=True)

                registrar_resultado(
                    codigo=codigo,
                    version_original=version,
                    nombre=nombre,
                    cantidad_servicios=cantidad_servicios,
                    resultado="ERROR",
                    detalle=str(error).replace("\n", " "),
                )

                print("Resultado: error")
                print(f"Detalle: {error}")
                print(f"Captura: {captura}")

                recuperar_pagina(page)

    pendientes = len(tarifas) - len(procesadas)

    print("\nPROCESO FINALIZADO")
    print("------------------")
    print(f"Rectificadas correctamente: {exitosas}")
    print(f"Omitidas por ejecución anterior: {omitidas}")
    print(f"Errores en esta ejecución: {errores}")
    print(f"Pendientes: {pendientes}")
    print(f"Resultados: {ARCHIVO_RESULTADOS.resolve()}")


if __name__ == "__main__":
    ejecutar()