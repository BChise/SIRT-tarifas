import csv
from pathlib import Path

from playwright.sync_api import BrowserContext, Locator, Page, sync_playwright


CDP_URL = "http://localhost:9222"

ARCHIVO_ENTRADA = Path("tarifas_ejecucion_511.csv")
ARCHIVO_RESULTADOS = Path("resultados_renombre.csv")
DIRECTORIO_CAPTURAS = Path("capturas_renombre")

SUFIJO_RENOMBRE = "_REN"
TIMEOUT = 60_000


def obtener_pagina_sirt(context: BrowserContext) -> Page:
    for page in context.pages:
        if "sirt.pe.be1business.com" in page.url:
            return page

    raise RuntimeError("No se encontró una pestaña abierta de SIRT.")


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
            if fila.get("resultado") not in {"EXITOSO", "YA_RENOMBRADA"}:
                continue

            codigo = fila.get("codigo_sirt", "").strip()
            version = fila.get("version", "").strip()

            if codigo and version:
                procesadas.add((codigo, version))

    return procesadas


def registrar_resultado(
    codigo: str,
    version: str,
    nombre_anterior: str,
    nombre_nuevo: str,
    resultado: str,
    detalle: str,
) -> None:
    archivo_nuevo = not ARCHIVO_RESULTADOS.exists()

    campos = [
        "codigo_sirt",
        "version",
        "nombre_anterior",
        "nombre_nuevo",
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
                "version": version,
                "nombre_anterior": nombre_anterior,
                "nombre_nuevo": nombre_nuevo,
                "resultado": resultado,
                "detalle": detalle,
            }
        )


def abrir_tarifas_registradas(page: Page) -> None:
    enlace = page.get_by_role("link", name="Tarifas registradas")

    if enlace.count() > 0:
        enlace.click()

    page.get_by_role("button", name="Filtro avanzado").wait_for(state="visible")


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


def obtener_fila_tarifa(page: Page, codigo: str, version: str) -> tuple[Locator, str]:
    filas = page.locator("tbody tr")
    cantidad_filas = filas.count()

    for indice in range(cantidad_filas):
        fila = filas.nth(indice)
        celdas = fila.locator("td")

        if celdas.count() < 7:
            continue

        codigo_fila = celdas.nth(2).inner_text().strip()
        version_fila = celdas.nth(3).inner_text().strip()
        nombre_fila = celdas.nth(6).inner_text().strip()

        if codigo_fila == codigo and version_fila == version:
            return fila, nombre_fila

    raise RuntimeError(f"No se encontró {codigo}, versión {version}, después de aplicar el filtro.")


def renombrar_tarifa(
    page: Page,
    codigo: str,
    version: str,
) -> tuple[str, str, str]:
    buscar_tarifa(page, codigo)

    fila, nombre_actual = obtener_fila_tarifa(page, codigo, version)

    if nombre_actual.upper().endswith(SUFIJO_RENOMBRE):
        return nombre_actual, nombre_actual, "YA_RENOMBRADA"

    nombre_nuevo = f"{nombre_actual}{SUFIJO_RENOMBRE}"

    fila.get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Renombre").click()

    campo_nombre = page.get_by_role("textbox", name="Ingresa el nombre de la tarifa")
    campo_nombre.wait_for(state="visible")
    campo_nombre.fill(nombre_nuevo)

    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    page.get_by_role("button", name="Aceptar", exact=True).click()

    page.get_by_role("textbox", name="Ingresa código tarifa").wait_for(state="visible")
    page.wait_for_timeout(500)

    return nombre_actual, nombre_nuevo, "EXITOSO"


def limpiar_filtro(page: Page) -> None:
    try:
        campo_codigo = page.get_by_role("textbox", name="Ingresa código tarifa")

        if campo_codigo.count() > 0 and campo_codigo.is_visible():
            campo_codigo.fill("")

            boton_limpiar = page.get_by_role("button", name="Limpiar selección")

            if boton_limpiar.count() > 0 and boton_limpiar.is_visible():
                boton_limpiar.click()

            page.get_by_role("button", name="Buscar", exact=True).click()
            page.wait_for_timeout(500)

    except Exception:
        page.keyboard.press("Escape")


def cerrar_dialogos(page: Page) -> None:
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
    except Exception:
        pass


def ejecutar() -> None:
    DIRECTORIO_CAPTURAS.mkdir(exist_ok=True)

    tarifas = leer_tarifas_objetivo()
    procesadas = leer_tarifas_procesadas()

    exitosas = 0
    ya_renombradas = 0
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

        for correlativo, tarifa in enumerate(tarifas, start=1):
            codigo = tarifa["codigo_sirt"]
            version = tarifa["version"]
            clave = (codigo, version)

            if clave in procesadas:
                omitidas += 1
                continue

            print(f"\n[{correlativo}/{len(tarifas)}] {codigo}, versión {version}")

            try:
                nombre_anterior, nombre_nuevo, resultado = renombrar_tarifa(
                    page,
                    codigo,
                    version,
                )

                registrar_resultado(
                    codigo=codigo,
                    version=version,
                    nombre_anterior=nombre_anterior,
                    nombre_nuevo=nombre_nuevo,
                    resultado=resultado,
                    detalle="",
                )

                procesadas.add(clave)

                if resultado == "YA_RENOMBRADA":
                    ya_renombradas += 1
                    print(f"Ya estaba renombrada: {nombre_anterior}")
                else:
                    exitosas += 1
                    print(f"Renombrada: {nombre_nuevo}")

            except Exception as error:
                errores += 1
                captura = DIRECTORIO_CAPTURAS / f"error_{codigo}_v{version}.png"

                page.screenshot(path=str(captura), full_page=True)

                registrar_resultado(
                    codigo=codigo,
                    version=version,
                    nombre_anterior="",
                    nombre_nuevo="",
                    resultado="ERROR",
                    detalle=str(error).replace("\n", " "),
                )

                print(f"Error: {error}")
                print(f"Captura: {captura}")

                cerrar_dialogos(page)

            finally:
                limpiar_filtro(page)

    pendientes = len(tarifas) - len(procesadas)

    print("\nPROCESO FINALIZADO")
    print("------------------")
    print(f"Renombradas correctamente: {exitosas}")
    print(f"Ya estaban renombradas: {ya_renombradas}")
    print(f"Omitidas por ejecución anterior: {omitidas}")
    print(f"Errores en esta ejecución: {errores}")
    print(f"Pendientes: {pendientes}")
    print(f"Resultados: {ARCHIVO_RESULTADOS.resolve()}")


if __name__ == "__main__":
    ejecutar()