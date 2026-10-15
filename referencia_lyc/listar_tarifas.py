import csv
import re
import time
from pathlib import Path

from playwright.sync_api import BrowserContext, Page, sync_playwright


CDP_URL = "http://localhost:9222"
ARCHIVO_SALIDA = Path("listado_tarifas_lyc.csv")
ARCHIVO_REPETIDOS = Path("tarifas_repetidas.csv")

PREFIJO_NOMBRE = "LyC-0"
TAMANO_PAGINA = 100
TIMEOUT_SEGUNDOS = 60

INDICE_FECHA = 1
INDICE_CODIGO = 2
INDICE_VERSION = 3
INDICE_TIPO = 4
INDICE_CONDICION = 5
INDICE_NOMBRE = 6
INDICE_ALCANCE = 7
INDICE_USUARIO = 8
INDICE_ESTADO = 9
INDICE_PRIORIDAD = 10


def obtener_pagina_sirt(context: BrowserContext) -> Page:
    for page in context.pages:
        if "sirt.pe.be1business.com" in page.url:
            return page

    raise RuntimeError("No se encontró una pestaña abierta de SIRT.")


def obtener_texto_paginador(page: Page) -> str:
    paginador = page.locator(".mat-mdc-paginator-range-label")
    paginador.wait_for(state="visible")
    return paginador.inner_text().strip()


def interpretar_paginador(texto: str) -> tuple[int, int, int]:
    numeros = [int(numero) for numero in re.findall(r"\d+", texto)]

    if len(numeros) < 3:
        raise ValueError(f"No se pudo interpretar el paginador: {texto}")

    return numeros[-3], numeros[-2], numeros[-1]


def obtener_filas_dom(page: Page) -> list[dict[str, object]]:
    return page.locator("tbody tr").evaluate_all(
        """
        filas => filas.map((fila, indice) => {
            const celdas = Array.from(fila.querySelectorAll("td"))
                .map(celda => celda.innerText.trim());

            return {
                numero_fila: indice + 1,
                celdas: celdas
            };
        })
        """
    )


def convertir_filas(page: Page) -> list[dict[str, str]]:
    filas_dom = obtener_filas_dom(page)
    tarifas = []

    for fila in filas_dom:
        celdas = fila["celdas"]

        if len(celdas) <= INDICE_PRIORIDAD:
            continue

        tarifas.append(
            {
                "numero_fila": str(fila["numero_fila"]),
                "fecha_registro": str(celdas[INDICE_FECHA]).strip(),
                "codigo_sirt": str(celdas[INDICE_CODIGO]).strip(),
                "version": str(celdas[INDICE_VERSION]).strip(),
                "tipo_tarifa": str(celdas[INDICE_TIPO]).strip(),
                "condicion": str(celdas[INDICE_CONDICION]).strip(),
                "nombre": str(celdas[INDICE_NOMBRE]).strip(),
                "alcance": str(celdas[INDICE_ALCANCE]).strip(),
                "usuario_registro": str(celdas[INDICE_USUARIO]).strip(),
                "estado_comercializacion": str(celdas[INDICE_ESTADO]).strip(),
                "prioridad": str(celdas[INDICE_PRIORIDAD]).strip(),
            }
        )

    return tarifas


def obtener_claves_visibles(page: Page) -> set[tuple[str, str]]:
    return {
        (tarifa["codigo_sirt"], tarifa["version"])
        for tarifa in convertir_filas(page)
        if tarifa["codigo_sirt"] and tarifa["version"]
    }


def esperar_cambio_paginador(page: Page, texto_anterior: str) -> str:
    page.wait_for_function(
        """
        textoAnterior => {
            const paginador = document.querySelector(".mat-mdc-paginator-range-label");

            return paginador &&
                paginador.textContent.trim() !== textoAnterior;
        }
        """,
        arg=texto_anterior,
        timeout=TIMEOUT_SEGUNDOS * 1000,
    )

    return obtener_texto_paginador(page)


def esperar_nueva_pagina(
    page: Page,
    claves_anteriores: set[tuple[str, str]],
    cantidad_esperada: int,
) -> None:
    tiempo_limite = time.monotonic() + TIMEOUT_SEGUNDOS
    ultima_secuencia = None
    instante_estable = None

    while time.monotonic() < tiempo_limite:
        tarifas_actuales = convertir_filas(page)

        secuencia_actual = tuple(
            (tarifa["codigo_sirt"], tarifa["version"])
            for tarifa in tarifas_actuales
            if tarifa["codigo_sirt"] and tarifa["version"]
        )

        claves_actuales = set(secuencia_actual)

        cantidad_correcta = len(tarifas_actuales) == cantidad_esperada
        pagina_diferente = claves_actuales != claves_anteriores
        contenido_valido = len(claves_actuales) > 0

        if cantidad_correcta and pagina_diferente and contenido_valido:
            if secuencia_actual == ultima_secuencia:
                if instante_estable is None:
                    instante_estable = time.monotonic()

                if time.monotonic() - instante_estable >= 3:
                    return
            else:
                ultima_secuencia = secuencia_actual
                instante_estable = time.monotonic()
        else:
            ultima_secuencia = None
            instante_estable = None

        page.wait_for_timeout(500)

    raise TimeoutError(
        f"La tabla no terminó de actualizarse. "
        f"Filas esperadas: {cantidad_esperada}."
    )


def esperar_primera_carga(page: Page, cantidad_esperada: int) -> None:
    tiempo_limite = time.monotonic() + TIMEOUT_SEGUNDOS
    claves_estables = None
    lecturas_estables = 0

    while time.monotonic() < tiempo_limite:
        tarifas_actuales = convertir_filas(page)
        claves_actuales = obtener_claves_visibles(page)

        cantidad_correcta = len(tarifas_actuales) == cantidad_esperada
        contenido_valido = len(claves_actuales) > 0

        if cantidad_correcta and contenido_valido:
            if claves_actuales == claves_estables:
                lecturas_estables += 1
            else:
                claves_estables = claves_actuales
                lecturas_estables = 1

            if lecturas_estables >= 2:
                return
        else:
            claves_estables = None
            lecturas_estables = 0

        page.wait_for_timeout(500)

    raise TimeoutError("La primera página no terminó de cargar correctamente.")


def configurar_elementos_por_pagina(page: Page) -> None:
    texto_actual = obtener_texto_paginador(page)
    inicio, fin, total = interpretar_paginador(texto_actual)

    if inicio == 1 and fin == min(TAMANO_PAGINA, total):
        esperar_primera_carga(page, fin - inicio + 1)
        print(f"Paginador ya configurado: {texto_actual}")
        return

    texto_anterior = texto_actual

    page.locator("mat-paginator mat-select").click()

    opcion = page.locator(".cdk-overlay-container").get_by_role(
        "option",
        name=str(TAMANO_PAGINA),
        exact=True,
    )

    opcion.wait_for(state="visible")
    opcion.click(force=True)

    texto_nuevo = esperar_cambio_paginador(page, texto_anterior)
    inicio, fin, total = interpretar_paginador(texto_nuevo)
    cantidad_esperada = min(fin, total) - inicio + 1

    esperar_primera_carga(page, cantidad_esperada)

    print(f"Paginador configurado: {texto_nuevo}")


def volver_primera_pagina(page: Page) -> None:
    texto_actual = obtener_texto_paginador(page)
    inicio, _, _ = interpretar_paginador(texto_actual)

    if inicio == 1:
        return

    boton_primera = page.get_by_role("button", name="Primera página")

    if boton_primera.count() == 0:
        raise RuntimeError("No se encontró el botón Primera página.")

    texto_anterior = texto_actual
    boton_primera.click()

    texto_nuevo = esperar_cambio_paginador(page, texto_anterior)
    inicio, fin, total = interpretar_paginador(texto_nuevo)
    cantidad_esperada = min(fin, total) - inicio + 1

    esperar_primera_carga(page, cantidad_esperada)


def avanzar_pagina(page: Page) -> None:
    texto_anterior = obtener_texto_paginador(page)
    claves_anteriores = obtener_claves_visibles(page)

    boton_siguiente = page.get_by_role("button", name="Página siguiente")
    boton_siguiente.click()

    texto_nuevo = esperar_cambio_paginador(page, texto_anterior)
    inicio, fin, total = interpretar_paginador(texto_nuevo)
    cantidad_esperada = min(fin, total) - inicio + 1

    esperar_nueva_pagina(page, claves_anteriores, cantidad_esperada)


def extraer_tarifas_pagina(page: Page, numero_pagina: int) -> list[dict[str, str]]:
    tarifas = []

    for tarifa in convertir_filas(page):
        if not tarifa["nombre"].lower().startswith(PREFIJO_NOMBRE.lower()):
            continue

        if not tarifa["codigo_sirt"] or not tarifa["version"]:
            continue

        tarifa["pagina"] = str(numero_pagina)
        tarifas.append(tarifa)

    return tarifas


def separar_tarifas_unicas(
    tarifas: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    tarifas_unicas = []
    tarifas_repetidas = []
    claves_registradas = set()

    for tarifa in tarifas:
        clave = (tarifa["codigo_sirt"], tarifa["version"])

        if clave in claves_registradas:
            tarifas_repetidas.append(tarifa)
            continue

        claves_registradas.add(clave)
        tarifas_unicas.append(tarifa)

    return tarifas_unicas, tarifas_repetidas


def guardar_csv(ruta: Path, tarifas: list[dict[str, str]]) -> None:
    campos = [
        "pagina",
        "numero_fila",
        "fecha_registro",
        "codigo_sirt",
        "version",
        "tipo_tarifa",
        "condicion",
        "nombre",
        "alcance",
        "usuario_registro",
        "estado_comercializacion",
        "prioridad",
    ]

    with ruta.open("w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=campos, delimiter=";")
        escritor.writeheader()
        escritor.writerows(tarifas)


def ejecutar() -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = obtener_pagina_sirt(context)

        page.bring_to_front()

        enlace_tarifas = page.get_by_role("link", name="Tarifas registradas")

        if enlace_tarifas.count() > 0:
            enlace_tarifas.click()

        page.locator(".mat-mdc-paginator-range-label").wait_for(state="visible")
        page.locator("tbody tr").first.wait_for(state="visible")

        configurar_elementos_por_pagina(page)
        volver_primera_pagina(page)
        ordenar_por_codigo_sirt(page)

        tarifas_totales = []
        numero_pagina = 1

        while True:
            texto_paginador = obtener_texto_paginador(page)
            inicio, fin, total = interpretar_paginador(texto_paginador)

            print(f"\nLeyendo página {numero_pagina}: {texto_paginador}")

            tarifas_pagina = extraer_tarifas_pagina(page, numero_pagina)
            tarifas_totales.extend(tarifas_pagina)

            print(f"Tarifas {PREFIJO_NOMBRE} encontradas: {len(tarifas_pagina)}")
            print(f"Total acumulado: {len(tarifas_totales)}")

            if fin >= total:
                print("\nSe llegó a la última página.")
                break

            avanzar_pagina(page)
            numero_pagina += 1

        tarifas_unicas, tarifas_repetidas = separar_tarifas_unicas(tarifas_totales)

        guardar_csv(ARCHIVO_SALIDA, tarifas_unicas)

        if tarifas_repetidas:
            guardar_csv(ARCHIVO_REPETIDOS, tarifas_repetidas)

        print("\nProceso finalizado.")
        print(f"Páginas recorridas: {numero_pagina}")
        print(f"Tarifas leídas: {len(tarifas_totales)}")
        print(f"Claves Código SIRT + Versión únicas: {len(tarifas_unicas)}")
        print(f"Lecturas repetidas: {len(tarifas_repetidas)}")
        print(f"Archivo generado: {ARCHIVO_SALIDA.resolve()}")

        if tarifas_repetidas:
            print(f"Repetidos guardados en: {ARCHIVO_REPETIDOS.resolve()}")

def ordenar_por_codigo_sirt(page: Page) -> None:
    encabezado = page.get_by_role("columnheader", name=re.compile(r"Código SIRT", re.IGNORECASE))
    encabezado.click()
    page.wait_for_timeout(3_000)

if __name__ == "__main__":
    ejecutar()