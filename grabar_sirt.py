"""
Grabador de SIRT: se conecta al Edge del proyecto (puerto 9222) y abre el Inspector de
Playwright. En el Inspector pulsa «Record» y haz el flujo en SIRT: el código aparece en el
Inspector; cópialo a grabaciones/<carpeta>/grabacion.py.

Uso (con Edge abierto en modo controlable y la sesión de SIRT iniciada):
    python grabar_sirt.py

Al terminar, cierra el Inspector (o pulsa Ctrl + C en la consola). No cierra Edge.
"""

import sys

from playwright.sync_api import sync_playwright

CDP_URL = "http://localhost:9222"
DOMINIO_SIRT = "sirt"


def main() -> None:
    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(CDP_URL)
        except Exception as error:
            sys.exit(f"No se pudo conectar con Edge en {CDP_URL}. Ábrelo con el puerto 9222.\n{error}")
        context = browser.contexts[0]
        paginas = [p for p in context.pages if DOMINIO_SIRT in p.url] or context.pages
        if not paginas:
            sys.exit("Edge no tiene ninguna pestaña abierta.")
        page = paginas[0]
        print(f"Conectado a: {page.url}")
        print("En el Inspector pulsa «Record» y haz el flujo en SIRT. Ciérralo al terminar.")
        page.pause()


if __name__ == "__main__":
    main()
