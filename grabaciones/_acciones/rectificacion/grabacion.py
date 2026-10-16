import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-OS-MUESTRA-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Rectificar").click()
    page.get_by_role("combobox", name="Residencial").click()
    page.get_by_role("option", name="Comercial").locator("mat-pseudo-checkbox").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("textbox", name="Tarifa de instalación").click()
    page.get_by_role("textbox", name="Tarifa de instalación").click()
    page.get_by_role("textbox", name="Tarifa de instalación").fill("11")
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectificacion")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_text("Tipo de monedaSoles").click()
    page.get_by_role("listbox", name="Tipo de moneda").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Condiciones generales de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectif")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("Otras restricciones de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectif")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

