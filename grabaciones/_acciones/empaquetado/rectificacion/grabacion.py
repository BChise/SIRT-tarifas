import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-EMP-MUESTRA-006")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Rectificar").click()
    page.get_by_role("textbox", name="Tarifa de instalación").click()
    page.get_by_role("textbox", name="Tarifa de instalación").fill("1000")
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada FEERRATAS rectificacion")
    page.get_by_role("combobox", name="Residencial").click()
    page.locator("#mat-select-13-panel").click()
    page.locator("#mat-select-13-panel").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("spinbutton", name="Valor del servicio").click()
    page.get_by_role("spinbutton", name="Valor del servicio").fill("28")
    page.get_by_role("tab", name="Servicio 2").click()
    page.get_by_role("spinbutton", name="Valor del servicio").click()
    page.get_by_role("spinbutton", name="Valor del servicio").fill("98")
    page.get_by_role("tab", name="Servicio 3").click()
    page.get_by_role("spinbutton", name="Valor del servicio").click()
    page.get_by_role("spinbutton", name="Valor del servicio").fill("48")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

