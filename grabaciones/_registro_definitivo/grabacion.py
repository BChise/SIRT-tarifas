import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("link", name="Tarifas temporales").click()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("cell", name="TE-IF-MUESTRA-").click()
    page.locator("body").press("ControlOrMeta+c")
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-IF-MUESTRA-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("checkbox", name="select row 2026-10-14T11:05:").check()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Ver Disponibilidad Comercial").click()
    page.get_by_role("checkbox", name="Amazonas").check()
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("checkbox", name="select row 2026-10-14T11:05:").check()
    page.get_by_role("button", name="Registro definitivo").click()
    page.get_by_role("button", name="Registrar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

