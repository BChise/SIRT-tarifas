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
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada ")
    page.get_by_role("textbox", name="Otros datos generales").press("CapsLock")
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada FEERRATAS")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

