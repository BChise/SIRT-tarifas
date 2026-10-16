import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-PTV-PROMO-003")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizadaafeerrata")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("combobox", name="Tipo de Contratación Postpago").click()
    page.get_by_role("listbox", name="Tipo de Contratación").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Condiciones generales de prueba automatizadafeerrata")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("Restricciones técnicas de prueba automatizadaferrata")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Continuar con el paso").press("F12")

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

