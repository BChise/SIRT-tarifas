import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TP-IF-PROMO-021")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Promoción de prueba automatizada_fe_errata")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("Otras condiciones de prueba_fe_errata")
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

