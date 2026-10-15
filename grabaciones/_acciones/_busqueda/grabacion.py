import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-SMO-PROMO-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Renombre").click()
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").click()
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").press("CapsLock")
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").fill("E")
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").press("CapsLock")
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").fill("")
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").press("CapsLock")
    page.get_by_role("textbox", name="Ingresa el nombre de la tarifa").fill("TE-SMO-PROMO-001_RENOMBRE")
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Guardar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

