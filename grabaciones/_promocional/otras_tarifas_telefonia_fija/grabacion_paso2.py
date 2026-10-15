import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Tarifa por minuto").check()
    page.get_by_role("combobox", name="Tipo de Variación Seleccione").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-select-value-0").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.get_by_role("spinbutton", name="Variación").click()
    page.get_by_role("spinbutton", name="Variación").fill("0.1")
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").fill("10")
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("spinbutton", name="Variación").click()
    page.get_by_role("spinbutton", name="Variación").fill("2")
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

