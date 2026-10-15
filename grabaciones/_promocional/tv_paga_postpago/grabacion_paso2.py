import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Canales analógicos").check()
    page.get_by_role("button").nth(3).click()
    page.locator("input[name=\"nameInputRamdom0bf87360-52db-4778-8136-14cdec81e3cb\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator("#mat-input-1").click()
    page.locator("#mat-input-1").fill("30")
    page.get_by_role("checkbox", name="Decodificadores o similares").check()
    page.get_by_role("combobox", name="Tipo de Variación Seleccione").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-6").click()
    page.locator("#mat-input-6").fill("3")
    page.locator("#mat-input-7").click()
    page.locator("#mat-input-7").fill("180")
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

