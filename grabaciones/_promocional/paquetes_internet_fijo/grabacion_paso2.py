import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Velocidad de descarga contratada").check()
    page.get_by_role("checkbox", name="Velocidad de descarga garantizada").check()
    page.locator("mat-stepper").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("mat-stepper").click()
    page.locator("#mat-input-54").fill("100")
    page.locator("mat-stepper").click()
    page.locator("#mat-input-55").fill("60")
    page.locator("#mat-select-38").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-56").click()
    page.locator("#mat-input-56").fill("40")
    page.locator("#mat-input-57").click()
    page.locator("#mat-input-57").fill("30")
    page.get_by_role("checkbox", name="Cantidad de datos").check()
    page.locator("#mat-select-41").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-63").click()
    page.locator("#mat-input-63").fill("60")
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

