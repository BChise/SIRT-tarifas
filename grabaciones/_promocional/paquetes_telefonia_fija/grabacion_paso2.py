import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Tarifa por minuto a fijo LDI").check()
    page.locator("#mat-select-0").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-0").click()
    page.locator("#mat-input-0").fill("1")
    page.locator("#mat-input-1").click()
    page.locator("#mat-input-1").fill("30")
    page.locator("mat-stepper").click()
    page.get_by_role("checkbox", name="Tarifa por minuto a móvil LDI").check()
    page.get_by_text("Tipo de VariaciónSeleccione").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-2").click()
    page.locator("#mat-input-2").fill("1")
    page.locator("#mat-input-3").click()
    page.locator("#mat-input-3").fill("30")
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

