import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Minutos todo destino nacional").check()
    page.get_by_role("checkbox", name="Minutos móvil on net").check()
    page.locator("mat-stepper").dblclick()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-select-3").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-0").click()
    page.locator("#mat-input-0").fill("100")
    page.locator("#mat-input-6").click()
    page.locator("#mat-input-6").fill("100")
    page.locator("#mat-input-1").click()
    page.locator("#mat-input-1").fill("30")
    page.locator("#mat-input-7").click()
    page.locator("#mat-input-7").fill("30")
    page.get_by_role("checkbox", name="Tarifa por minuto a móvil LDI").check()
    page.locator("#mat-select-6").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-12").click()
    page.locator("#mat-input-12").fill("10")
    page.locator(".mat-mdc-form-field-infix.ng-tns-c508571215-33").click()
    page.locator("#mat-input-13").fill("30")
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

