import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Minutos todo destino nacional").check()
    page.get_by_role("checkbox", name="Minutos red fija on net").check()
    page.locator("mat-stepper").click()
    page.get_by_role("checkbox", name="Minutos red fija off net").check()
    page.get_by_text("Tipo de VariaciónSeleccione").first.click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-16").click()
    page.locator("#mat-input-16").fill("50")
    page.locator("#mat-input-17").click()
    page.locator("#mat-input-17").fill("60")
    page.locator("#mat-select-9").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-18").click()
    page.locator("#mat-input-18").fill("50")
    page.locator("mat-stepper").click()
    page.locator("#mat-input-19").fill("100")
    page.locator("#mat-select-10").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator(".ng-tns-c508571215-52 > app-label-info > .flex.justify-between").click()
    page.locator("#mat-input-20").click()
    page.locator("#mat-input-20").click()
    page.locator("#mat-input-20").click()
    page.locator("#mat-input-20").click()
    page.locator("#mat-input-20").fill("50")
    page.locator("#mat-input-21").click()
    page.locator("#mat-input-21").fill("100")
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

