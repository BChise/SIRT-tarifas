import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Velocidad de descarga contratada", exact=True).check()
    page.get_by_role("checkbox", name="Velocidad de descarga garantizada", exact=True).check()
    page.locator("#mat-select-0").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-0").click()
    page.locator("#mat-input-0").fill("50")
    page.locator("#mat-input-1").click()
    page.locator("#mat-input-1").fill("30")
    page.locator(".mat-mdc-select-arrow.ng-tns-c1073613477-11 > svg").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-2").click()
    page.locator("#mat-select-1").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.get_by_role("spinbutton", name="Variación", description="Este apartado es obligatorio.").click()
    page.get_by_role("spinbutton", name="Variación", description="Este apartado es obligatorio.").fill("100")
    page.locator("#mat-input-3").click()
    page.locator("#mat-input-3").fill("60")
    page.locator("mat-stepper").click()
    page.get_by_role("checkbox", name="Velocidad de carga contratada", exact=True).check()
    page.locator("#mat-select-2").click()
    page.locator(".cdk-overlay-backdrop").dblclick()
    page.get_by_role("combobox", name="Tipo de Variación Seleccione", description="Este apartado es obligatorio.").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-4").click()
    page.locator("#mat-input-4").fill("30")
    page.locator("#mat-input-5").click()
    page.locator("#mat-input-5").fill("30")
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

