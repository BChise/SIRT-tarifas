import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="Duración").check()
    page.locator("#mat-select-0").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("div").filter(has_text=re.compile(r"^Variación$")).first.click()
    page.locator("#mat-input-0").click()
    page.locator("#mat-input-0").fill("60")
    page.locator("#mat-input-1").click()
    page.locator("#mat-input-1").fill("180")
    page.get_by_role("checkbox", name="Datos para tethering").check()
    page.locator("#mat-select-1").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("mat-stepper").click()
    page.locator("#mat-input-2").fill("10")
    page.locator("#mat-input-3").click()
    page.locator("#mat-input-3").fill("60")
    page.get_by_role("checkbox", name="Datos para Apps 1").check()
    page.locator("mat-stepper").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-8").click()
    page.locator("#mat-input-8").fill("12")
    page.locator("#mat-input-9").click()
    page.locator("#mat-input-9").fill("120")
    page.get_by_role("checkbox", name="Datos tope descarga").check()
    page.locator("mat-stepper").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-17").click()
    page.locator("#mat-input-17").fill("50")
    page.locator("#mat-input-18").click()
    page.locator("#mat-input-18").fill("120")
    page.get_by_role("checkbox", name="Datos Roaming incluidos en").check()
    page.locator("#mat-select-21").click()
    page.get_by_role("listbox", name="Tipo de Variación").click()
    page.locator("#mat-input-41").click()
    page.locator("#mat-input-41").fill("120")
    page.locator("mat-stepper").click()
    page.locator("#mat-input-42").fill("40")
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

