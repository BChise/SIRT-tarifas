import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("1245")
    page.get_by_role("tablist").click()
    page.get_by_role("spinbutton", name="Duración Días").fill("125")
    page.get_by_role("tablist").click()
    page.get_by_role("option", name="Postpago").click()
    page.get_by_role("listbox", name="Tipo de línea").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("button", name="Agregar").first.click()
    page.get_by_role("combobox", name="Tipo").click()
    page.get_by_role("listbox", name="Tipo").click()
    page.get_by_role("spinbutton", name="Tarifa").click()
    page.get_by_role("spinbutton", name="Tarifa").fill("12")
    page.get_by_role("combobox", name="Países").click()
    page.get_by_role("option", name="ALEMANIA").click()
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Agregar").nth(1).click()
    page.get_by_role("combobox", name="Tipo").click()
    page.get_by_role("listbox", name="Tipo").click()
    page.get_by_role("spinbutton", name="Tarifa").click()
    page.get_by_role("spinbutton", name="Tarifa").fill("124")
    page.get_by_role("combobox", name="Países").click()
    page.get_by_role("combobox", name="Países").fill("")
    page.get_by_role("option", name="ALEMANIA").click()
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asfasf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asfasf")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

