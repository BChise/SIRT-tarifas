import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("combobox", name="Tipo de moneda Seleccione").click()
    page.get_by_role("listbox", name="Tipo de moneda").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("124")
    page.get_by_role("combobox", name="Periodicidad Seleccione").click()
    page.get_by_role("listbox", name="Periodicidad").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asff")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("asfsg")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asfg")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("asdasf")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

