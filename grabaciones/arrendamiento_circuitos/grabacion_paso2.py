import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("combobox", name="Tipo de moneda Seleccione").click()
    page.get_by_role("listbox", name="Tipo de moneda").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("12")
    page.get_by_role("spinbutton", name="Periodo de contrato Meses").click()
    page.get_by_role("spinbutton", name="Periodo de contrato Meses").fill("123")
    page.get_by_role("tablist").click()
    page.get_by_role("listbox", name="Periodicidad").click()
    page.get_by_role("textbox", name="Tecnología de transmisión").click()
    page.get_by_role("textbox", name="Tecnología de transmisión").fill("tecnologia_1_escrito")
    page.get_by_role("textbox", name="Velocidad de transmisión Mbps").click()
    page.get_by_role("textbox", name="Velocidad de transmisión Mbps").fill("14")
    page.get_by_role("textbox", name="Vel. transmisión garant. Mbps").click()
    page.get_by_role("textbox", name="Vel. transmisión garant. Mbps").fill("14")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asdf")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("asdf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asdf")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("asdf")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

