import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("1245")
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").fill("1251")
    page.get_by_role("combobox", name="Tecnología (IF) Seleccione").click()
    page.get_by_role("listbox", name="Tecnología (IF)").click()
    page.get_by_role("button", name="Agregar").click()
    page.get_by_role("combobox", name="Tecnología").click()
    page.get_by_role("listbox", name="Tecnología").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").fill("124")
    page.locator("#mat-mdc-dialog-4").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga garantizada Mbps").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga garantizada Mbps").fill("124")
    page.get_by_role("spinbutton", name="Velocidad de carga contratada").click()
    page.get_by_role("spinbutton", name="Velocidad de carga contratada").fill("124")
    page.get_by_role("spinbutton", name="Velocidad de carga garantizada Mbps").click()
    page.get_by_role("spinbutton", name="Velocidad de carga garantizada Mbps").fill("124")
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("tablist").click()
    page.get_by_role("spinbutton", name="Cantidad de datos GB").fill("12")
    page.get_by_role("spinbutton", name="Repetidores y/o equipos").click()
    page.get_by_role("spinbutton", name="Repetidores y/o equipos").fill("52")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("sadasf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asfasfas")
    page.get_by_role("combobox", name="Tecnología (IF) VHDL").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("button", name="Agregar").click()
    page.get_by_role("combobox", name="Tecnología").click()
    page.get_by_role("listbox", name="Tecnología").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").fill("123")
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").click()
    page.get_by_role("spinbutton", name="Velocidad de descarga contratada Mbps").fill("")
    page.get_by_role("button").first.click()
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

