import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("124214")
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").fill("12647547")
    page.get_by_role("combobox", name="Tipo de Acceso (TV)").click()
    page.get_by_role("listbox", name="Tipo de Acceso (TV)").click()
    page.get_by_role("option", name="Migración").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("combobox", name="Tecnologia Seleccione").click()
    page.get_by_role("listbox", name="Tecnologia").click()
    page.get_by_role("button").nth(5).click()
    page.locator("input[name=\"nameInputRamdom1a1af60e-d959-43a2-a31f-34160250cafb\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdomeedce8b2-5bcb-4cca-970b-3c0f507d96e8\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdoma8dc4853-28d8-4d3a-9e2e-b97c001a95c7\"]").set_input_files("Canales analógicos_1.xlsx")
    page.get_by_role("spinbutton", name="Decodificadores o similares").click()
    page.get_by_role("spinbutton", name="Decodificadores o similares").fill("4")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asdfg")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asdfg")
    page.get_by_role("tablist").click()
    page.get_by_role("listbox", name="Tecnologia").click()
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

