import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("124")
    page.get_by_role("combobox", name="Periodicidad Seleccione").click()
    page.get_by_role("listbox", name="Periodicidad").click()
    page.get_by_role("combobox", name="Tipo de Acceso (TV)").click()
    page.get_by_role("listbox", name="Tipo de Acceso (TV)").click()
    page.get_by_role("listbox", name="Tipo de Acceso (TV)").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("combobox", name="Tecnologia Seleccione").click()
    page.get_by_role("listbox", name="Tecnologia").click()
    page.locator(".ml-4.mdc-button.mdc-button--unelevated").first.click()
    page.locator("input[name=\"nameInputRamdom5cf8952d-b90c-4fbf-80ef-0df60371ab43\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdomb7269a98-61f8-4a9e-8934-c00fd18fd137\"]").set_input_files("Canales analógicos_1.xlsx")
    page.get_by_role("spinbutton", name="Decodificadores o similares").click()
    page.get_by_role("spinbutton", name="Decodificadores o similares").fill("12")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asdf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asdf")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdomd8f3187c-3944-4703-81ac-a5a8f920208d\"]").set_input_files("Canales analógicos_1.xlsx")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

