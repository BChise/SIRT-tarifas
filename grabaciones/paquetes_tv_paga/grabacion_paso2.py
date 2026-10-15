import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("1241254")
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("spinbutton", name="Duración Días").fill("1675")
    page.get_by_role("combobox", name="Tipo de Contratación").click()
    page.get_by_role("option", name="Postpago").click()
    page.get_by_role("option", name="Prepago").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("tablist").click()
    page.locator(".cdk-overlay-backdrop").click()
    page.get_by_role("button").nth(4).click()
    page.locator("input[name=\"nameInputRamdom73207942-75cf-469a-85c4-3edf0ce2d700\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdom4c974fb0-44e4-443e-86ab-5fafb329e99e\"]").set_input_files("Canales analógicos_1.xlsx")
    page.locator(".ml-4.mdc-button.mdc-button--unelevated.mat-mdc-unelevated-button.mat-primary.mat-mdc-button-base.cdk-focused").click()
    page.locator("input[name=\"nameInputRamdom8904b4e0-a471-4e38-89e7-70d4e636471d\"]").set_input_files("Canales analógicos_1.xlsx")
    page.get_by_role("tablist").click()
    page.get_by_role("spinbutton", name="Decodificadores o similares").fill("123")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("asfasfg")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("asggg")
    page.get_by_role("button", name="Guardar y continuar con el paso").click()
    page.get_by_role("button", name="Confirmar Tarifa").click()
    page.get_by_role("button", name="Confirmar tarifa").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

