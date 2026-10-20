import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TP-TF-PROMO-017")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Renovar").click()
    page.locator("mat-form-field").filter(has_text="Fecha inicio de la").get_by_label("Open calendar").click()
    page.get_by_role("button", name="15 de diciembre de").click()
    page.locator("mat-form-field").filter(has_text="Fecha fin de la").get_by_label("Open calendar").click()
    page.get_by_role("button", name="Next month").click()
    page.get_by_role("button", name="8 de enero de 2027", exact=True).click()
    page.get_by_role("tabpanel", name="Editable Datos de Tarifa").locator("mat-form-field").filter(has_text="Tipo de VariaciónPorcentual").click()
    page.get_by_role("tabpanel", name="Editable Datos de Tarifa").get_by_role("spinbutton", name="Duración Días").click()
    page.get_by_role("tabpanel", name="Editable Datos de Tarifa").get_by_role("spinbutton", name="Duración Días").fill("25")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.locator("#mat-input-19").click()
    page.locator("#mat-input-18").click()
    page.locator("#mat-input-19").fill("25")
    page.locator(".mat-mdc-form-field-icon-suffix.ng-tns-c508571215-59").click()
    page.locator("#mat-input-21").fill("25")
    page.locator(".mat-icon.notranslate.icon-size-5.mat-icon-no-color.ng-tns-c508571215-63 > svg").click()
    page.locator("#mat-input-23").fill("25")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

