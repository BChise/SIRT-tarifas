import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("LYC de prueba 16 oct")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Rectificar").click()
    page.locator("#mat-input-14").click()
    page.locator("#mat-input-14").press("ArrowRight")
    page.locator("#mat-input-14").fill("16/11/2026")
    page.locator("#mat-input-15").click()
    page.get_by_role("button", name="Open calendar").first.click()
    page.get_by_role("button", name="20 de octubre de").click()
    page.get_by_role("button", name="Open calendar").nth(1).click()
    page.get_by_role("button", name="Next month").click()
    page.get_by_role("button", name="31 de diciembre de").click()
    page.get_by_role("button", name="Open calendar").nth(1).click()
    page.get_by_role("button", name="Next month").click()
    page.get_by_role("button", name="31 de enero de").click()
    page.get_by_role("combobox", name="Ámbito Público").click()
    page.get_by_role("option", name="Privado").click()
    page.get_by_role("textbox", name="Valor de contrato Soles").click()
    page.get_by_role("textbox", name="Valor de contrato Soles").fill("12124")
    page.get_by_text("MonedaSoles").click()
    page.get_by_role("listbox", name="Moneda").click()
    page.get_by_role("textbox", name="Nombre de empresa").click()
    page.get_by_role("textbox", name="Nombre de empresa").fill("example sac_fe_rect")
    page.get_by_role("textbox", name="Descripción del Servicio").click()
    page.get_by_role("textbox", name="Descripción del Servicio").fill("asdf_fe_rect")
    page.get_by_role("textbox", name="Atributos de la tarifa").fill("asdf_fe_rect")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("textbox", name="Detalle del servicio").click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("sadasffasfasf_fe_rect")
    page.get_by_role("tab", name="Servicio 2").click()
    page.get_by_role("textbox", name="Detalle del servicio").click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("asdasfasfasf_fe-_rect")
    page.locator("span").filter(has_text="Servicio").nth(5).click()
    page.get_by_role("textbox", name="Detalle del servicio").click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("safasfasfasf_fe_rect")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

