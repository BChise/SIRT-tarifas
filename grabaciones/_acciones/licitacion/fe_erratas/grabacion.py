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
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("textbox", name="Documento de Empresa").click()
    page.get_by_role("textbox", name="Documento de Empresa").fill("1242154525123")
    page.get_by_role("textbox", name="Nombre de empresa").click()
    page.get_by_role("textbox", name="Nombre de empresa").fill("example sac_fe")
    page.get_by_role("textbox", name="Descripción del Servicio").click()
    page.get_by_role("textbox", name="Descripción del Servicio").fill("asdf_fe")
    page.get_by_role("textbox", name="Atributos de la tarifa").click()
    page.get_by_role("textbox", name="Atributos de la tarifa").fill("asdf_fe")
    page.get_by_role("textbox", name="Ámbito geográfico").fill("asdf_fe")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.locator("div").filter(has_text=re.compile(r"^Detalle del servicio$")).nth(2).click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("sadasffasfasf_fe")
    page.get_by_role("tab", name="Servicio 2").click()
    page.get_by_role("textbox", name="Detalle del servicio").click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("asdasfasfasf_fe")
    page.get_by_role("tab", name="Servicio 3").click()
    page.get_by_role("textbox", name="Detalle del servicio").click()
    page.get_by_role("textbox", name="Detalle del servicio").fill("safasfasfasf_fe")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

