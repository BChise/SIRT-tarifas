import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-OS-MUESTRA-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.locator("div").filter(has_text=re.compile(r"^Otros datos generales$")).nth(2).click()
    page.locator("div").filter(has_text=re.compile(r"^Otros datos generales$")).nth(2).click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizadaasfasfasfasf")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Condiciones generales de prueba automatizadaasfasfasf")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("Otras condiciones de prueba automatizadaasfasfasf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("Restricciones técnicas de prueba automatizadaasfasfasfasf")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("Otras restricciones de prueba automatizadaasfasfasfasf")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Continuar con el paso").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

