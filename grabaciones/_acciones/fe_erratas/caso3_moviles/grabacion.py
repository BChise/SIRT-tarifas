import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.locator(".mat-mdc-form-field-infix.ng-tns-c508571215-23").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-SMP-MUESTRA-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Fe de Erratas").click()
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizadaasfasfasfas")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Condiciones generales de prueba automatizadaasfasf")
    page.get_by_role("textbox", name="Condiciones cantidad datos").click()
    page.get_by_role("textbox", name="Condiciones cantidad datos").fill("Condiciones de datos de prueba automatizadaasfasfasf")
    page.get_by_role("textbox", name="Condiciones bolsa de datos").click()
    page.get_by_role("textbox", name="Condiciones bolsa de datos").fill("Condiciones de bolsa de pruebafsafasf")
    page.get_by_role("textbox", name="Condiciones apps zero rating").click()
    page.get_by_role("textbox", name="Condiciones apps zero rating").fill("Condiciones zero rating de pruebaasfasfsaf")
    page.get_by_role("textbox", name="Condiciones datos tethering").click()
    page.get_by_role("textbox", name="Condiciones datos tethering").fill("Condiciones tethering de pruebaasfasfsa")
    page.get_by_role("textbox", name="Condiciones traslada datos").click()
    page.get_by_role("textbox", name="Condiciones traslada datos").fill("Condiciones traslada datos de pruebaasfsafas")
    page.get_by_role("textbox", name="Condiciones velocidades").click()
    page.get_by_role("textbox", name="Condiciones velocidades").fill("Condiciones velocidades de pruebasafasfasf")
    page.get_by_role("textbox", name="Condiciones Roaming").click()
    page.get_by_role("textbox", name="Condiciones Roaming").fill("Condiciones roaming de pruebaasfsafasf")
    page.get_by_role("textbox", name="Condiciones LDI").click()
    page.get_by_role("textbox", name="Condiciones LDI").fill("Condiciones LDI de pruebaasfasfsaf")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("Otras condiciones de prueba automatizadaasfasfasfasf")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("Restricciones técnicas de prueba automatizadaasfasf")
    page.get_by_role("textbox", name="Restricciones geográficas").click()
    page.get_by_role("textbox", name="Restricciones geográficas").fill("Restricciones geográficas de prueba automatizadaasfasfsaf")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("Otras restricciones de prueba automatizadaasfasfsafsafaf")
    page.get_by_role("button", name="Continuar con el paso").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

