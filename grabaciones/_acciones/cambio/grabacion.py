import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("button", name="Filtro avanzado").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").click()
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill("TE-OS-MUESTRA-001")
    page.get_by_role("button", name="Buscar").click()
    page.get_by_role("cell", name="Acciones").get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Cambio de Tarifa").click()
    page.get_by_role("button", name="Open calendar").click()
    page.get_by_role("button", name="19 de octubre de").click()
    page.locator("div").filter(has_text=re.compile(r"^Tarifa de instalación$")).nth(5).click()
    page.get_by_role("textbox", name="Tarifa de instalación").fill("10")
    page.get_by_role("textbox", name="Otros datos generales").click()
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectificacion | Rectificación 15/10/2026 17:03 ")
    page.get_by_role("textbox", name="Otros datos generales").press("CapsLock")
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectificacion | Rectificación 15/10/2026 17:03 C")
    page.get_by_role("textbox", name="Otros datos generales").press("CapsLock")
    page.get_by_role("textbox", name="Otros datos generales").fill("Otros datos de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectificacion | Rectificación 15/10/2026 17:03 CAmbio")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").click()
    page.get_by_role("textbox", name="Valor del servicio Soles").fill("49")
    page.get_by_role("textbox", name="Condiciones generales").click()
    page.get_by_role("textbox", name="Condiciones generales").fill("Condiciones generales de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectif | Rectificación 15/10/2026 17:03 cambio")
    page.get_by_role("textbox", name="Otras condiciones tarifa").click()
    page.get_by_role("textbox", name="Otras condiciones tarifa").fill("Otras condiciones de prueba automatizada | Fe de erratas 15/10/2026 15:21 | Rectificación 15/10/2026 17:03 cambio")
    page.get_by_role("textbox", name="Restricciones técnicas").click()
    page.get_by_role("textbox", name="Restricciones técnicas").fill("Restricciones técnicas de prueba automatizada | Fe de erratas 15/10/2026 15:21 | Rectificación 15/10/2026 17:03 cambio")
    page.get_by_role("textbox", name="Otras restricciones").click()
    page.get_by_role("textbox", name="Otras restricciones").fill("Otras restricciones de prueba automatizada | Fe de erratas 15/10/2026 15:21 rectif | Rectificación 15/10/2026 17:03 cambio")
    page.get_by_role("button", name="Continuar con el paso").click()
    page.get_by_role("button", name="Guardar cambios").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

