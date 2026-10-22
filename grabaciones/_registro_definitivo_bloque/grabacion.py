# Pega aquí el código que genera el Inspector de Playwright (python grabar_sirt.py > Record).
import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page.get_by_role("checkbox", name="select row 2026-10-16T15:24:421").check()
    page.get_by_role("checkbox", name="select row 2026-10-16T15:24:311").check()
    page.get_by_role("checkbox", name="select row 2026-10-16T15:24:181").check()
    page.get_by_role("button", name="Registro de Disponibilidad").click()
    page.get_by_role("radio", name="Selector geográfico").check()
    page.get_by_role("button", name="Continuar").click()
    page.locator("div:nth-child(15) > .region-header").click()
    page.get_by_role("checkbox", name="Lima").check()
    page.get_by_role("button", name="Guardar").click()
    page.get_by_role("button", name="Aceptar").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
