"""
Completa el «Código SIRT» (p. ej. TERDS2026000123) de las tarifas YA registradas en los Excel de
automatizacion/ y deja el número interno (ID, p. ej. 26853) en la columna «ID de registro».

Desde el 21/10 el registro definitivo anota el código solo (resultados/codigos_sirt.csv); este
script es para lo registrado antes. Pasos:
1. Lee los Excel (establecidas, empaquetados, LyC y promociones): filas con «Registro definitivo»
   (o «Ya registrada») cuyo «Código SIRT» esté vacío o sea el ID.
2. Busca en SIRT («Tarifas registradas»): primero por grupo (p. ej. «TE-SMP-CM-»); lo que no
   aparezca así, una por una. Si la tarifa fue renombrada, la busca por su nombre actual.
3. Anota los códigos en resultados/codigos_sirt.csv y los escribe en los Excel (con copia).

Uso (Excel CERRADOS, Edge con SIRT abierto):
    python completar_codigos_sirt.py --solo-listar
    python completar_codigos_sirt.py --limite 3
    python completar_codigos_sirt.py
    python completar_codigos_sirt.py --solo-excel     (sin SIRT: solo pasa al Excel lo ya anotado)
"""

import argparse
import re
import time

import comun
import excel_comun as xc

EXCELS = [("registro_tarifas_establecidas.xlsx", None), ("empaquetados.xlsx", "Empaquetados"),
          ("registro_licitaciones.xlsx", "Licitaciones"), ("promociones.xlsx", "Promociones")]
ESTADOS = ("Registro definitivo", "Ya registrada")


def leer_argumentos():
    p = argparse.ArgumentParser(description="Completa el Código SIRT en los Excel.")
    p.add_argument("--solo-listar", action="store_true")
    p.add_argument("--solo-excel", action="store_true", help="no entra a SIRT; solo escribe lo ya anotado")
    p.add_argument("--limite", type=int, default=0, help="buscar solo N tarifas en SIRT (prueba)")
    return p.parse_args()


def hojas(libro, solo):
    for titulo in ([solo] if solo else libro.sheetnames):
        if titulo not in libro.sheetnames:
            continue
        hoja = xc.Hoja(libro[titulo])
        if hoja.c_estado and hoja.c_nombre and hoja.c_codigo:
            yield hoja


def pendientes_de(hoja) -> list[tuple[int, str]]:
    salida = []
    for r in range(xc.FILA_CABECERA + 1, hoja.hoja.max_row + 1):
        nombre = xc.texto(hoja.hoja.cell(r, hoja.c_nombre).value)
        codigo = xc.texto(hoja.hoja.cell(r, hoja.c_codigo).value)
        if nombre and hoja.estado(r) in ESTADOS and (not codigo or codigo.isdigit()):
            salida.append((r, nombre))
    return salida


def grupo(nombre: str) -> str:
    """«TE-SMP-CM-011» → «TE-SMP-CM-» (la parte común para buscar varias de una vez)."""
    return re.sub(r"\d+$", "", nombre) if re.search(r"\d+$", nombre) else nombre


def buscar_en_sirt(nombres: list[str]) -> int:
    from playwright.sync_api import sync_playwright
    import acciones
    from acciones_excel import nombre_actual, nombres_renombrados
    renombrados = nombres_renombrados()
    actual = {n: nombre_actual(n, renombrados) for n in nombres}
    faltan = set(nombres)
    hallados = 0

    def anotar(items) -> None:
        nonlocal hallados
        por_nombre = {}
        for item in items:
            clave = acciones._sin_espacios(item.get("nombre"))
            previo = por_nombre.get(clave)
            if previo is None or (item.get("version") or 0) >= (previo.get("version") or 0):
                por_nombre[clave] = item
        for n in list(faltan):
            item = por_nombre.get(acciones._sin_espacios(actual[n]))
            if item and item.get("codigoSIRT"):
                comun.guardar_codigo_sirt(n, item["codigoSIRT"], str(item.get("idTarifa") or ""))
                faltan.discard(n)
                hallados += 1
                print(f"   {n} → {item['codigoSIRT']} (ID {item.get('idTarifa')})")

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        page = comun.obtener_pagina_sirt(browser.contexts[0])
        page.set_default_timeout(comun.TIMEOUT)
        acciones.abrir_tarifas_registradas(page)
        grupos = sorted({grupo(actual[n]) for n in nombres})
        print(f"\nBúsqueda por grupo ({len(grupos)} grupo(s)):")
        for g in grupos:
            try:
                items, total = acciones.items_del_filtro(page, g, con_total=True)
                print(f" · {g}: SIRT devolvió {len(items)} de {total}")
                anotar(items)
            except Exception as error:
                print(f" · {g}: no se pudo ({str(error)[:100]})")
        if faltan:
            print(f"\nUna por una ({len(faltan)}):")
        for n in sorted(faltan):
            try:
                anotar(acciones.items_del_filtro(page, actual[n]))
                if n in faltan:
                    print(f"   {n}: no aparece en Tarifas registradas")
            except Exception as error:
                print(f"   {n}: error ({str(error)[:100]})")
    return hallados


def main() -> None:
    args = leer_argumentos()
    inicio = time.monotonic()
    libros = []
    for archivo, hoja_unica in EXCELS:
        ruta = xc.AUTOMATIZACION / archivo
        if not ruta.exists():
            continue
        libro = xc.abrir_libro(ruta, args.solo_listar)
        filas = [(h, r, n) for h in hojas(libro, hoja_unica) for r, n in pendientes_de(h)]
        libros.append((ruta, libro, filas))
        print(f"{archivo}: {len(filas)} fila(s) sin Código SIRT")
    todas = [n for _, _, filas in libros for _, _, n in filas]
    sin_codigo = [n for n in dict.fromkeys(todas) if not comun.codigo_sirt_registrado(n)]
    print(f"Total: {len(todas)} fila(s); ya anotadas en codigos_sirt.csv: {len(todas) - len(sin_codigo)}; "
          f"a buscar en SIRT: {len(sin_codigo)}")
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT ni los Excel.")
        return
    if sin_codigo and not args.solo_excel:
        buscar_en_sirt(sin_codigo[: args.limite] if args.limite else sin_codigo)

    print()
    for ruta, libro, filas in libros:
        escritas = 0
        for hoja, r, nombre in filas:
            codigo, ident = comun.codigo_sirt_registrado(nombre, True)
            if not codigo:
                continue
            anterior = xc.texto(hoja.hoja.cell(r, hoja.c_codigo).value)
            ident = anterior if anterior.isdigit() else ident
            if ident:
                hoja.hoja.cell(r, hoja._columna_id(), ident)
            hoja.hoja.cell(r, hoja.c_codigo, codigo)
            escritas += 1
        if escritas:
            copia = xc.guardar_con_copia(libro, ruta, "CODIGOS_SIRT", ruta.stem)
            print(f"{ruta.name}: {escritas} código(s) escrito(s) (copia: {copia.name})")
        else:
            print(f"{ruta.name}: sin cambios")
    segundos = int(time.monotonic() - inicio)
    print(f"Tiempo total: {segundos // 60} min {segundos % 60} s")


if __name__ == "__main__":
    main()
