"""
Acciones sobre tarifas ya registradas (página "Tarifas registradas").

Subcomandos:
    buscar    muestra código SIRT, versión y estados (comercialización / vigencia)
              de cada tarifa, sin cambiar nada.
    renombre  cambia el nombre de cada tarifa (por defecto: nombre actual + sufijo).
    cese      programa el cese de comercialización (fin de comercialización; por
              defecto mañana). Exige "En comercialización" y vigente; pide confirmación.
    fe_erratas agrega "| Fe de erratas <fecha hora>" a todos los campos de texto de
              los pasos 1 y 2 (cualquier servicio), revisa el resumen y guarda.
              --hasta-resumen se detiene en el paso 3 sin guardar.

Flujo grabado (grabaciones/_acciones/):
    Tarifas registradas → Filtro avanzado → "Ingresa nombre de tarifa" → Buscar →
    fila → ⋮ Acciones → Renombre → "Ingresa el nombre de la tarifa" → Guardar →
    Guardar tarifa → Aceptar.

Los estados se leen de la respuesta del filtro (data.items: comercializacion,
vigencia, esUltimaVersion), no de la pantalla. Cese y fe de erratas exigirán
"En comercialización" y vigente; el renombre funciona también con "Por iniciar".

Uso:
    python acciones.py buscar --tarifas TE-SMO-PROMO-001 TE-IF-MUESTRA-001
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --solo-listar
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --diagnostico
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --nuevo-nombre TE-SMO-PROMO-002_NUEVO
    python acciones.py renombre --tarifas TE-AC-PROMO-001 TE-AC-PROMO-002 --sufijo _R2
    python acciones.py cese --tarifas TE-TF-MUESTRA-001 --solo-listar
    python acciones.py cese --tarifas TE-TF-MUESTRA-001 --fecha 2026-10-20 --diagnostico
    python acciones.py fe_erratas --tarifas TE-IF-MUESTRA-001 --hasta-resumen --diagnostico

Resultados (reanudable): resultados/ACCIONES/acciones.csv
"""

import argparse
import csv
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import comun
from empaquetar import texto_dialogo
from ejecutar import Diagnostico


RUTA_RESULTADOS = comun.DIRECTORIO_RESULTADOS / "ACCIONES" / "acciones.csv"
CAMPOS = ["fecha_hora", "accion", "nombre", "codigo_sirt", "version", "nombre_nuevo",
          "comercializacion", "vigencia", "resultado", "detalle"]
TIMEOUT_BUSQUEDA = 20_000
SUFIJO_POR_DEFECTO = "_REN"


# --- Resultados ------------------------------------------------------------------

def registrar(fila: dict) -> None:
    RUTA_RESULTADOS.parent.mkdir(parents=True, exist_ok=True)
    nuevo = not RUTA_RESULTADOS.exists()
    with RUTA_RESULTADOS.open("a", newline="", encoding="utf-8-sig") as salida:
        escritor = csv.DictWriter(salida, fieldnames=CAMPOS, delimiter=";")
        if nuevo:
            escritor.writeheader()
        escritor.writerow({c: fila.get(c, "") for c in CAMPOS})


def ya_hechas(accion: str) -> set[str]:
    if not RUTA_RESULTADOS.exists():
        return set()
    with RUTA_RESULTADOS.open("r", newline="", encoding="utf-8-sig") as entrada:
        return {f["nombre"] for f in csv.DictReader(entrada, delimiter=";")
                if f["accion"] == accion and f["resultado"] == "EXITOSO"}


# --- Búsqueda en "Tarifas registradas" ----------------------------------------------

def abrir_tarifas_registradas(page) -> None:
    origen = re.match(r"https?://[^/]+", page.url).group(0)
    page.keyboard.press("Escape")
    page.goto(f"{origen}/extranet/consulta-tarifa", wait_until="domcontentloaded")
    page.get_by_role("button", name=re.compile("Filtro (avanzado|principal)")).first.wait_for(state="visible")


def abrir_filtro_avanzado(page) -> None:
    campo = page.get_by_role("textbox", name="Ingresa nombre de tarifa")
    if campo.count() and campo.is_visible():
        return
    page.get_by_role("button", name="Filtro avanzado").click()
    campo.wait_for(state="visible")


def _sin_espacios(texto: str) -> str:
    return re.sub(r"\s+", "", texto or "")


def buscar_tarifa(page, nombre: str) -> dict | None:
    """Busca por nombre y devuelve la última versión con ese nombre EXACTO:
    {id, codigo, version, nombre, comercializacion, vigencia, condicion} o None.
    Los datos salen de la respuesta del filtro (la tabla no muestra la vigencia)."""
    respuestas = []

    def capturar(respuesta):
        pedido = respuesta.request
        if pedido.method == "POST" and pedido.resource_type in ("xhr", "fetch"):
            respuestas.append(respuesta)

    abrir_filtro_avanzado(page)
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill(nombre)

    page.on("response", capturar)
    try:
        page.get_by_role("button", name="Buscar", exact=True).click()
        items = None
        for _ in range(TIMEOUT_BUSQUEDA // 250):
            for respuesta in list(respuestas):
                try:
                    datos = respuesta.json().get("data")
                except Exception:
                    respuestas.remove(respuesta)
                    continue
                if isinstance(datos, dict) and isinstance(datos.get("items"), list):
                    items = datos["items"]
                    break
                respuestas.remove(respuesta)
            if items is not None:
                break
            page.wait_for_timeout(250)
    finally:
        page.remove_listener("response", capturar)

    if items is None:
        raise RuntimeError(f"No llegó la respuesta del filtro al buscar '{nombre}'.")

    exactos = [i for i in items if _sin_espacios(i.get("nombre")) == _sin_espacios(nombre)]
    if not exactos:
        return None
    ultimo = max(exactos, key=lambda i: (bool(i.get("esUltimaVersion")), i.get("version") or 0))
    page.wait_for_timeout(300)        # que la tabla termine de pintarse
    return {
        "id": ultimo.get("idTarifa"),
        "codigo": ultimo.get("codigoSIRT", ""),
        "version": str(ultimo.get("version", "")),
        "nombre": ultimo.get("nombre", ""),
        "comercializacion": (ultimo.get("comercializacion") or {}).get("nombre", ""),
        "vigencia": (ultimo.get("vigencia") or {}).get("nombre", ""),
        "condicion": (ultimo.get("condicion") or {}).get("nombre", ""),
        "tipo": ((ultimo.get("tipoTarifa") or {}).get("nombre", "")),
    }


def fila_en_tabla(page, tarifa: dict):
    """La fila de la tabla con ese código SIRT y versión (columnas leídas por su título)."""
    js = """
        ([codigo, version]) => {
            const titulos = Array.from(document.querySelectorAll("thead th"))
                .map(th => th.innerText.normalize("NFD").replace(/[^A-Za-z ]/g, "").replace(/\\s+/g, " ").trim());
            const col = t => titulos.findIndex(x => x.startsWith(t));
            const cCod = col("Codigo SIRT"), cVer = col("Version");
            return Array.from(document.querySelectorAll("tbody tr")).findIndex(fila => {
                const celdas = Array.from(fila.querySelectorAll("td")).map(td => td.innerText.trim());
                const esCodigo = cCod >= 0 ? celdas[cCod] === codigo : celdas.includes(codigo);
                return esCodigo && (cVer < 0 || celdas[cVer] === version);
            });
        }
    """
    argumentos = [tarifa["codigo"], tarifa["version"]]
    try:   # la tabla se pinta un instante después de llegar la respuesta
        page.wait_for_function(f"a => ({js})(a) >= 0", arg=argumentos, timeout=TIMEOUT_BUSQUEDA)
    except Exception:
        raise RuntimeError(f"La tabla no muestra {tarifa['codigo']} versión {tarifa['version']}.")
    indice = page.evaluate(js, argumentos)
    return page.locator("tbody tr").nth(indice)


def abrir_accion(page, fila, accion: str) -> None:
    fila.get_by_label("Acciones").click()
    opcion = page.get_by_role("menuitem", name=accion)
    opcion.wait_for(state="visible")
    if not opcion.is_enabled():
        page.keyboard.press("Escape")
        raise RuntimeError(f"La acción '{accion}' está deshabilitada para esta tarifa.")
    opcion.click()


def confirmar_con_aceptar(page) -> str:
    """Lee el mensaje final, pulsa Aceptar y avisa si es un error."""
    aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    aceptar.first.wait_for(state="visible")
    mensaje = texto_dialogo(page)
    aceptar.first.click()
    # Esperar a que se cierre ESTA ventana: o desaparece el botón, o aparece otra
    # ventana distinta (SIRT a veces muestra dos "Aceptar" seguidos).
    for _ in range(40):                                  # hasta ~10 s
        page.wait_for_timeout(250)
        if not aceptar.first.is_visible() or texto_dialogo(page) != mensaje:
            break
    if re.search(r"error|no se pudo|no se cumplieron|object reference", mensaje, re.I):
        raise RuntimeError(f"SIRT respondió: {mensaje[:400]}")
    return mensaje


# --- Renombre --------------------------------------------------------------------

def renombrar(page, tarifa: dict, nombre_nuevo: str) -> str:
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, "Renombre")

    campo = page.get_by_role("textbox", name="Ingresa el nombre de la tarifa")
    campo.wait_for(state="visible")
    campo.fill(nombre_nuevo)
    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    mensaje = confirmar_con_aceptar(page)

    # Verificación: la tarifa debe encontrarse con el nombre nuevo.
    abrir_tarifas_registradas(page)
    renombrada = buscar_tarifa(page, nombre_nuevo)
    if renombrada is None or renombrada["codigo"] != tarifa["codigo"]:
        raise RuntimeError(f"Tras el renombre, '{nombre_nuevo}' no aparece con el código {tarifa['codigo']}.")
    return mensaje


# --- Cese de comercialización ------------------------------------------------------

class NoAplica(Exception):
    """La tarifa no está en el estado que exige la acción: no se tocó SIRT."""


def validar_para_cese(tarifa: dict, sin_validar: bool, accion: str = "el cese") -> None:
    """SIRT da error si la tarifa no está En comercialización y vigente."""
    problemas = []
    if "en comercializ" not in tarifa["comercializacion"].lower():
        problemas.append(f"comercialización '{tarifa['comercializacion']}'")
    if re.search(r"por iniciar|no vigente|fin", tarifa["vigencia"], re.I):
        problemas.append(f"vigencia '{tarifa['vigencia']}'")
    if problemas and not sin_validar:
        raise NoAplica(f"No se puede aplicar {accion}: " + " y ".join(problemas)
                           + " (debe estar En comercialización y vigente; pasa a ese estado el día"
                             " siguiente al registro definitivo). Usa --sin-validar para intentarlo igual.")


def elegir_fecha_calendario(page, fecha: date) -> None:
    """Abre el único calendario del formulario y avanza mes a mes hasta la fecha."""
    page.get_by_role("button", name="Open calendar").first.click()
    texto = f"{fecha.day} de {comun.MESES[fecha.month - 1]} de {fecha.year}"
    for _ in range(24):
        dia = page.get_by_role("button", name=texto, exact=True)
        if dia.count():
            dia.click()
            page.locator("mat-calendar").wait_for(state="hidden")
            return
        siguiente = page.get_by_role("button", name="Next month")
        if not siguiente.is_enabled():
            break
        siguiente.click()
    raise RuntimeError(f"El calendario no permite elegir el {texto}.")


def cesar(page, tarifa: dict, fecha: date) -> str:
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, "Cese de Comercialización")
    page.get_by_text("Fin de comercialización").first.wait_for(state="visible")
    elegir_fecha_calendario(page, fecha)
    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    return confirmar_con_aceptar(page)


# --- Fe de erratas ----------------------------------------------------------------

def _texto_nuevo(actual: str, sello: str, maximo: int | None) -> str:
    """Quita un sello anterior ('| Fe de erratas ...') y agrega el nuevo, sin pasar del máximo."""
    base = re.sub(r"\s*\|[^|]*\d{2}/\d{2}/\d{4} \d{2}:\d{2}\s*$", "", actual or "").rstrip()
    nuevo = f"{base} {sello}".strip()
    if maximo and len(nuevo) > maximo:
        nuevo = (base[: max(0, maximo - len(sello) - 1)] + " " + sello)[:maximo]
    return nuevo


def corregir_textos(page) -> list[str]:
    """Agrega el sello a TODOS los campos de texto visibles y editables de la pantalla.
    Sirve para cualquier servicio: solo cambia cuántos campos hay."""
    corregidos = []
    for campo in page.locator("textarea").all():
        if not (campo.is_visible() and campo.is_editable()):
            continue
        maximo = campo.get_attribute("maxlength")
        nuevo = _texto_nuevo(campo.input_value(), corregir_textos.sello, int(maximo) if maximo else None)
        campo.fill(nuevo)
        corregidos.append(nuevo)
    return corregidos


def fe_de_erratas(page, tarifa: dict, sello: str, hasta_resumen: bool) -> tuple[int, str]:
    corregir_textos.sello = sello
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, "Fe de Erratas")

    # Paso 1: datos de la tarifa (p. ej. "Otros datos generales").
    page.get_by_role("button", name="Continuar con el paso 2").wait_for(state="visible")
    textos = corregir_textos(page)
    page.get_by_role("button", name="Continuar con el paso 2").click()

    # Paso 2: datos del servicio; si hay pestañas "Servicio N" (paquetes), se corrige cada una.
    continuar3 = page.get_by_role("button", name="Continuar con el paso 3")
    continuar3.wait_for(state="visible")
    page.wait_for_timeout(500)
    pestanas = page.get_by_role("tab", name=re.compile(r"^Servicio \d+$"))
    if pestanas.count():
        for i in range(pestanas.count()):
            pestanas.nth(i).click()
            page.wait_for_timeout(500)
            textos += corregir_textos(page)
    else:
        textos += corregir_textos(page)
    if not textos:
        raise RuntimeError("No se encontró ningún campo de texto para corregir.")
    continuar3.click()

    # Paso 3: el resumen debe mostrar los textos nuevos.
    guardar = page.get_by_role("button", name="Guardar cambios")
    guardar.wait_for(state="visible")
    resumen = page.locator("body").inner_text()
    faltan = [t for t in textos if _sin_espacios(t) not in _sin_espacios(resumen)]
    if faltan:
        raise RuntimeError(f"El resumen no muestra {len(faltan)} de los textos corregidos: {faltan[0][:80]}")
    if hasta_resumen:
        return len(textos), "detenido en el resumen (sin guardar)"

    guardar.click()
    # Puede pedir una confirmación antes del mensaje final.
    for _ in range(60):                                  # hasta ~15 s
        if page.get_by_role("button", name="Aceptar", exact=True).is_visible():
            # Pueden salir varias ventanas seguidas con "Aceptar" (confirmación y resultado):
            # se aceptan todas hasta que pasen ~4 s sin ninguna.
            mensajes = []
            aceptar = page.get_by_role("button", name="Aceptar", exact=True)
            while len(mensajes) < 5:
                mensajes.append(confirmar_con_aceptar(page))
                try:
                    aceptar.first.wait_for(state="visible", timeout=4_000)
                except Exception:
                    break
            return len(textos), " / ".join(mensajes)
        dialogo = page.locator("mat-dialog-container")
        if dialogo.count():
            for texto in ("Guardar tarifa", "Guardar cambios", "Confirmar", "Guardar", "Sí", "Si"):
                boton = dialogo.last.get_by_role("button", name=texto, exact=True)
                if boton.count() and boton.is_visible():
                    boton.click()
                    break
        page.wait_for_timeout(250)
    raise RuntimeError(f"No apareció el mensaje final tras 'Guardar cambios'. Pantalla: {texto_dialogo(page)[:200]}")


# --- Principal -------------------------------------------------------------------

def leer_argumentos():
    parser = argparse.ArgumentParser(description="Acciones sobre tarifas registradas.")
    parser.add_argument("accion", choices=["buscar", "renombre", "cese", "fe_erratas"])
    parser.add_argument("--tarifas", nargs="+", required=True, help="nombres actuales de las tarifas")
    parser.add_argument("--sufijo", default=SUFIJO_POR_DEFECTO, help="renombre: se agrega al nombre actual")
    parser.add_argument("--nuevo-nombre", help="renombre: nombre exacto (solo con una tarifa)")
    parser.add_argument("--fecha", help="cese: fin de comercialización AAAA-MM-DD (por defecto, mañana)")
    parser.add_argument("--sin-validar", action="store_true",
                        help="cese / fe_erratas: intentar aunque la tarifa no esté En comercialización")
    parser.add_argument("--texto", default="Fe de erratas",
                        help="fe_erratas: texto que se agrega a cada campo (con fecha y hora)")
    parser.add_argument("--hasta-resumen", action="store_true",
                        help="fe_erratas: llega al paso 3 y se detiene SIN guardar (para revisar)")
    parser.add_argument("--si", action="store_true", help="no pedir confirmación")
    parser.add_argument("--solo-listar", action="store_true", help="muestra lo que haría, sin tocar SIRT")
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    return parser.parse_args()


def nombre_nuevo_de(nombre: str, args) -> str:
    return args.nuevo_nombre if args.nuevo_nombre else f"{nombre}{args.sufijo}"


def main() -> None:
    args = leer_argumentos()
    if args.nuevo_nombre and len(args.tarifas) != 1:
        sys.exit("--nuevo-nombre solo se puede usar con una tarifa.")

    hechas = ya_hechas(args.accion) if args.accion != "buscar" else set()
    pendientes = [t for t in args.tarifas if t not in hechas]
    for t in args.tarifas:
        if t in hechas:
            print(f"  {t}: ya tiene {args.accion} EXITOSO en {RUTA_RESULTADOS} (se omite)")

    if args.accion == "renombre":
        for t in pendientes:
            print(f"  {t}  →  {nombre_nuevo_de(t, args)}")
    fecha_cese = date.fromisoformat(args.fecha) if args.fecha else date.today() + timedelta(days=1)
    if args.accion == "cese":
        if fecha_cese <= date.today():
            sys.exit("La fecha de cese debe ser posterior a hoy.")
        for t in pendientes:
            print(f"  {t}  →  fin de comercialización {fecha_cese:%d/%m/%Y}")
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT.")
        return
    if not pendientes:
        print("No hay tarifas pendientes.")
        return
    if args.accion == "fe_erratas":
        for t in pendientes:
            print(f"  {t}  →  fe de erratas: a cada campo de texto se le agrega '| {args.texto} <fecha hora>'"
                  + ("  (se detiene en el paso 3, sin guardar)" if args.hasta_resumen else ""))
    if args.accion in ("cese", "fe_erratas") and not args.si and not args.hasta_resumen:
        accion_txt = "el CESE" if args.accion == "cese" else "una FE DE ERRATAS"
        print(f"\nATENCIÓN: se registrará {accion_txt} en {len(pendientes)} tarifa(s) (probablemente irreversible).")
        if input("Escribe 'si' para continuar: ").strip().lower() not in ("si", "sí"):
            print("Cancelado.")
            return

    from playwright.sync_api import sync_playwright

    diag = Diagnostico(args.diagnostico, False, RUTA_RESULTADOS.parent / "diagnostico")
    exitosas = errores = 0

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        diag.iniciar(context)
        if args.al_frente:
            page.bring_to_front()

        for numero, nombre in enumerate(pendientes, start=1):
            print("\n" + "=" * 70)
            print(f"[{numero}/{len(pendientes)}] {nombre}")
            estado = {"etapa": ""}
            fila = {"fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "accion": args.accion, "nombre": nombre}
            diag.iniciar_caso({"correlativo": numero, "nombre": nombre})

            def etapa(nombre_etapa, accion):
                estado["etapa"] = nombre_etapa
                resultado = accion()
                diag.captura_paso(page, nombre_etapa)
                return resultado

            try:
                etapa("abrir_tarifas_registradas", lambda: abrir_tarifas_registradas(page))
                tarifa = etapa("buscar", lambda: buscar_tarifa(page, nombre))
                if tarifa is None:
                    raise RuntimeError(f"'{nombre}' no aparece en Tarifas registradas.")
                fila.update(codigo_sirt=tarifa["codigo"], version=tarifa["version"],
                            comercializacion=tarifa["comercializacion"], vigencia=tarifa["vigencia"])
                print(f"  {tarifa['codigo']} v{tarifa['version']} | {tarifa['tipo']} {tarifa['condicion']}"
                      f" | {tarifa['comercializacion']} | {tarifa['vigencia']}")

                if args.accion == "buscar":
                    fila["resultado"] = "ENCONTRADA"
                elif args.accion == "renombre":
                    nuevo = nombre_nuevo_de(nombre, args)
                    fila["nombre_nuevo"] = nuevo
                    mensaje = etapa("renombre", lambda: renombrar(page, tarifa, nuevo))
                    fila["resultado"] = "EXITOSO"
                    fila["detalle"] = mensaje[:300]
                    print(f"  Renombrada: {nuevo}")
                elif args.accion == "cese":
                    validar_para_cese(tarifa, args.sin_validar)
                    mensaje = etapa("cese", lambda: cesar(page, tarifa, fecha_cese))
                    fila["resultado"] = "EXITOSO"
                    fila["detalle"] = f"fin comercialización {fecha_cese:%d/%m/%Y} | {mensaje[:250]}"
                    print(f"  Cese programado: fin de comercialización {fecha_cese:%d/%m/%Y}")
                elif args.accion == "fe_erratas":
                    validar_para_cese(tarifa, args.sin_validar, "la fe de erratas")
                    sello = f"| {args.texto} {datetime.now():%d/%m/%Y %H:%M}"
                    cambios, mensaje = etapa("fe_erratas", lambda: fe_de_erratas(page, tarifa, sello, args.hasta_resumen))
                    fila["resultado"] = "EN_RESUMEN" if args.hasta_resumen else "EXITOSO"
                    fila["detalle"] = f"{cambios} campo(s) {sello} | {mensaje[:250]}"
                    print(f"  {cambios} campo(s) modificados"
                          + (" — detenido en el paso 3 SIN guardar." if args.hasta_resumen else ". Guardado."))
                exitosas += 1

            except NoAplica as aviso:
                fila["resultado"] = "NO_APLICA"
                fila["detalle"] = str(aviso)[:500]
                print(f"  NO APLICA (no se tocó SIRT): {aviso}")

            except Exception as error:
                diag.registrar_fallo(page, error, estado["etapa"])
                fila["resultado"] = "ERROR"
                fila["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
                errores += 1
                print(f"  ERROR en '{estado['etapa']}': {error}")
                page.keyboard.press("Escape")

            diag.cerrar_caso(fila["resultado"] == "ERROR")
            if args.accion != "buscar":
                registrar(fila)

        diag.finalizar()

    print("\n" + "=" * 70)
    print(f"Correctas: {exitosas} | Errores: {errores}")
    if args.accion != "buscar":
        print(f"Resultados: {RUTA_RESULTADOS.resolve()}")


if __name__ == "__main__":
    main()
