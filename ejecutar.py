"""
Ejecutor único: crea las combinaciones de tarifas de un servicio.

Uso:
    python ejecutar.py                              # lista los servicios disponibles
    python ejecutar.py internet_fijo --solo-calcular
    python ejecutar.py internet_fijo --limite 2
    python ejecutar.py internet_fijo

Opciones:
    --solo-calcular   muestra el total de casos y guarda el plan, sin abrir el navegador
    --limite N        ejecuta solo los primeros N casos pendientes
                      (sin --limite pide confirmación antes de crear más de 1 tarifa)
    --si              no pide confirmación (para lotes completos ya decididos)
    --al-frente       trae la ventana de Edge al frente al empezar (por defecto
                      no, para poder seguir usando el equipo mientras corre)
    --lote TEXTO      identificador del lote (por defecto, la fecha de hoy);
                      con el mismo lote, una nueva ejecución continúa donde quedó
    --diagnostico     guarda una captura después de cada paso y, si un caso falla,
                      el "mapa" de la página (aria snapshot) y el detalle del error
                      en resultados/<CODIGO>/diagnostico_<lote>/<correlativo>/
    --trace           además del diagnóstico, graba un trace de Playwright y lo
                      guarda solo de los casos que fallan (implica --diagnostico).
                      Se abre con: playwright show-trace <archivo>.zip

Ejemplo de prueba con diagnóstico:
    python ejecutar.py telefonia_fija --limite 1 --diagnostico

Requisitos para ejecutar (no para --solo-calcular):
- Edge abierto en modo controlable (puerto 9222).
- Sesión iniciada en SIRT y pantalla en "Tarifas registradas".
"""

import argparse
import importlib
import sys
import traceback
from datetime import datetime
from pathlib import Path

import comun


PAUSA_ENTRE_CASOS_MS = 1_000


def servicios_disponibles() -> list[str]:
    carpeta = Path(__file__).parent / "servicios"
    return sorted(
        archivo.stem
        for archivo in carpeta.glob("*.py")
        if not archivo.stem.startswith("_")
    )


def cargar_servicio(nombre: str):
    modulo = importlib.import_module(f"servicios.{nombre}")

    for atributo in ("SERVICIO", "CODIGO", "DIMENSIONES", "VALORES", "completar_paso_2"):
        if not hasattr(modulo, atributo):
            raise AttributeError(f"servicios/{nombre}.py no define {atributo}.")

    if "alcance" not in modulo.DIMENSIONES:
        raise ValueError(f"servicios/{nombre}.py: DIMENSIONES debe incluir 'alcance'.")

    return modulo


def leer_argumentos():
    parser = argparse.ArgumentParser(description="Combinaciones de tarifas SIRT por servicio.")
    parser.add_argument("servicio", nargs="?", help="nombre del archivo en servicios/, sin .py")
    parser.add_argument("--solo-calcular", action="store_true")
    parser.add_argument("--limite", type=int, default=0)
    parser.add_argument("--lote", default=datetime.now().strftime("%Y%m%d"))
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--trace", action="store_true")
    parser.add_argument("--si", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    parser.add_argument("--nombre", help="registra UNA tarifa con este nombre exacto (p. ej. desde el Excel de pruebas)")
    parser.add_argument("--alcance", help="con --nombre: 'Residencial', 'Comercial' o 'Residencial + Comercial'")
    args = parser.parse_args()

    if args.trace:
        args.diagnostico = True

    return args


# =============================================================================
# DIAGNÓSTICO
# =============================================================================

class Diagnostico:
    """
    Con --diagnostico:
    - una captura (pantalla completa) al terminar cada paso del caso,
    - si el caso falla: captura, mapa de la página (aria snapshot), URL, etapa
      y traceback completo.
    Con --trace, además, un trace de Playwright por caso fallido.

    Sin --diagnostico no hace nada (todos los métodos retornan de inmediato).
    """

    def __init__(self, activo: bool, con_trace: bool, carpeta: Path):
        self.activo = activo
        self.con_trace = con_trace
        self.carpeta = carpeta
        self.context = None
        self.carpeta_caso = None
        self.numero_paso = 0

    def iniciar(self, context) -> None:
        if not self.activo:
            return

        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.context = context

        if not self.con_trace:
            return

        try:
            context.tracing.start(screenshots=True, snapshots=True, sources=False)
        except Exception as error:
            print(f"[diagnóstico] No se pudo iniciar el trace ({error}). Se continúa sin trace.")
            self.con_trace = False

    def iniciar_caso(self, caso: dict) -> None:
        if not self.activo:
            return

        self.carpeta_caso = self.carpeta / f"{caso['correlativo']:03d}"
        self.carpeta_caso.mkdir(parents=True, exist_ok=True)
        self.numero_paso = 0

        if self.con_trace:
            self.context.tracing.start_chunk(title=caso["nombre"])

    def captura_paso(self, page, etapa: str) -> None:
        if not self.activo:
            return

        self.numero_paso += 1
        ruta = self.carpeta_caso / f"{self.numero_paso:02d}_{etapa}.png"

        try:
            page.screenshot(path=str(ruta), full_page=True)
        except Exception as error:
            print(f"[diagnóstico] No se pudo capturar '{etapa}': {error}")

    def registrar_fallo(self, page, error: Exception, etapa: str) -> None:
        if not self.activo:
            return

        prefijo = f"{self.numero_paso + 1:02d}_FALLO_{etapa}"

        try:
            page.screenshot(path=str(self.carpeta_caso / f"{prefijo}.png"), full_page=True)
        except Exception as error_captura:
            print(f"[diagnóstico] No se pudo capturar el fallo: {error_captura}")

        # Mapa de la página: lo que Playwright "ve" (roles y nombres accesibles).
        # Sirve para corregir selectores: si un botón o lista no aparece aquí
        # con el nombre esperado, get_by_role no lo va a encontrar.
        try:
            mapa = page.locator("body").aria_snapshot()
            (self.carpeta_caso / f"{prefijo}_mapa.yaml").write_text(mapa, encoding="utf-8")
        except Exception as error_mapa:
            print(f"[diagnóstico] No se pudo obtener el mapa (aria snapshot): {error_mapa}")
            try:
                (self.carpeta_caso / f"{prefijo}_pagina.html").write_text(
                    page.content(), encoding="utf-8"
                )
            except Exception:
                pass

        try:
            url = page.url
        except Exception:
            url = "(no disponible)"

        detalle = (
            f"Fecha:  {datetime.now():%Y-%m-%d %H:%M:%S}\n"
            f"Etapa:  {etapa}\n"
            f"URL:    {url}\n\n"
            f"Error:\n{error}\n\n"
            f"Traceback:\n{''.join(traceback.format_exception(error))}"
        )
        (self.carpeta_caso / f"{prefijo}_error.txt").write_text(detalle, encoding="utf-8")

    def cerrar_caso(self, fallo: bool) -> None:
        if not self.activo or not self.con_trace:
            return

        try:
            if fallo:
                ruta = self.carpeta_caso / "trace.zip"
                self.context.tracing.stop_chunk(path=str(ruta))
                print(f"[diagnóstico] Trace: {ruta}  (playwright show-trace {ruta})")
            else:
                self.context.tracing.stop_chunk()   # caso exitoso: se descarta
        except Exception as error:
            print(f"[diagnóstico] No se pudo guardar el trace: {error}")

    def finalizar(self) -> None:
        if not self.activo:
            return

        if self.con_trace:
            try:
                self.context.tracing.stop()
            except Exception:
                pass

        print(f"Diagnóstico: {self.carpeta.resolve()}")


# =============================================================================
# EJECUCIÓN DE UN CASO
# =============================================================================

def ejecutar_caso(page, servicio, caso: dict, paso1: dict, estado: dict, diag: Diagnostico) -> None:
    valores = caso["valores"]

    def etapa(nombre: str, accion):
        estado["etapa"] = nombre
        resultado = accion()
        diag.captura_paso(page, nombre)
        return resultado

    etapa("abrir_nueva_tarifa", lambda: comun.abrir_nueva_tarifa(page))
    etapa("paso_1_completo", lambda: comun.completar_paso_1(
        page, caso["nombre"], servicio.SERVICIO, valores["alcance"], paso1
    ))
    estado["id_registro"] = etapa("paso_1_guardado", lambda: comun.guardar_paso_1(page))
    etapa("paso_2_completo", lambda: servicio.completar_paso_2(page, valores))
    etapa("paso_2_guardado", lambda: comun.guardar_paso_2(page))
    etapa("confirmada", lambda: comun.confirmar_tarifa(page))


def nombres_exitosos(ruta: Path) -> set[str]:
    if not ruta.exists():
        return set()
    import csv
    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        return {f["nombre"] for f in csv.DictReader(entrada, delimiter=";") if f.get("resultado") == "EXITOSO"}


def siguiente_correlativo(ruta: Path) -> int:
    if not ruta.exists():
        return 1
    import csv
    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        usados = [int(f["correlativo"]) for f in csv.DictReader(entrada, delimiter=";")
                  if (f.get("correlativo") or "").isdigit()]
    return max(usados, default=0) + 1


def main() -> None:
    args = leer_argumentos()
    disponibles = servicios_disponibles()

    if not args.servicio or args.servicio not in disponibles:
        if args.servicio:
            print(f"No existe el servicio '{args.servicio}'.\n")
        print("Servicios disponibles:")
        for nombre in disponibles:
            print(f"  - {nombre}")
        print("\nEjemplo: python ejecutar.py internet_fijo --solo-calcular")
        sys.exit(1)

    servicio = cargar_servicio(args.servicio)
    paso1 = {**comun.PASO1_POR_DEFECTO, **getattr(servicio, "PASO1", {})}
    rutas = comun.rutas_lote(servicio.CODIGO, args.lote)

    print(f"Servicio: {servicio.SERVICIO}")
    print(f"Lote:     {args.lote}")
    if args.diagnostico:
        print(f"Diagnóstico: activado{' (con trace)' if args.trace else ''}")
    print()

    casos = comun.generar_casos(servicio.DIMENSIONES, servicio.CODIGO, args.lote)
    comun.mostrar_calculo(servicio.DIMENSIONES, casos)
    comun.guardar_plan(rutas["plan"], servicio.DIMENSIONES, casos)
    print(f"\nPlan guardado: {rutas['plan'].resolve()}")

    if args.solo_calcular:
        print("\n--solo-calcular: no se creó ninguna tarifa.")
        return

    exitosas_previas = comun.leer_claves_exitosas(rutas["resultados"])
    pendientes = [caso for caso in casos if caso["clave"] not in exitosas_previas]

    if args.nombre:
        # Una sola tarifa con nombre propio: la primera combinación con ese alcance.
        candidatos = casos
        if args.alcance:
            buscado = {a.strip().lower() for a in args.alcance.split("+")}
            candidatos = [c for c in casos if {a.strip().lower() for a in
                          comun.texto_valor(c["valores"].get("alcance", "")).split("+")} == buscado] or casos
        elegido = {**candidatos[0], "nombre": args.nombre}
        ya = nombres_exitosos(rutas["resultados"])
        elegido["correlativo"] = siguiente_correlativo(rutas["resultados"])
        pendientes = [] if args.nombre in ya else [elegido]
        if args.nombre in ya:
            print(f"\n'{args.nombre}' ya está registrada (EXITOSO) en este lote: no se repite.")

    if args.limite > 0:
        pendientes = pendientes[: args.limite]

    print(f"\nYa creados en este lote: {len(exitosas_previas)}")
    print(f"A ejecutar ahora:        {len(pendientes)}")

    if not pendientes:
        print("No hay casos pendientes.")
        return

    # Protección: sin --limite se crearía todo el lote de golpe.
    if args.limite == 0 and len(pendientes) > 1 and not args.si:
        print(f"\nATENCIÓN: sin --limite se crearán {len(pendientes)} tarifas CONFIRMADAS en SIRT.")
        respuesta = input("Escribe 'si' para continuar (Enter para cancelar): ").strip().lower()
        if respuesta not in ("si", "sí"):
            print("Cancelado. Para probar usa --limite 1 o --limite 2.")
            return

    from playwright.sync_api import sync_playwright

    rutas["capturas"].mkdir(parents=True, exist_ok=True)

    diag = Diagnostico(
        activo=args.diagnostico,
        con_trace=args.trace,
        carpeta=rutas["base"] / f"diagnostico_{args.lote}",
    )

    exitosos = 0
    errores = 0

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        diag.iniciar(context)

        page.set_default_timeout(comun.TIMEOUT)
        # Por defecto NO se trae Edge al frente: así se puede seguir usando el
        # equipo (Playwright no necesita la ventana activa). --al-frente para verlo.
        if args.al_frente:
            page.bring_to_front()

        page.get_by_role("button", name="Nueva Tarifa").wait_for(state="visible")
        url_listado = page.url

        for numero, caso in enumerate(pendientes, start=1):
            print("\n" + "=" * 70)
            print(f"[{numero}/{len(pendientes)}] {caso['nombre']}  (caso {caso['correlativo']} del plan)")

            for nombre, valor in caso["valores"].items():
                print(f"  {nombre:<22} {comun.texto_valor(valor)}")

            estado = {"id_registro": "", "etapa": ""}
            fila = {
                "correlativo": caso["correlativo"],
                "nombre": caso["nombre"],
                "clave": caso["clave"],
                "id_registro": "",
                "resultado": "",
                "detalle": "",
                "captura": "",
                "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            diag.iniciar_caso(caso)
            fallo = False

            try:
                ejecutar_caso(page, servicio, caso, paso1, estado, diag)

                fila["resultado"] = "EXITOSO"
                exitosos += 1
                print(f"Resultado: CONFIRMADA (registro {estado['id_registro']})")

            except Exception as error:
                fallo = True
                captura = rutas["capturas"] / f"{caso['correlativo']:03d}_error.png"
                try:
                    page.screenshot(path=str(captura), full_page=True)
                except Exception as error_captura:
                    print(f"No se pudo tomar la captura de error: {error_captura}")

                diag.registrar_fallo(page, error, estado["etapa"])

                fila["resultado"] = "ERROR"
                fila["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
                fila["captura"] = str(captura)
                errores += 1

                print(f"Resultado: ERROR en '{estado['etapa']}' - {error}")
                print(f"Captura: {captura}")

            diag.cerrar_caso(fallo)

            fila["id_registro"] = estado["id_registro"]
            comun.registrar_resultado(rutas["resultados"], fila)

            try:
                comun.volver_al_listado(page, url_listado)
            except Exception as error:
                print(f"No se pudo volver al listado: {error}")
                print("Se detiene el proceso. Vuelve a ejecutar para reanudar.")
                break

            page.wait_for_timeout(PAUSA_ENTRE_CASOS_MS)

        diag.finalizar()

    print("\n" + "=" * 70)
    print("PROCESO FINALIZADO")
    print(f"Confirmadas en esta ejecución: {exitosos}")
    print(f"Errores en esta ejecución:     {errores}")
    print(f"Resultados: {rutas['resultados'].resolve()}")


if __name__ == "__main__":
    main()
