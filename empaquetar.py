"""
Crea una tarifa EMPAQUETADA (Tarifa establecida / Empaquetada) a partir de
tarifas monoproducto que ya tienen registro definitivo.

Flujo (según grabaciones/_empaquetada/):
  Paso 1: Tipo de Tarifa → Condición "Empaquetada" → por cada tarifa:
          "Agregar Tarifa" → "Ingresa el nombre de la tarifa" → Buscar →
          marcar la fila → Agregar.  Luego Datos de la tarifa (nombre, alcance,
          fechas, tarifa de instalación, otros datos) → "Guardar y continuar con el paso".
  Paso 2: una pestaña por servicio ("Servicio 1", "Servicio 2"...): Valor del
          servicio → Guardar (obligatorio en cada una) → "Continuar con el paso".
  Paso 3: Confirmar Tarifa → Confirmar tarifa → Aceptar.

Reglas de SIRT (vistas en pruebas):
- El alcance del paquete debe ser igual al de cada tarifa que lo forma.
- El "Valor del plan" del paquete es la suma de los valores del servicio y debe
  cumplir:  valor de la tarifa más cara  <=  suma  <  suma de las tarifas base.
  (Con 30 + 50 → 80 SIRT rechaza: no es menor que 80. Con 29 + 49 → 78 acepta.
  Con 3 + 5 → 8 rechaza: es menor que 50.)
  Por defecto cada servicio lleva el valor de su tarifa menos 1 (29 y 49).
  Ese chequeo se hace también antes de empezar.
- El orden de las pestañas no es el orden en que se agregan las tarifas:
  cada pestaña se identifica por el nombre del servicio que muestra.

Uso:
    python empaquetar.py --tarifas TE-TFP-MUESTRA-001 TE-SMP-MUESTRA-001 --lote MUESTRA --solo-listar
    python empaquetar.py --tarifas TE-TFP-MUESTRA-001 TE-SMP-MUESTRA-001 --lote MUESTRA --diagnostico

Opciones:
    --tarifas N1 N2 ...  nombres de las tarifas a empaquetar (2 o más)
    --lote TEXTO         lote del paquete (el nombre será TE-EMP-<lote>-NNN)
    --nombre TEXTO       nombre del paquete (en vez del automático)
    --valores V1 V2 ...  valor del servicio de cada tarifa, en el mismo orden que --tarifas
                         (por defecto: el valor de cada tarifa menos 1)
    --solo-listar        muestra lo que haría, sin tocar SIRT
    --diagnostico        capturas por paso y mapa de la página si falla
    --al-frente          trae la ventana de Edge al frente al empezar

Resultados en resultados/EMP/resultados_<lote>.csv (mismo formato que ejecutar.py,
así registro_definitivo.py puede usarlos después).

Requisitos: Edge abierto en modo controlable y sesión iniciada en SIRT
(en cualquier pantalla: el script entra solo a "Nueva tarifa").
"""

import argparse
import csv
import importlib
import re
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import comun
from ejecutar import Diagnostico


CODIGO_PAQUETE = "EMP"
CONDICION_EMPAQUETADA = "Empaquetada"
TIMEOUT_BUSQUEDA = 20_000


# --- Datos de las tarifas base ------------------------------------------------

def valor_por_defecto(valor_tarifa):
    """Valor de la tarifa menos 1 (o el 90 % si vale 1 o menos), para que el
    paquete salga más barato que la suma de sus tarifas."""
    valor = Decimal(str(valor_tarifa))
    nuevo = valor - 1 if valor > 1 else (valor * Decimal("0.9")).quantize(Decimal("0.01"))
    return format(nuevo.normalize(), "f")


def servicios_por_codigo() -> dict:
    """CODIGO → módulo del servicio (para saber su nombre en SIRT y su valor)."""
    modulos = {}
    for ruta in sorted((Path(__file__).parent / "servicios").glob("*.py")):
        if ruta.stem.startswith("_"):
            continue
        modulo = importlib.import_module(f"servicios.{ruta.stem}")
        modulos[modulo.CODIGO] = modulo
    return modulos


def buscar_tarifa_base(nombre: str) -> dict | None:
    """Busca la tarifa en los resultados de ejecutar.py (cualquier servicio y lote)."""
    for ruta in sorted(comun.DIRECTORIO_RESULTADOS.glob("*/resultados_*.csv")):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            for fila in csv.DictReader(entrada, delimiter=";"):
                if fila.get("nombre") == nombre and fila.get("resultado") == "EXITOSO":
                    return {
                        "nombre": nombre,
                        "codigo": ruta.parent.name,
                        # La clave empieza por el alcance: "Residencial | ..." o "Residencial + Comercial | ..."
                        "alcance": tuple(a.strip() for a in fila["clave"].split(" | ")[0].split(" + ")),
                        "id_registro": fila.get("id_registro", ""),
                    }
    return None


def tiene_registro_definitivo(nombre: str) -> bool:
    for ruta in comun.DIRECTORIO_RESULTADOS.glob("registro_definitivo_*.csv"):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            for fila in csv.DictReader(entrada, delimiter=";"):
                if fila.get("nombre") == nombre and fila.get("resultado") == "EXITOSO":
                    return True
    return False


def siguiente_nombre(lote: str, ruta_resultados: Path) -> tuple[int, str]:
    correlativo = 1
    if ruta_resultados.exists():
        with ruta_resultados.open("r", newline="", encoding="utf-8-sig") as entrada:
            usados = [int(f["correlativo"]) for f in csv.DictReader(entrada, delimiter=";") if f.get("correlativo")]
        correlativo = max(usados, default=0) + 1
    return correlativo, f"TE-{CODIGO_PAQUETE}-{lote}-{correlativo:03d}"


# --- SIRT ---------------------------------------------------------------------

def abrir_nueva_tarifa(page) -> None:
    origen = re.match(r"https?://[^/]+", page.url).group(0)
    page.keyboard.press("Escape")
    page.goto(f"{origen}/extranet/nueva-tarifa", wait_until="domcontentloaded")
    page.get_by_role("combobox", name="Tipo de Tarifa").wait_for(state="visible")


def texto_dialogo(page) -> str:
    dialogo = page.locator("mat-dialog-container, [role='dialog']")
    return dialogo.last.inner_text().replace("\n", " ").strip() if dialogo.count() else ""


def texto_visible(page, texto: str) -> bool:
    """¿Hay en pantalla un elemento visible cuyo texto sea exactamente `texto`?"""
    return page.evaluate(
        """
        texto => Array.from(document.querySelectorAll("body *")).some(el =>
            el.childElementCount === 0 &&
            el.textContent.trim() === texto &&
            el.offsetParent !== null)
        """,
        texto,
    )


def total_listado(page) -> int:
    """Número que muestra 'Listado de tarifas: Total N'."""
    texto = page.get_by_text(re.compile(r"Listado de tarifas:\s*Total\s*\d+")).first.inner_text()
    return int(re.search(r"Total\s*(\d+)", texto).group(1))


def limpiar_listado(page) -> int:
    """SIRT recuerda las tarifas agregadas en un intento que no se guardó: al
    volver a 'Nueva tarifa' siguen en el listado. Se quitan todas antes de
    empezar. Devuelve cuántas quitó."""
    quitadas = 0
    for _ in range(100):
        quitar = page.get_by_role("button", name="Quitar Tarifa")
        if quitar.count() == 0:
            return quitadas
        antes = total_listado(page)
        quitar.first.click()
        for _ in range(40):                     # hasta ~10 s
            page.wait_for_timeout(250)
            dialogo = page.locator("mat-dialog-container")
            if dialogo.count():
                for texto in ("Aceptar", "Sí", "Si", "Confirmar"):
                    boton = dialogo.last.get_by_role("button", name=texto, exact=True)
                    if boton.count():
                        boton.click()
                        break
            if total_listado(page) < antes:
                break
        else:
            raise RuntimeError("No se pudo quitar una tarifa que quedó de un intento anterior.")
        quitadas += 1
    return quitadas


def agregar_tarifa_al_paquete(page, nombre: str, numero: int) -> None:
    """'Agregar Tarifa' → buscar por nombre → marcar su fila → Agregar."""
    page.get_by_role("button", name="Agregar Tarifa").click()

    campo = page.get_by_role("textbox", name="Ingresa el nombre de la tarifa")
    campo.wait_for(state="visible")
    campo.fill(nombre)
    page.get_by_role("button", name="Buscar", exact=True).click()

    # Fila con ese nombre (sin espacios: en pantalla se parte) y con casilla:
    # las filas del listado del paquete no tienen casilla.
    buscar_fila = """
        nombre => Array.from(document.querySelectorAll("tbody tr")).findIndex(fila =>
            fila.querySelector("input[type=checkbox]") &&
            Array.from(fila.querySelectorAll("td"))
                .some(c => c.innerText.replace(/\\s+/g, "") === nombre))
    """
    try:
        page.wait_for_function(f"nombre => ({buscar_fila})(nombre) >= 0", arg=nombre, timeout=TIMEOUT_BUSQUEDA)
    except Exception:
        raise RuntimeError(
            f"'{nombre}' no aparece en 'Agregar tarifas' (¿tiene registro definitivo?)."
        )

    indice = page.evaluate(f"nombre => ({buscar_fila})(nombre)", nombre)
    page.locator("tbody tr").nth(indice).get_by_role("checkbox").check()
    page.get_by_role("button", name="Agregar", exact=True).click()
    campo.wait_for(state="hidden")

    # El listado del paquete debe mostrar ya `numero` tarifas.
    page.get_by_text(re.compile(rf"Listado de tarifas.*Total\s*{numero}\b")).first.wait_for(state="visible")


PASO1_EXTRA: dict = {}   # fechas, instalación, otros datos y contacto pedidos por línea de comandos


def completar_paso_1(page, nombre_paquete: str, componentes: list[dict], alcance: tuple) -> str:
    comun.elegir_opcion(page, "Tipo de Tarifa", comun.TIPO_TARIFA)
    comun.elegir_opcion(page, "Condición", CONDICION_EMPAQUETADA)
    limpiar_listado(page)

    for numero, componente in enumerate(componentes, start=1):
        agregar_tarifa_al_paquete(page, componente["nombre"], numero)

    paso1 = {**comun.PASO1_POR_DEFECTO, **{k: v for k, v in PASO1_EXTRA.items() if k != "contacto"}}
    hoy = date.today()

    page.get_by_role("textbox", name="Nombre de la tarifa").fill(nombre_paquete)
    comun.elegir_varias_opciones(page, "Alcance de la tarifa", alcance)
    comun.elegir_fecha(page, "Inicio de vigencia", paso1["fecha_inicio_vigencia"] or hoy)
    comun.elegir_fecha(page, "Inicio comercialización", paso1["fecha_inicio_comercializacion"] or hoy)
    comun.llenar_si_editable(page.get_by_role("spinbutton", name="Tarifa Instalación"), paso1["tarifa_instalacion"])
    page.get_by_role("textbox", name="Otros datos generales").fill(paso1["otros_datos"])
    for aviso in comun.llenar_contacto(page, PASO1_EXTRA.get("contacto", {})):
        print(f"  Aviso: {aviso}")

    return comun.guardar_paso_1(page)


def completar_paso_2(page, componentes: list[dict]) -> None:
    """Una pestaña por servicio: valor + Guardar. Cada pestaña se identifica
    por el nombre del servicio que muestra (no por el orden de agregado)."""
    pendientes = list(componentes)

    for numero in range(1, len(componentes) + 1):
        pestana = page.get_by_role("tab", name=f"Servicio {numero}", exact=True)
        pestana.wait_for(state="visible")
        pestana.click()
        page.wait_for_timeout(500)

        componente = next((c for c in pendientes if texto_visible(page, c["servicio"])), None)
        if componente is None:
            raise RuntimeError(
                f"La pestaña 'Servicio {numero}' no muestra ninguno de los servicios esperados: "
                + ", ".join(c["servicio"] for c in pendientes)
            )
        pendientes.remove(componente)

        page.get_by_role("spinbutton", name="Valor del servicio").fill(componente["valor"])
        page.get_by_role("button", name="Guardar", exact=True).click()
        page.wait_for_timeout(1_000)

    page.get_by_role("button", name="Continuar con el paso").click()

    confirmar = page.get_by_role("button", name="Confirmar Tarifa", exact=True)
    try:
        confirmar.wait_for(state="visible", timeout=15_000)
    except Exception:
        raise RuntimeError(f"No pasó al resumen. Mensaje de SIRT: {texto_dialogo(page)[:300]}")


def verificar_resumen(page, componentes: list[dict]) -> None:
    """En el resumen, cada servicio debe mostrar el valor que se le asignó
    (comprueba que ninguna pestaña recibió el valor de otro servicio)."""
    for componente in componentes:
        bloque = page.get_by_role("region", name=f"Servicio {componente['servicio']}", exact=True)
        if bloque.count() == 0:
            print(f"  Aviso: el resumen no muestra el bloque de '{componente['servicio']}'; no se verificó su valor.")
            continue
        encontrado = re.search(r"Nuevo valor del servicio[^:]*:\s*([\d.,]+)", bloque.first.inner_text())
        if not encontrado:
            print(f"  Aviso: no se leyó el valor de '{componente['servicio']}' en el resumen.")
            continue
        if Decimal(encontrado.group(1).replace(",", "")) != Decimal(componente["valor"]):
            raise RuntimeError(
                f"En el resumen, '{componente['servicio']}' tiene {encontrado.group(1)} "
                f"y se esperaba {componente['valor']}: no se confirmó."
            )


def confirmar_paquete(page) -> None:
    """Como comun.confirmar_tarifa, pero detecta el aviso de validaciones
    ('No se cumplieron las validaciones...') que SIRT muestra al confirmar."""
    page.get_by_role("button", name="Confirmar Tarifa", exact=True).click()

    boton_confirmar = page.get_by_role("button", name="Confirmar tarifa", exact=True)
    boton_confirmar.wait_for(state="visible")
    boton_confirmar.click()

    boton_aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    boton_aceptar.wait_for(state="visible")
    mensaje = texto_dialogo(page)

    boton_aceptar.click()
    boton_aceptar.wait_for(state="hidden")

    if "no se cumplieron" in mensaje.lower() or "validaci" in mensaje.lower():
        raise RuntimeError(f"SIRT rechazó el paquete: {mensaje[:400]}")


# --- Principal ----------------------------------------------------------------

def leer_argumentos():
    parser = argparse.ArgumentParser(description="Crea una tarifa empaquetada.")
    parser.add_argument("--tarifas", nargs="+", required=True)
    parser.add_argument("--lote", required=True)
    parser.add_argument("--nombre")
    parser.add_argument("--valores", nargs="+")
    parser.add_argument("--alcance", help="alcance pedido (se comprueba que coincida con el de las tarifas)")
    parser.add_argument("--inicio-vigencia", help="AAAA-MM-DD (por defecto hoy)")
    parser.add_argument("--inicio-comercializacion", help="AAAA-MM-DD (por defecto hoy)")
    parser.add_argument("--instalacion", help="tarifa de instalación")
    parser.add_argument("--otros-datos", help="texto de «Otros datos generales»")
    parser.add_argument("--enlace")
    parser.add_argument("--fono")
    parser.add_argument("--email")
    parser.add_argument("--solo-listar", action="store_true")
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    return parser.parse_args()


def preparar_componentes(args) -> tuple[list[dict], tuple]:
    if len(args.tarifas) < 2:
        sys.exit("Un paquete necesita al menos 2 tarifas.")
    if args.valores and len(args.valores) != len(args.tarifas):
        sys.exit("--valores debe tener un valor por cada tarifa, en el mismo orden.")

    modulos = servicios_por_codigo()
    componentes = []
    problemas = []

    for indice, nombre in enumerate(args.tarifas):
        base = buscar_tarifa_base(nombre)
        if base is None:
            problemas.append(f"'{nombre}': no está como EXITOSO en resultados/*/resultados_*.csv")
            continue
        if not tiene_registro_definitivo(nombre):
            problemas.append(f"'{nombre}': no tiene registro definitivo (registro_definitivo.py)")
        modulo = modulos.get(base["codigo"])
        if modulo is None:
            problemas.append(f"'{nombre}': servicio de código {base['codigo']} desconocido")
            continue

        base["servicio"] = modulo.SERVICIO
        # Valor con que se registró (Excel) o, si no se anotó, el valor fijo del código.
        base["valor_tarifa"] = comun.valor_tarifa_registrado(nombre) or modulo.VALORES.get("valor_servicio", "")
        if args.valores:
            base["valor"] = args.valores[indice]
        else:
            base["valor"] = valor_por_defecto(base["valor_tarifa"])
        componentes.append(base)

    if problemas:
        sys.exit("No se puede armar el paquete:\n  - " + "\n  - ".join(problemas))

    alcances = {c["alcance"] for c in componentes}
    if len(alcances) != 1:
        sys.exit("Las tarifas tienen alcances distintos (SIRT lo rechaza): "
                 + "; ".join(f"{c['nombre']}: {' + '.join(c['alcance'])}" for c in componentes))
    if args.alcance:
        pedido = {a.strip().lower() for a in args.alcance.split("+")}
        propio = {a.lower() for a in next(iter(alcances))}
        if pedido != propio:
            sys.exit(f"El alcance pedido ({args.alcance}) no coincide con el de las tarifas "
                     f"({' + '.join(next(iter(alcances)))}): SIRT lo rechaza.")

    servicios = [c["servicio"] for c in componentes]
    if len(set(servicios)) != len(servicios):
        sys.exit("Hay dos tarifas del mismo servicio: las pestañas del paso 2 no se podrían distinguir.")

    # Reglas de SIRT: valor de la más cara <= suma de valores < suma de las tarifas base.
    suma = sum(Decimal(c["valor"]) for c in componentes)
    mayor = max(Decimal(c["valor_tarifa"]) for c in componentes)
    total_base = sum(Decimal(c["valor_tarifa"]) for c in componentes)
    if suma < mayor:
        sys.exit(f"La suma de valores ({suma}) es menor que el valor de la tarifa más cara ({mayor}): SIRT lo rechaza.")
    if suma >= total_base:
        sys.exit(f"La suma de valores ({suma}) debe ser MENOR que la suma de las tarifas base ({total_base}): SIRT lo rechaza.")

    return componentes, alcances.pop()


def main() -> None:
    args = leer_argumentos()
    if args.inicio_vigencia:
        PASO1_EXTRA["fecha_inicio_vigencia"] = date.fromisoformat(args.inicio_vigencia)
    if args.inicio_comercializacion:
        PASO1_EXTRA["fecha_inicio_comercializacion"] = date.fromisoformat(args.inicio_comercializacion)
    if args.instalacion:
        PASO1_EXTRA["tarifa_instalacion"] = args.instalacion
    if args.otros_datos:
        PASO1_EXTRA["otros_datos"] = args.otros_datos
    PASO1_EXTRA["contacto"] = {"enlace": args.enlace, "fono": args.fono, "email": args.email}
    componentes, alcance = preparar_componentes(args)

    rutas = comun.rutas_lote(CODIGO_PAQUETE, args.lote)
    correlativo, nombre_auto = siguiente_nombre(args.lote, rutas["resultados"])
    nombre_paquete = args.nombre or nombre_auto
    clave = " | ".join([" + ".join(alcance)] + [c["nombre"] for c in componentes])

    print(f"Paquete: {nombre_paquete}")
    print(f"Alcance: {' + '.join(alcance)}")
    for c in componentes:
        print(f"  - {c['nombre']:<22} {c['servicio']:<40} valor {c['valor']} (tarifa: {c['valor_tarifa']})")
    print(f"Suma de valores: {sum(Decimal(c['valor']) for c in componentes)}")

    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT.")
        return

    from playwright.sync_api import sync_playwright

    diag = Diagnostico(args.diagnostico, False, rutas["base"] / f"diagnostico_{args.lote}")
    estado = {"etapa": "", "id_registro": ""}
    fila = {
        "correlativo": correlativo, "nombre": nombre_paquete, "clave": clave,
        "id_registro": "", "resultado": "", "detalle": "", "captura": "",
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        diag.iniciar(context)
        diag.iniciar_caso({"correlativo": correlativo, "nombre": nombre_paquete})
        if args.al_frente:
            page.bring_to_front()

        def etapa(nombre, accion):
            estado["etapa"] = nombre
            resultado = accion()
            diag.captura_paso(page, nombre)
            return resultado

        try:
            etapa("abrir_nueva_tarifa", lambda: abrir_nueva_tarifa(page))
            estado["id_registro"] = etapa(
                "paso_1_guardado", lambda: completar_paso_1(page, nombre_paquete, componentes, alcance)
            )
            etapa("paso_2_completo", lambda: completar_paso_2(page, componentes))
            etapa("resumen_verificado", lambda: verificar_resumen(page, componentes))
            etapa("confirmada", lambda: confirmar_paquete(page))

            fila["resultado"] = "EXITOSO"
            print(f"\nResultado: PAQUETE CONFIRMADO (registro {estado['id_registro']})")

        except Exception as error:
            rutas["capturas"].mkdir(parents=True, exist_ok=True)
            captura = rutas["capturas"] / f"{correlativo:03d}_error.png"
            try:
                page.screenshot(path=str(captura), full_page=True)
            except Exception:
                pass
            diag.registrar_fallo(page, error, estado["etapa"])
            fila["resultado"] = "ERROR"
            fila["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
            fila["captura"] = str(captura)
            print(f"\nResultado: ERROR en '{estado['etapa']}' - {error}")

        diag.cerrar_caso(fila["resultado"] == "ERROR")
        diag.finalizar()

    fila["id_registro"] = estado["id_registro"]
    comun.registrar_resultado(rutas["resultados"], fila)
    print(f"Resultados: {rutas['resultados'].resolve()}")


if __name__ == "__main__":
    main()
