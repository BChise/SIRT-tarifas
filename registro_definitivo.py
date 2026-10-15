"""
Registro definitivo de las tarifas creadas por ejecutar.py.

Toma las tarifas EXITOSAS de un lote (resultados/<CODIGO>/resultados_<lote>.csv),
las busca en "Tarifas temporales" por su nombre y, una por una:
  1. si el servicio lo requiere, asigna la disponibilidad comercial
     (Acciones → "Ver Disponibilidad Comercial" → ubicaciones → Guardar → Aceptar);
  2. hace el registro definitivo
     ("Registro definitivo" → "Registrar tarifa" → Aceptar).

Uso:
    python registro_definitivo.py --lote MUESTRA --solo-listar      # solo muestra qué haría
    python registro_definitivo.py --lote MUESTRA --limite 1 --diagnostico
    python registro_definitivo.py --lote MUESTRA --servicio TVP --limite 1
    python registro_definitivo.py --lote MUESTRA                    # todas (pide confirmación)

Opciones:
    --lote TEXTO       lote de ejecutar.py cuyas tarifas se registran (obligatorio)
    --servicio CODIGO  solo ese servicio (código: TF, TVP, SMP...); se puede repetir
    --nombre NOMBRE    solo esa tarifa (p. ej. TE-EMP-MUESTRA-005); se puede repetir
    --limite N         procesa solo las primeras N pendientes
    --ubigeos LISTA    ubicaciones separadas por coma (por defecto: Amazonas).
                       En empaquetadas, por defecto: la intersección de las
                       ubicaciones de sus tarifas que requieren disponibilidad.
    --solo-listar      muestra las tarifas y si llevan disponibilidad; no toca SIRT
    --diagnostico      capturas por paso y mapa de la página si falla
    --si               no pide confirmación
    --al-frente        trae la ventana de Edge al frente al empezar

Es reanudable: lo ya registrado (EXITOSO en resultados/registro_definitivo_<lote>.csv)
se omite en la siguiente ejecución.

Requisitos: Edge abierto en modo controlable y sesión iniciada en SIRT.

IMPORTANTE: el registro definitivo probablemente no se puede deshacer.
Probar primero con --limite 1.
"""

import argparse
import csv
import re
import sys
from datetime import datetime
from pathlib import Path

import comun
from ejecutar import Diagnostico


# --- Configuración -----------------------------------------------------------

UBIGEOS_POR_DEFECTO = ["Amazonas"]

# Servicios que requieren disponibilidad comercial antes del registro definitivo.
# Según flagDisponibilidadComercial de cada paso2.json (TF e IF por patrón: no
# hay JSON). Confirmado en pruebas: (pendiente).
REQUIERE_DISPONIBILIDAD = {
    "TF": True,     # Telefonía fija postpago (por patrón)
    "TFP": True,    # Telefonía fija prepago
    "OTF": True,    # Otras tarifas de telefonía fija
    "PTF": True,    # Paquetes adicionales de telefonía fija
    "IF": True,     # Internet fijo postpago (por patrón; la grabación lo confirma)
    "PIF": True,    # Paquetes adicionales y otros servicios de internet fijo
    "TVP": True,    # Televisión de paga Postpago
    "TVPR": True,   # Televisión de paga Prepago
    "PTV": True,    # Paquetes adicionales y otras tarifas de televisión de paga
    "SMP": False,   # Servicio público móvil postpago
    "SMC": False,   # Servicio público móvil control
    "SMQ": False,   # Servicio público móvil paquete
    "SMO": False,   # Servicio móvil otros
    "AC": False,    # Arrendamiento de circuitos
    "OS": False,    # Otros servicios
}

# Empaquetadas (EMP): lo deciden sus tarifas. Confirmado el 14/10: TFP + SMP
# pidió disponibilidad; SMP + OS no. Basta que una de sus tarifas la requiera.
CODIGO_PAQUETE = "EMP"


def requiere_disponibilidad(tarifa: dict) -> bool:
    codigo = tarifa["codigo_servicio"]
    if codigo == CODIGO_PAQUETE:
        # clave = "alcance | TE-TFP-MUESTRA-001 | TE-SMP-MUESTRA-001"
        componentes = [n.strip() for n in tarifa.get("clave", "").split(" | ")[1:]]
        codigos = [n.split("-")[1] for n in componentes if n.count("-") >= 2]
        if not codigos:
            sys.exit(f"No se pudo saber qué tarifas forman {tarifa['nombre']} (columna clave).")
        return any(requiere_disponibilidad({"codigo_servicio": c}) for c in codigos)
    if codigo not in REQUIERE_DISPONIBILIDAD:
        print(f"Código de servicio desconocido: {codigo}. Agrégalo a REQUIERE_DISPONIBILIDAD.")
        sys.exit(1)
    return REQUIERE_DISPONIBILIDAD[codigo]


def disponibilidad_registrada() -> dict[str, list[str]]:
    """Nombre → ubicaciones con que se registró (de todos los registro_definitivo_*.csv)."""
    resultado = {}
    for ruta in sorted(comun.DIRECTORIO_RESULTADOS.glob("registro_definitivo_*.csv")):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            for fila in csv.DictReader(entrada, delimiter=";"):
                if fila.get("resultado") == "EXITOSO":
                    resultado[fila["nombre"]] = [u.strip() for u in fila.get("disponibilidad", "").split(",") if u.strip()]
    return resultado


def ubigeos_del_paquete(tarifa: dict, pedidos: list[str] | None,
                        registradas: dict[str, list[str]]) -> list[str]:
    """Regla (14/10): la disponibilidad del paquete debe estar dentro de la
    intersección de las ubicaciones de sus tarifas que requieren disponibilidad.
    Las que no la requieren no influyen. Lanza ValueError si no se puede."""
    componentes = [n.strip() for n in tarifa.get("clave", "").split(" | ")[1:]]
    comunes = None
    for nombre in componentes:
        codigo = nombre.split("-")[1] if nombre.count("-") >= 2 else ""
        if not REQUIERE_DISPONIBILIDAD.get(codigo, False):
            continue
        ubicaciones = registradas.get(nombre)
        if not ubicaciones:
            raise ValueError(f"no se sabe con qué ubicaciones se registró {nombre}")
        comunes = set(ubicaciones) if comunes is None else comunes & set(ubicaciones)

    if not comunes:
        raise ValueError("sus tarifas no tienen ninguna ubicación en común")
    if pedidos is None:
        return sorted(comunes)
    fuera = [u for u in pedidos if u not in comunes]
    if fuera:
        raise ValueError(f"{', '.join(fuera)} no está en todas sus tarifas (en común: {', '.join(sorted(comunes))})")
    return pedidos


PAUSA_ENTRE_TARIFAS_MS = 1_000
TIMEOUT_BUSQUEDA = 20_000

CAMPOS_REGISTRO = [
    "nombre", "codigo_servicio", "id_registro", "disponibilidad",
    "resultado", "detalle", "fecha_hora",
]


# --- Lectura de tarifas a registrar -------------------------------------------

def leer_tarifas_del_lote(lote: str, servicios: list[str] | None) -> list[dict]:
    """Tarifas EXITOSAS del lote, de todos los servicios (o de los indicados)."""
    tarifas = []

    for ruta in sorted(comun.DIRECTORIO_RESULTADOS.glob(f"*/resultados_{lote}.csv")):
        codigo = ruta.parent.name

        if servicios and codigo not in servicios:
            continue

        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            for fila in csv.DictReader(entrada, delimiter=";"):
                if fila.get("resultado") != "EXITOSO":
                    continue

                tarifas.append({
                    "nombre": fila["nombre"],
                    "codigo_servicio": codigo,
                    "id_registro": fila.get("id_registro", ""),
                    "clave": fila.get("clave", ""),
                })

    # Un mismo nombre puede repetirse si un caso se reintentó: se deja uno.
    vistos = set()
    unicas = []
    for tarifa in tarifas:
        if tarifa["nombre"] not in vistos:
            vistos.add(tarifa["nombre"])
            unicas.append(tarifa)

    return unicas


def leer_registradas(ruta: Path) -> set[str]:
    if not ruta.exists():
        return set()

    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        return {
            fila["nombre"]
            for fila in csv.DictReader(entrada, delimiter=";")
            if fila.get("resultado") == "EXITOSO"
        }


def registrar(ruta: Path, fila: dict) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    nuevo = not ruta.exists()

    with ruta.open("a", newline="", encoding="utf-8-sig") as salida:
        escritor = csv.DictWriter(salida, fieldnames=CAMPOS_REGISTRO, delimiter=";")
        if nuevo:
            escritor.writeheader()
        escritor.writerow(fila)


# --- SIRT: Tarifas temporales -------------------------------------------------

def recargar_temporales(page) -> None:
    """
    Carga de nuevo "Tarifas temporales". Se hace al empezar cada tarifa para
    partir sin ninguna fila marcada: una marca que quedara de una tarifa
    anterior (oculta por el filtro) podría registrarse sin querer.
    """
    origen = re.match(r"https?://[^/]+", page.url).group(0)
    page.keyboard.press("Escape")
    page.goto(f"{origen}/extranet/tarifa-temporal", wait_until="domcontentloaded")
    page.get_by_role("button", name="Registro definitivo").wait_for(state="visible")


def abrir_tarifas_temporales(page) -> None:
    campo_nombre = page.get_by_role("textbox", name="Ingresa nombre de tarifa")
    if campo_nombre.count() > 0 and campo_nombre.is_visible():
        return

    if "tarifa-temporal" not in page.url:
        recargar_temporales(page)

    page.get_by_role("button", name="Registro definitivo").wait_for(state="visible")

    # El filtro por nombre está en "Filtro avanzado" (abierto, el botón pasa
    # a llamarse "Filtro principal").
    if not campo_nombre.is_visible():
        page.get_by_role("button", name="Filtro avanzado").click()
        campo_nombre.wait_for(state="visible")


def fila_de_tarifa(page, nombre: str):
    """
    Devuelve la fila de la tabla con ese nombre y en estado confirmado, o None.
    Compara el nombre sin espacios (en pantalla se parte en varias líneas) y
    descarta borradores ("En proceso - Paso 1"), que pueden tener el mismo nombre.
    """
    indice = page.locator("tbody tr").evaluate_all(
        """
        (filas, nombre) => filas.findIndex(fila => {
            const celdas = Array.from(fila.querySelectorAll("td"))
                .map(c => c.innerText.replace(/\\s+/g, ""));
            const texto = fila.innerText.toLowerCase();
            return celdas.includes(nombre) && texto.includes("confirmado");
        })
        """,
        nombre,
    )
    return None if indice < 0 else page.locator("tbody tr").nth(indice)


def buscar_tarifa(page, nombre: str):
    abrir_tarifas_temporales(page)

    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill(nombre)
    page.get_by_role("button", name="Buscar", exact=True).click()

    try:
        page.wait_for_function(
            """
            nombre => Array.from(document.querySelectorAll("tbody tr")).some(fila =>
                Array.from(fila.querySelectorAll("td"))
                    .some(c => c.innerText.replace(/\\s+/g, "") === nombre))
            """,
            arg=nombre,
            timeout=TIMEOUT_BUSQUEDA,
        )
    except Exception:
        raise RuntimeError(
            f"'{nombre}' no aparece en Tarifas temporales "
            "(¿ya registrada, o el filtro no la encuentra?)."
        )

    page.wait_for_timeout(500)
    fila = fila_de_tarifa(page, nombre)

    if fila is None:
        raise RuntimeError(f"'{nombre}' aparece, pero no en estado confirmado (¿borrador?).")

    return fila


def aceptar_dialogos(page, maximo: int = 2) -> list[str]:
    """
    Pulsa "Aceptar" en los diálogos que vayan apareciendo (hasta `maximo`) y
    devuelve sus textos, para revisar resúmenes como "Fallidos: 0".
    """
    textos = []

    for numero in range(maximo):
        boton = page.get_by_role("button", name="Aceptar", exact=True)
        try:
            boton.first.wait_for(state="visible", timeout=comun.TIMEOUT if numero == 0 else 5_000)
        except Exception:
            if numero == 0:
                raise
            break

        dialogo = page.locator("mat-dialog-container, [role='dialog']")
        if dialogo.count() > 0:
            textos.append(dialogo.last.inner_text().replace("\n", " ").strip())

        boton.first.click()
        boton.first.wait_for(state="hidden", timeout=10_000)
        page.wait_for_timeout(500)

    return textos


def revisar_fallidos(textos: list[str], operacion: str) -> None:
    for texto in textos:
        coincidencia = re.search(r"Fallidos:\s*(\d+)", texto)
        if coincidencia and int(coincidencia.group(1)) > 0:
            raise RuntimeError(f"{operacion}: SIRT informa {coincidencia.group(1)} fallido(s). {texto[:200]}")


def asignar_disponibilidad(page, fila, ubigeos: list[str]) -> list[str]:
    fila.get_by_role("checkbox").check()
    fila.get_by_label("Acciones").click()
    page.get_by_role("menuitem", name="Ver Disponibilidad Comercial").click()

    primera = page.get_by_role("checkbox", name=ubigeos[0], exact=True)
    primera.wait_for(state="visible")

    for ubigeo in ubigeos:
        page.get_by_role("checkbox", name=ubigeo, exact=True).check()

    page.get_by_role("button", name="Guardar", exact=True).click()

    textos = aceptar_dialogos(page, maximo=2)
    revisar_fallidos(textos, "Disponibilidad comercial")
    return textos


def cantidad_seleccionada(texto: str) -> int | None:
    """Lee "Tarifas seleccionadas: N" (o parecido) de un diálogo, si lo dice."""
    coincidencia = re.search(r"seleccionad[ao]s?\s*:?\s*(\d+)", texto, re.IGNORECASE)
    return int(coincidencia.group(1)) if coincidencia else None


def registro_definitivo(page, fila) -> list[str]:
    fila.get_by_role("checkbox").check()

    marcadas = page.locator("tbody tr").get_by_role("checkbox", checked=True).count()
    if marcadas != 1:
        raise RuntimeError(f"Hay {marcadas} filas marcadas; se esperaba 1. No se registra.")

    page.get_by_role("button", name="Registro definitivo").click()

    boton_registrar = page.get_by_role("button", name="Registrar tarifa")
    boton_registrar.wait_for(state="visible")

    # Si la confirmación dice cuántas tarifas se van a registrar, debe ser 1.
    dialogo = page.locator("mat-dialog-container, [role='dialog']")
    texto_confirmacion = dialogo.last.inner_text().replace("\n", " ") if dialogo.count() else ""
    cantidad = cantidad_seleccionada(texto_confirmacion)
    if cantidad is not None and cantidad != 1:
        page.keyboard.press("Escape")
        raise RuntimeError(f"La confirmación indica {cantidad} tarifas; se esperaba 1. Cancelado.")

    boton_registrar.click()

    textos = aceptar_dialogos(page, maximo=2)
    revisar_fallidos(textos, "Registro definitivo")
    return textos


def sigue_en_temporales(page, nombre: str) -> bool:
    """Tras el registro, la tarifa ya no debería figurar como confirmada en temporales."""
    abrir_tarifas_temporales(page)
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill(nombre)
    page.get_by_role("button", name="Buscar", exact=True).click()
    page.wait_for_timeout(2_500)
    return fila_de_tarifa(page, nombre) is not None


def procesar_tarifa(page, tarifa: dict, ubigeos: list[str], estado: dict, diag: Diagnostico) -> str:
    def etapa(nombre_etapa: str, accion):
        estado["etapa"] = nombre_etapa
        resultado = accion()
        diag.captura_paso(page, nombre_etapa)
        return resultado

    nombre = tarifa["nombre"]
    detalles = []

    estado["etapa"] = "recargar"
    recargar_temporales(page)

    fila = etapa("buscar", lambda: buscar_tarifa(page, nombre))

    if tarifa["requiere_disponibilidad"]:
        textos = etapa("disponibilidad", lambda: asignar_disponibilidad(page, fila, ubigeos))
        detalles.append("disp: " + " / ".join(t[:120] for t in textos))
        # Al volver, se busca otra vez para tener la fila actualizada.
        fila = etapa("buscar_de_nuevo", lambda: buscar_tarifa(page, nombre))

    textos = etapa("registro_definitivo", lambda: registro_definitivo(page, fila))
    detalles.append("registro: " + " / ".join(t[:120] for t in textos))

    if etapa("verificacion", lambda: sigue_en_temporales(page, nombre)):
        detalles.append("AVISO: sigue figurando en Tarifas temporales")

    return " | ".join(detalles)


# --- Principal ----------------------------------------------------------------

def leer_argumentos():
    parser = argparse.ArgumentParser(description="Registro definitivo de tarifas de un lote.")
    parser.add_argument("--lote", required=True)
    parser.add_argument("--servicio", action="append", help="código del servicio (TF, TVP...)")
    parser.add_argument("--limite", type=int, default=0)
    parser.add_argument("--nombre", action="append", help="solo la tarifa con ese nombre; se puede repetir")
    parser.add_argument("--ubigeos", default=None)
    parser.add_argument("--solo-listar", action="store_true")
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--si", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = leer_argumentos()
    pedidos = [u.strip() for u in args.ubigeos.split(",") if u.strip()] if args.ubigeos else None
    ubigeos = pedidos or list(UBIGEOS_POR_DEFECTO)
    ruta_registro = comun.DIRECTORIO_RESULTADOS / f"registro_definitivo_{args.lote}.csv"

    tarifas = leer_tarifas_del_lote(args.lote, args.servicio)
    if not tarifas:
        print(f"No hay tarifas EXITOSAS del lote '{args.lote}' en resultados/*/resultados_{args.lote}.csv")
        sys.exit(1)

    for tarifa in tarifas:
        tarifa["requiere_disponibilidad"] = requiere_disponibilidad(tarifa)

    registradas = leer_registradas(ruta_registro)
    pendientes = [t for t in tarifas if t["nombre"] not in registradas]
    if args.nombre:
        faltan = [n for n in args.nombre if n not in {t["nombre"] for t in tarifas}]
        if faltan:
            print(f"No están entre las tarifas EXITOSAS del lote: {', '.join(faltan)}")
            sys.exit(1)
        pendientes = [t for t in pendientes if t["nombre"] in args.nombre]

    # Ubicaciones de cada tarifa: las pedidas (o Amazonas); en paquetes, según sus tarifas.
    con_disponibilidad = disponibilidad_registrada()
    descartadas = []
    for tarifa in list(pendientes):
        tarifa["ubigeos"] = ubigeos if tarifa["requiere_disponibilidad"] else []
        if tarifa["codigo_servicio"] == CODIGO_PAQUETE and tarifa["requiere_disponibilidad"]:
            try:
                tarifa["ubigeos"] = ubigeos_del_paquete(tarifa, pedidos, con_disponibilidad)
            except ValueError as problema:
                descartadas.append(f"{tarifa['nombre']}: {problema}")
                pendientes.remove(tarifa)

    if args.limite > 0:
        pendientes = pendientes[: args.limite]

    print(f"Lote: {args.lote} | Ubicaciones (monoproductos): {', '.join(ubigeos)}")
    print(f"Tarifas del lote: {len(tarifas)} | Ya registradas: {len(registradas)} | A procesar: {len(pendientes)}\n")
    for tarifa in pendientes:
        marca = f"disponibilidad: {', '.join(tarifa['ubigeos'])}" if tarifa["requiere_disponibilidad"] else "sin disponibilidad"
        print(f"  {tarifa['nombre']:<24} {marca:<34} (registro {tarifa['id_registro']})")
    if descartadas:
        print("\nNo se procesarán (disponibilidad del paquete):")
        for texto in descartadas:
            print(f"  - {texto}")

    if args.solo_listar or not pendientes:
        print("\n--solo-listar: no se tocó SIRT." if args.solo_listar else "\nNo hay pendientes.")
        return

    if args.limite == 0 and len(pendientes) > 1 and not args.si:
        print(f"\nATENCIÓN: se harán {len(pendientes)} REGISTROS DEFINITIVOS (probablemente irreversibles).")
        respuesta = input("Escribe 'si' para continuar (Enter para cancelar): ").strip().lower()
        if respuesta not in ("si", "sí"):
            print("Cancelado. Para probar usa --limite 1.")
            return

    from playwright.sync_api import sync_playwright

    diag = Diagnostico(
        activo=args.diagnostico,
        con_trace=False,
        carpeta=comun.DIRECTORIO_RESULTADOS / f"diagnostico_registro_{args.lote}",
    )

    exitosas = 0
    errores = 0

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        diag.iniciar(context)

        if args.al_frente:
            page.bring_to_front()

        for numero, tarifa in enumerate(pendientes, start=1):
            print("\n" + "=" * 70)
            print(f"[{numero}/{len(pendientes)}] {tarifa['nombre']} "
                  f"({'con' if tarifa['requiere_disponibilidad'] else 'sin'} disponibilidad)")

            estado = {"etapa": ""}
            diag.iniciar_caso({"correlativo": numero, "nombre": tarifa["nombre"]})
            fila = {
                "nombre": tarifa["nombre"],
                "codigo_servicio": tarifa["codigo_servicio"],
                "id_registro": tarifa["id_registro"],
                "disponibilidad": ",".join(tarifa["ubigeos"]),
                "resultado": "",
                "detalle": "",
                "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            try:
                fila["detalle"] = procesar_tarifa(page, tarifa, tarifa["ubigeos"], estado, diag)
                fila["resultado"] = "EXITOSO"
                exitosas += 1
                print(f"Resultado: REGISTRO DEFINITIVO OK. {fila['detalle']}")

            except Exception as error:
                diag.registrar_fallo(page, error, estado["etapa"])
                fila["resultado"] = "ERROR"
                fila["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
                errores += 1
                print(f"Resultado: ERROR en '{estado['etapa']}' - {error}")

                # Cerrar lo que haya quedado abierto; la siguiente tarifa
                # empieza recargando la página (sin filas marcadas).
                for _ in range(3):
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(300)

            registrar(ruta_registro, fila)
            page.wait_for_timeout(PAUSA_ENTRE_TARIFAS_MS)

        diag.finalizar()

    print("\n" + "=" * 70)
    print("PROCESO FINALIZADO")
    print(f"Registradas en esta ejecución: {exitosas}")
    print(f"Errores en esta ejecución:     {errores}")
    print(f"Resultados: {ruta_registro.resolve()}")


if __name__ == "__main__":
    main()
