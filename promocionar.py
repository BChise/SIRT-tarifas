"""
Crea una TARIFA PROMOCIONAL a partir de una o más tarifas con registro definitivo
(monoproductos o empaquetadas).

Flujo en SIRT (grabaciones/_promocional/):
  Paso 1: Tipo de Tarifa "Tarifa promocional" → "Agregar Tarifa" por cada tarifa
          (buscar por nombre, marcar, Agregar; la Condición se llena sola) →
          nombre, comercialización, condiciones → Datos promocionales (Valor del
          plan y, si el servicio la tiene, Tarifa instalación) → Guardar y continuar.
  Paso 2: una pestaña "Servicio N" por SERVICIO (no por tarifa). En cada una se
          marcan los atributos a promocionar; cada fila lleva Tipo de variación,
          Variación y Duración (días) → Guardar.
  Paso 3: Confirmar Tarifa → Confirmar tarifa → Aceptar.

Qué atributos se promocionan
  - Por defecto, los que se usaron en la grabación de cada servicio (probados).
  - --atributos todos: todos los que trae el JSON del paso 2 del servicio
    (grabaciones/_promocional/<servicio>/paso2.json, 2.º JSON del archivo).
  - --atributos "A;B": solo esos (nombres exactos de las casillas).

Valores
  - Tipo de variación: se elige en pantalla, con preferencia Porcentual →
    Cantidad → Ilimitado (Ilimitado deja la Variación vacía).
  - Variación: 10 si es Porcentual, 1 si es Cantidad. Duración: 30 días.
    Se cambian con --variacion-porcentual, --variacion-cantidad y --duracion.
  - Canales (Carga y descarga): se adjunta el Excel de tv_paga_postpago.
  - Apps sin consumo de datos: se elige la primera plataforma.
  - Regla de SIRT: la duración no supera 180 días.

Uso:
    python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --solo-listar
    python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --diagnostico
    python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --atributos todos --solo-listar
    python promocionar.py --tarifas TE-EMP-MUESTRA-006 --lote PROMO --diagnostico

Nombre de la promoción: TP-<código>-<lote>-NNN (código del servicio; EMP si es
un paquete; MIX si se mezclan varias tarifas). Resultados en
resultados/PRM/resultados_<lote>.csv.
"""

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import comun
from empaquetar import (
    abrir_nueva_tarifa,
    agregar_tarifa_al_paquete,
    buscar_tarifa_base,
    confirmar_paquete,
    limpiar_listado,
    servicios_por_codigo,
    siguiente_nombre,
    texto_dialogo,
    texto_visible,
    tiene_registro_definitivo,
)
from ejecutar import Diagnostico


CODIGO_RESULTADOS = "PRM"
TIPO_TARIFA_PROMOCIONAL = "Tarifa promocional"
CARPETA_GRABACIONES = Path(__file__).parent / "grabaciones" / "_promocional"
ARCHIVO_CANALES = Path(__file__).parent / "grabaciones" / "tv_paga_postpago" / "Canales analógicos_1.xlsx"

PREFERENCIA_TIPO = ["Porcentual", "Cantidad", "Ilimitado"]
DURACION_MAXIMA = 180

# Atributos usados en las grabaciones (probados en SIRT el 15/10).
ATRIBUTOS_PROBADOS = {
    "TF": ["Minutos todo destino nacional", "Minutos red fija on net", "Minutos red fija off net"],
    "TFP": ["Minutos todo destino nacional", "Minutos móvil on net", "Tarifa por minuto a móvil LDI"],
    "OTF": ["Tarifa por minuto"],
    "PTF": ["Tarifa por minuto a fijo LDI", "Tarifa por minuto a móvil LDI"],
    "IF": ["Velocidad de descarga contratada", "Velocidad de descarga garantizada",
           "Velocidad de carga contratada"],
    "PIF": ["Velocidad de descarga contratada", "Velocidad de descarga garantizada", "Cantidad de datos"],
    "TVP": ["Canales analógicos", "Decodificadores o similares"],
    "TVPR": ["Canales analógicos", "Decodificadores o similares"],
    "PTV": ["Canales analógicos", "Decodificadores o similares"],
    "SMP": ["Datos para tethering", "Traslada datos", "Datos para Apps 1",
            "Apps sin consumo de datos", "Tarifa por MB adicional"],
    "SMC": ["Datos para tethering", "Traslada datos", "Datos para Apps 1",
            "Apps sin consumo de datos", "SMS Roaming dentro del plan", "Min. Roaming dentro del plan"],
    "SMQ": ["Duración", "Datos para tethering", "Datos para Apps 1", "Datos tope descarga",
            "Datos Roaming incluidos en los datos del plan"],
    "SMO": ["Duración", "Traslada datos", "Datos para Apps 1", "Minutos móvil nacional",
            "Tarifa minuto móvil todo destino", "SMS Roaming dentro del plan"],
    "AC": ["Velocidad de transmisión", "Vel. transmisión garant."],
    "OS": [],   # su JSON no trae atributos promocionables: solo se promociona el paso 1
}

# Servicios con "Tarifa instalación" (registraTarifaInstalacion del paso 2 de cada servicio).
CON_INSTALACION = {"TFP", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "OS"}


# --- Catálogo de atributos (JSON del paso 2) -------------------------------------

def cargar_jsons(ruta: Path) -> list:
    """El archivo puede traer varios JSON seguidos (respuestas pegadas una tras otra)."""
    texto = ruta.read_text(encoding="utf-8") if ruta.exists() else ""
    decodificador, posicion, documentos = json.JSONDecoder(), 0, []
    while True:
        while posicion < len(texto) and texto[posicion].isspace():
            posicion += 1
        if posicion >= len(texto):
            return documentos
        documento, posicion = decodificador.raw_decode(texto, posicion)
        documentos.append(documento)


def catalogo_atributos(carpeta: str) -> list[dict] | None:
    """Atributos promocionables de un servicio: [{nombre, tipo, opciones, unidad}].
    Un apartado con sub-apartados (tabla de velocidades) aporta una fila por sub-apartado.
    None si no hay JSON del catálogo."""
    for documento in cargar_jsons(CARPETA_GRABACIONES / carpeta / "paso2.json"):
        if not isinstance(documento.get("data"), list):
            continue
        atributos = []
        for seccion in documento["data"]:
            for apartado in seccion.get("apartados") or []:
                opciones = [v["nombre"] for v in apartado.get("variacion") or []]
                for sub in apartado.get("subApartados") or [apartado]:
                    atributos.append({
                        "nombre": sub["nombre"].strip(),
                        "tipo": apartado["tipo"],
                        "opciones": opciones,
                        "unidad": sub.get("unidadMedida") or apartado.get("unidadMedida") or "",
                    })
        return atributos
    return None


# --- Filas de atributos en pantalla ----------------------------------------------

# Marca los controles que siguen a una casilla, hasta la siguiente casilla, y
# devuelve qué es cada uno. Sirve tanto si la fila es un contenedor como si las
# columnas están separadas: se usa el orden del documento.
MARCAR_CONTROLES = """
([casilla, id]) => {
    const todos = Array.from(document.querySelectorAll(
        "input[type=checkbox], mat-select, input:not([type=checkbox]), textarea, button"));
    const inicio = todos.indexOf(casilla);
    const controles = [];
    for (let k = inicio + 1; k < todos.length; k++) {
        const el = todos[k];
        if (el.tagName === "INPUT" && el.type === "checkbox") break;
        controles.push(el);
    }
    return controles.map((el, n) => {
        el.setAttribute("data-promo", `${id}-${n}`);
        const campo = el.closest("mat-form-field");
        const texto = [
            el.getAttribute("placeholder") || "",
            el.getAttribute("aria-label") || "",
            campo ? (campo.innerText || "") : "",
        ].join(" ");
        return {
            sel: `[data-promo="${id}-${n}"]`,
            tag: el.tagName,
            type: el.type || "",
            texto: texto.replace(/\\s+/g, " ").trim(),
            visible: !!(el.offsetParent || el.type === "file"),
            deshabilitado: !!(el.disabled || el.getAttribute("aria-disabled") === "true"),
        };
    });
}
"""

_contador = {"n": 0}


def casilla_visible(page, etiqueta: str):
    """La casilla visible con ese nombre (exacto; si no, que empiece igual)."""
    for candidata in (
        page.get_by_role("checkbox", name=etiqueta, exact=True),
        page.get_by_role("checkbox", name=re.compile(rf"^\s*{re.escape(etiqueta)}")),
    ):
        for casilla in candidata.all():
            if casilla.is_visible():
                return casilla
    return None


def controles_de(page, casilla) -> dict:
    """Tipo de variación / plataforma (mat-select), Variación, Duración y archivo de la fila."""
    _contador["n"] += 1
    marcados = page.evaluate(MARCAR_CONTROLES, [casilla.element_handle(), f"f{_contador['n']}"])

    controles = {"select": None, "variacion": None, "duracion": None, "archivo": None, "botones": []}
    entradas = []
    for c in marcados:
        if c["tag"] == "BUTTON":
            if not entradas and controles["select"] is None:   # botones antes de los campos
                controles["botones"].append(page.locator(c["sel"]))
        elif c["tag"] == "MAT-SELECT" and controles["select"] is None:
            controles["select"] = page.locator(c["sel"])
        elif c["type"] == "file":
            controles["archivo"] = page.locator(c["sel"])
        elif c["tag"] in ("INPUT", "TEXTAREA") and c["visible"]:
            entradas.append(c)

    for c in entradas:
        if "Duraci" in c["texto"] and controles["duracion"] is None:
            controles["duracion"] = page.locator(c["sel"])
        elif "Variaci" in c["texto"] and controles["variacion"] is None:
            controles["variacion"] = page.locator(c["sel"])
    # Respaldo por posición si los campos no dicen qué son: primero la variación,
    # al final la duración.
    if controles["variacion"] is None and controles["duracion"] is None and entradas:
        controles["variacion"] = page.locator(entradas[0]["sel"])
        if len(entradas) >= 2:
            controles["duracion"] = page.locator(entradas[-1]["sel"])
    return controles


TEXTO_DE_FILA = """
casilla => {
    const casillas = Array.from(document.querySelectorAll("input[type=checkbox]"));
    const siguiente = casillas[casillas.indexOf(casilla) + 1];
    const rango = document.createRange();
    rango.setStartAfter(casilla);
    if (siguiente) rango.setEndBefore(siguiente); else rango.setEndAfter(document.body.lastChild);
    return rango.toString();
}
"""


def cargar_en_fila(page, casilla, controles: dict, etiqueta: str, ruta: Path) -> None:
    """Adjunta el archivo con el botón de carga de ESTA fila (el último botón
    antes de los campos: descarga, carga). El campo oculto de archivo no sirve:
    en SIRT está antes de su casilla y se confundía con el de la fila siguiente."""
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo a cargar: {ruta.resolve()}")
    botones = [b for b in controles["botones"] if b.is_enabled()]
    if not botones:
        raise RuntimeError(f"'{etiqueta}': no se encontró el botón para cargar el archivo.")

    with page.expect_file_chooser() as seleccion:
        botones[-1].click()
    seleccion.value.set_files(str(ruta))

    for _ in range(40):                                   # hasta ~10 s
        if ruta.name in casilla.evaluate(TEXTO_DE_FILA):
            return
        page.wait_for_timeout(250)
    raise RuntimeError(f"'{etiqueta}': el archivo {ruta.name} no aparece en su fila tras cargarlo.")


def opciones_abiertas(page) -> list[str]:
    opciones = page.locator(".cdk-overlay-container [role='option']")
    opciones.first.wait_for(state="visible")
    return [o.strip() for o in opciones.all_inner_texts()]


def elegir_tipo(page, select, preferido: str | None) -> str:
    select.click()
    disponibles = opciones_abiertas(page)
    orden = ([preferido] if preferido else []) + PREFERENCIA_TIPO
    tipo = next((t for t in orden if t in disponibles), disponibles[0])
    comun.opcion_en_overlay(page, tipo).click()
    return tipo


def cerrar_overlay(page) -> None:
    fondo = page.locator(".cdk-overlay-backdrop")
    if fondo.count():
        fondo.last.click()
    page.keyboard.press("Escape")


def promocionar_fila(page, etiqueta: str, config: dict, con_duracion: bool = True) -> str:
    """Marca la casilla y llena su fila. Devuelve un resumen de lo puesto."""
    casilla = casilla_visible(page, etiqueta)
    if casilla is None:
        raise RuntimeError(f"No se encontró la casilla '{etiqueta}' en esta pantalla.")
    casilla.check()
    page.wait_for_timeout(300)

    c = controles_de(page, casilla)
    partes = []

    if config.get("archivo"):
        cargar_en_fila(page, casilla, c, etiqueta, Path(config["archivo"]))
        partes.append(f"archivo {Path(config['archivo']).name}")

    elif config.get("plataforma"):
        if c["select"] is None:
            raise RuntimeError(f"'{etiqueta}': no se encontró la lista de plataformas.")
        c["select"].click()
        primera = opciones_abiertas(page)[0]
        comun.opcion_en_overlay(page, primera).click()
        cerrar_overlay(page)
        partes.append(f"plataforma {primera}")

    elif c["select"] is not None:
        tipo = elegir_tipo(page, c["select"], config.get("tipo"))
        page.wait_for_timeout(200)
        c = controles_de(page, casilla)       # Ilimitado puede quitar la variación
        valor = config["porcentual"] if tipo == "Porcentual" else config["cantidad"]
        if tipo != "Ilimitado" and c["variacion"] is not None and comun.llenar_si_editable(c["variacion"], valor):
            partes.append(f"{tipo} {valor}")
        else:
            partes.append(tipo)

    if con_duracion:
        if c["duracion"] is not None and comun.llenar_si_editable(c["duracion"], config["duracion"]):
            partes.append(f"{config['duracion']} días")
        elif c["duracion"] is None:
            partes.append("(sin duración)")

    return f"{etiqueta}: " + ", ".join(partes)


# --- Paso 1 --------------------------------------------------------------------

def elegir_fecha_futura(page, etiqueta: str, fecha: date) -> None:
    """Abre el calendario y avanza mes a mes hasta la fecha (como en la grabación).
    comun.elegir_fecha solo sirve para fechas del mes que muestra el calendario."""
    page.locator("mat-form-field").filter(has_text=etiqueta).get_by_label("Open calendar").click()
    texto = f"{fecha.day} de {comun.MESES[fecha.month - 1]} de {fecha.year}"
    for _ in range(24):
        dia = page.get_by_role("button", name=texto, exact=True)
        if dia.count():
            dia.click()
            page.get_by_role("dialog", name=etiqueta).wait_for(state="hidden")
            return
        page.get_by_role("button", name="Next month").click()
    raise RuntimeError(f"No se encontró el día {texto} en el calendario de '{etiqueta}'.")


def completar_paso_1(page, nombre: str, tarifas: list[str], args, con_instalacion: bool) -> tuple[str, list]:
    comun.elegir_opcion(page, "Tipo de Tarifa", TIPO_TARIFA_PROMOCIONAL)
    quitadas = limpiar_listado(page)
    if quitadas:
        print(f"  Se quitaron {quitadas} tarifa(s) que quedaban de un intento anterior.")
    for numero, tarifa in enumerate(tarifas, start=1):
        agregar_tarifa_al_paquete(page, tarifa, numero)

    hoy = date.today()
    page.get_by_role("textbox", name="Nombre de la tarifa").fill(nombre)
    comun.elegir_fecha(page, "Inicio comercialización", hoy)
    elegir_fecha_futura(page, "Fin comercialización", hoy + timedelta(days=args.dias_comercializacion))
    page.get_by_role("textbox", name="Condiciones generales").fill("Promoción de prueba automatizada")
    otras = page.get_by_role("textbox", name="Otras condiciones tarifa")
    if otras.count():
        otras.fill("Otras condiciones de prueba")

    hechos = []
    base = {"porcentual": args.variacion_porcentual, "cantidad": args.variacion_cantidad, "duracion": args.duracion}
    if not args.sin_valor_plan:
        hechos.append(promocionar_fila(page, "Valor del plan", base))
    if con_instalacion and not args.sin_instalacion:
        if casilla_visible(page, "Tarifa instalación") is not None:
            hechos.append(promocionar_fila(page, "Tarifa instalación", base, con_duracion=False))
        else:
            hechos.append("Tarifa instalación: no aparece (se omite)")

    return comun.guardar_paso_1(page), hechos


# --- Paso 2 --------------------------------------------------------------------

def guardar_pestana(page) -> None:
    """Guardar; si SIRT abre una ventana (error), se lee, se cierra y se informa.
    Con un solo servicio el botón es "Guardar"; en un empaquetado (varias pestañas
    "Servicio N") es "Guardar e ir al siguiente" (visto el 16/10 con TE-EMP-MUESTRA-005)."""
    boton = page.get_by_role("button", name=re.compile(r"^Guardar( e ir al siguiente)?$"))
    visibles = [b for b in boton.all() if b.is_visible()]
    if not visibles:
        raise RuntimeError("No aparece el botón 'Guardar' / 'Guardar e ir al siguiente' de la pestaña.")
    visibles[0].click()
    aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    try:
        aceptar.wait_for(state="visible", timeout=2_500)
    except Exception:
        return
    mensaje = texto_dialogo(page)
    aceptar.click()
    raise RuntimeError(f"SIRT no aceptó la pestaña: {mensaje[:400]}")


def completar_paso_2(page, servicios: list[dict], args) -> list[str]:
    hechos = []
    pendientes = list(servicios)

    for numero in range(1, len(servicios) + 1):
        pestana = page.get_by_role("tab", name=f"Servicio {numero}", exact=True)
        pestana.wait_for(state="visible")
        pestana.click()
        page.wait_for_timeout(500)

        servicio = next((s for s in pendientes if texto_visible(page, s["nombre_sirt"])), None)
        if servicio is None:
            raise RuntimeError(
                f"La pestaña 'Servicio {numero}' no muestra ninguno de: "
                + ", ".join(s["nombre_sirt"] for s in pendientes))
        pendientes.remove(servicio)

        for atributo in servicio["atributos"]:
            hechos.append(f"[{servicio['codigo']}] " + promocionar_fila(page, atributo["nombre"], atributo["config"]))

        if servicio["atributos"]:
            guardar_pestana(page)
            page.wait_for_timeout(800)

    page.get_by_role("button", name="Continuar con el paso").click()
    confirmar = page.get_by_role("button", name="Confirmar Tarifa", exact=True)
    try:
        confirmar.wait_for(state="visible", timeout=15_000)
    except Exception:
        raise RuntimeError(f"No pasó al resumen. Mensaje de SIRT: {texto_dialogo(page)[:300]}")
    return hechos


# --- Preparación ---------------------------------------------------------------

def config_atributo(atributo: dict, args) -> dict:
    base = {"porcentual": args.variacion_porcentual, "cantidad": args.variacion_cantidad,
            "duracion": args.duracion, "tipo": None}
    if atributo["tipo"] == "PROMOCIONAL_ADJUNTO":
        return {**base, "archivo": ARCHIVO_CANALES}
    if atributo["tipo"] == "PROMOCIONAL_APP":
        return {**base, "plataforma": True}
    return base


def servicios_de_las_tarifas(tarifas: list[str], args) -> tuple[list[dict], list[str]]:
    """Servicios que tendrán pestaña en el paso 2 (los de un paquete se abren en sus tarifas)."""
    modulos = servicios_por_codigo()
    problemas, codigos = [], []

    for nombre in tarifas:
        base = buscar_tarifa_base(nombre)
        if base is None:
            problemas.append(f"'{nombre}': no está como EXITOSO en resultados/*/resultados_*.csv")
            continue
        if not tiene_registro_definitivo(nombre):
            problemas.append(f"'{nombre}': no tiene registro definitivo")
        if base["codigo"] == "EMP":
            fila = next(iter(_filas_resultado("EMP", nombre)), None)
            componentes = fila["clave"].split(" | ")[1:] if fila else []
            codigos += [c.split("-")[1] for c in componentes]
        else:
            codigos.append(base["codigo"])

    servicios = []
    for codigo in dict.fromkeys(codigos):           # sin repetir, en orden
        modulo = modulos.get(codigo)
        if modulo is None:
            problemas.append(f"Servicio de código {codigo} desconocido")
            continue
        carpeta = Path(modulo.__file__).stem
        catalogo = catalogo_atributos(carpeta)

        if args.atributos == "probados":
            nombres = ATRIBUTOS_PROBADOS.get(codigo, [])
        elif args.atributos == "todos":
            if catalogo is None:
                problemas.append(f"{codigo}: no hay catálogo en grabaciones/_promocional/{carpeta}/paso2.json "
                                 "(falta el 2.º JSON); usa --atributos probados")
                continue
            nombres = [a["nombre"] for a in catalogo]
        else:
            nombres = [n.strip() for n in args.atributos.split(";") if n.strip()]

        por_nombre = {a["nombre"]: a for a in catalogo or []}
        atributos = []
        for n in nombres:
            info = por_nombre.get(n, {"nombre": n, "tipo": "PROMOCIONAL", "opciones": [], "unidad": ""})
            if catalogo is not None and n not in por_nombre:
                print(f"  Aviso: '{n}' no está en el catálogo de {codigo}; se intentará igual.")
            atributos.append({**info, "config": config_atributo(info, args)})

        servicios.append({"codigo": codigo, "nombre_sirt": modulo.SERVICIO, "atributos": atributos})

    return servicios, problemas


def _filas_resultado(codigo: str, nombre: str):
    import csv
    for ruta in sorted((comun.DIRECTORIO_RESULTADOS / codigo).glob("resultados_*.csv")):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            for fila in csv.DictReader(entrada, delimiter=";"):
                if fila.get("nombre") == nombre and fila.get("resultado") == "EXITOSO":
                    yield fila


def leer_argumentos():
    parser = argparse.ArgumentParser(description="Crea una tarifa promocional.")
    parser.add_argument("--tarifas", nargs="+", required=True, help="tarifas a promocionar (con registro definitivo)")
    parser.add_argument("--lote", required=True)
    parser.add_argument("--nombre", help="nombre de la promoción (en vez del automático)")
    parser.add_argument("--atributos", default="probados", help="probados | todos | 'Nombre A;Nombre B'")
    parser.add_argument("--variacion-porcentual", default="10")
    parser.add_argument("--variacion-cantidad", default="1")
    parser.add_argument("--duracion", default="30", help="días (máximo 180)")
    parser.add_argument("--dias-comercializacion", type=int, default=60)
    parser.add_argument("--sin-valor-plan", action="store_true")
    parser.add_argument("--sin-instalacion", action="store_true")
    parser.add_argument("--solo-listar", action="store_true")
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = leer_argumentos()
    if int(args.duracion) > DURACION_MAXIMA:
        sys.exit(f"La duración no puede superar {DURACION_MAXIMA} días.")

    servicios, problemas = servicios_de_las_tarifas(args.tarifas, args)
    if problemas:
        sys.exit("No se puede armar la promoción:\n  - " + "\n  - ".join(problemas))

    codigos = [s["codigo"] for s in servicios]
    if any(t.split("-")[1] == "EMP" for t in args.tarifas):
        prefijo = "EMP" if len(args.tarifas) == 1 else "MIX"
    else:
        prefijo = codigos[0] if len(set(codigos)) == 1 else "MIX"
    rutas = comun.rutas_lote(CODIGO_RESULTADOS, args.lote)
    correlativo, nombre_auto = siguiente_nombre(args.lote, rutas["resultados"])
    nombre = args.nombre or nombre_auto.replace(f"TE-EMP-", f"TP-{prefijo}-")
    con_instalacion = any(c in CON_INSTALACION for c in codigos)

    print(f"Promoción: {nombre}")
    print(f"Tarifas:   {', '.join(args.tarifas)}")
    print(f"Paso 1:    Valor del plan {'NO' if args.sin_valor_plan else 'sí'}"
          f" | Tarifa instalación {'sí' if con_instalacion and not args.sin_instalacion else 'no'}")
    for s in servicios:
        print(f"Paso 2 - {s['nombre_sirt']} ({s['codigo']}): {len(s['atributos'])} atributo(s)")
        for a in s["atributos"]:
            extra = "archivo" if a["config"].get("archivo") else "plataforma" if a["config"].get("plataforma") else \
                "/".join(a["opciones"]) or "tipo en pantalla"
            print(f"    - {a['nombre']:<50} {extra} {a['unidad']}")
    print(f"Valores: Porcentual {args.variacion_porcentual} | Cantidad {args.variacion_cantidad}"
          f" | Duración {args.duracion} días")

    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT.")
        return

    from playwright.sync_api import sync_playwright

    clave = " + ".join(args.tarifas) + f" | {args.atributos}"
    diag = Diagnostico(args.diagnostico, False, rutas["base"] / f"diagnostico_{args.lote}")
    estado = {"etapa": "", "id_registro": ""}
    fila = {
        "correlativo": correlativo, "nombre": nombre, "clave": clave,
        "id_registro": "", "resultado": "", "detalle": "", "captura": "",
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        diag.iniciar(context)
        diag.iniciar_caso({"correlativo": correlativo, "nombre": nombre})
        if args.al_frente:
            page.bring_to_front()

        def etapa(nombre_etapa, accion):
            estado["etapa"] = nombre_etapa
            resultado = accion()
            diag.captura_paso(page, nombre_etapa)
            return resultado

        try:
            etapa("abrir_nueva_tarifa", lambda: abrir_nueva_tarifa(page))
            id_registro, hechos1 = etapa(
                "paso_1_guardado", lambda: completar_paso_1(page, nombre, args.tarifas, args, con_instalacion))
            estado["id_registro"] = id_registro
            for h in hechos1:
                print(f"  {h}")
            hechos2 = etapa("paso_2_completo", lambda: completar_paso_2(page, servicios, args))
            for h in hechos2:
                print(f"  {h}")
            etapa("confirmada", lambda: confirmar_paquete(page))

            fila["resultado"] = "EXITOSO"
            fila["detalle"] = " | ".join(hechos1 + hechos2)[:900]
            print(f"\nResultado: PROMOCIÓN CONFIRMADA (registro {estado['id_registro']})")

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
