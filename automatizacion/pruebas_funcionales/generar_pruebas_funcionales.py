"""
Genera los Excel de la PRUEBA FUNCIONAL (día 1) en automatizacion/pruebas_funcionales/, con el
MISMO formato que los Excel de trabajo (se copian y se cambian solo las filas). No toca los
originales. Cada fila trae dos columnas más al final: «Qué se prueba» y «Resultado esperado»
(las lee verificar_funcional.py).

  registro_funcional.xlsx      32 filas: 2 tarifas por servicio (TE-<COD>-FN-011 y -012) + 2 de control
  empaquetados_funcional.xlsx  10 filas: de 2 a 6 tarifas, valores vacíos / escritos / mixtos, errores
  licitaciones_funcional.xlsx   6 filas: 1, 3 y 6 servicios, instalación, monedas, errores
  promociones_funcional.xlsx   12 filas: todos los modos de atributos, preferidos, mezclas, errores

Tarifas: FN-011 = base de empaquetados y promociones (no recibirá acciones); FN-012 = libre para
las acciones del día 2.

Cada serie: registro con FN-011 (bases), FN-012 (acciones), FN-013/014 (modificaciones del día 2).
Uso (los Excel de trabajo solo se leen):
    python automatizacion\\pruebas_funcionales\\generar_pruebas_funcionales.py --serie P4 --carpeta Prueba4
    python automatizacion\\pruebas_funcionales\\generar_pruebas_funcionales.py              (serie FN)
    python automatizacion\\pruebas_funcionales\\generar_pruebas_funcionales.py --serie FA   (repetir con tarifas nuevas)
Antes de reemplazar, copia los Excel anteriores (con sus resultados) en pruebas_funcionales/anteriores/<fecha>/.
"""

import shutil
import sys
from copy import copy
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
PROYECTO = CARPETA.parent.parent
sys.path.insert(0, str(PROYECTO))
import excel_comun as xc  # noqa: E402

AUTO = PROYECTO / "automatizacion"
# Excel de donde se copian el formato y las filas modelo (TE-…-CM-001, etc.): la carpeta «base»
# junto a este archivo si existe (paquete de entrega, donde automatizacion/ trae plantillas vacías);
# si no, los Excel de trabajo de automatizacion/.
ORIGEN = CARPETA / "base" if (CARPETA / "base").is_dir() else AUTO
MANANA = date.today() + timedelta(days=1)
SERVICIOS = ["TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "SMP", "SMC", "SMO", "SMQ", "OS", "AC"]
CON_DOS_MODIFICACIONES = {"TF", "IF", "SMC", "SMP", "TVP"}   # servicios con 2 filas en modificaciones
CON_DISPONIBILIDAD = {"TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV"}
SALIDA = CARPETA      # --carpeta NOMBRE → pruebas_funcionales/NOMBRE (una carpeta por serie)
SERIE = "FN"          # --serie: otra letra/número para repetir la prueba con tarifas nuevas
RESULTADO = ("Estado", "Código SIRT", "Fecha de registro", "Fecha de ejecución", "Detalle", "ID de registro")


def fn(codigo: str, n: int) -> str:
    return f"TE-{codigo}-{SERIE}-0{n}"


# --- utilidades de copia ----------------------------------------------------------------------

def cabecera(ws) -> dict:
    return {xc.texto(c.value): c.column for c in ws[xc.FILA_CABECERA] if c.value is not None}


def filas_con_nombre(ws, columna: int) -> dict:
    return {xc.texto(ws.cell(r, columna).value): r for r in range(xc.FILA_CABECERA + 1, ws.max_row + 1)
            if xc.texto(ws.cell(r, columna).value)}


def vaciar(ws) -> None:
    for fila in ws.iter_rows(min_row=xc.FILA_CABECERA + 1, max_row=ws.max_row):
        for celda in fila:
            celda.value = None


def copiar_fila(ws_origen, r_origen, ws_destino, r_destino, cambios: dict, col: dict) -> None:
    """Copia valores y estilos de una fila; aplica cambios {título: valor}; deja vacío el resultado."""
    from openpyxl.formula.translate import Translator
    for c in range(1, ws_origen.max_column + 1):
        a, b = ws_origen.cell(r_origen, c), ws_destino.cell(r_destino, c)
        valor = a.value
        if isinstance(valor, str) and valor.startswith("="):
            valor = Translator(valor, origin=a.coordinate).translate_formula(b.coordinate)
        b.value = valor
        if a.has_style:
            for atributo in ("font", "fill", "border", "alignment"):
                setattr(b, atributo, copy(getattr(a, atributo)))
            b.number_format = a.number_format
    for titulo in RESULTADO:
        if titulo in col:
            ws_destino.cell(r_destino, col[titulo]).value = None
    for titulo, valor in cambios.items():
        if isinstance(valor, str):
            valor = valor.replace("-FN-", f"-{SERIE}-")
        celda = ws_destino.cell(r_destino, col[titulo])
        celda.value = valor                     # (cell(..., None) no borra: se asigna aparte)
        if isinstance(valor, date):
            celda.number_format = "DD/MM/YYYY"


def columnas_prueba(ws) -> tuple[int, int]:
    """Agrega «Qué se prueba» y «Resultado esperado» al final (si no están)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    col = cabecera(ws)
    if "Qué se prueba" in col:
        return col["Qué se prueba"], col["Resultado esperado"]
    base = max(col.values())
    for k, (titulo, ancho) in enumerate((("Qué se prueba", 60), ("Resultado esperado", 18)), start=1):
        celda = ws.cell(xc.FILA_CABECERA, base + k, titulo)
        celda.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="7030A0")
        celda.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(base + k)].width = ancho
    return base + 1, base + 2


def anotar(ws, r: int, cols: tuple[int, int], prueba: str, esperado: str) -> None:
    from openpyxl.styles import Alignment, Font
    for c, v in zip(cols, (prueba, esperado)):
        celda = ws.cell(r, c, v)
        celda.font = Font(name="Arial", size=9, bold=c == cols[1])
        celda.alignment = Alignment(wrap_text=True, vertical="top")


def copiar_libro(origen: Path, destino: Path):
    libro = xc.abrir_libro(origen, True)            # restaura las listas desplegables
    libro.save(destino)
    return xc.abrir_libro(destino, True)


# --- 1. Registro de tarifas establecidas -----------------------------------------------------------

def registro() -> dict:
    """Devuelve {nombre FN: valor del servicio} de las FN-011 (para los empaquetados)."""
    destino = SALIDA / "registro_funcional.xlsx"
    libro = copiar_libro(ORIGEN / "registro_tarifas_establecidas.xlsx", destino)
    original = xc.abrir_libro(ORIGEN / "registro_tarifas_establecidas.xlsx", True)
    valores = {}
    tablas = [t for t in libro.sheetnames if t.startswith("Tablas")]
    filas_tablas = {t: [] for t in tablas}
    for t in tablas:                                   # filas de tabla de las tarifas origen
        ws_o = original[t]
        col = cabecera(ws_o)
        for r in range(xc.FILA_CABECERA + 1, ws_o.max_row + 1):
            filas_tablas[t].append((r, xc.texto(ws_o.cell(r, col["Nombre de la tarifa"]).value)))
        vaciar(libro[t])
    proximas = {t: xc.FILA_CABECERA + 1 for t in tablas}
    total_filas = 0

    for codigo in SERVICIOS:
        ws_o, ws = original[codigo], libro[codigo]
        col = cabecera(ws)
        nombres = filas_con_nombre(ws_o, col["Nombre de la tarifa"])
        origen = nombres[f"TE-{codigo}-CM-001"]
        vaciar(ws)
        cols = columnas_prueba(ws)
        disp = "Disponibilidad comercial" in col
        filas = [
            (fn(codigo, 11), {"Alcance de la tarifa": "Residencial", "Registro definitivo": "Sí"}
             | ({"Disponibilidad comercial": "Lima"} if disp else {}),
             "Alcance Residencial · tablas copiadas · registro definitivo"
             + (" con disponibilidad Lima" if disp else ""), "Registro definitivo", True),
            (fn(codigo, 12), {"Alcance de la tarifa": "Residencial + Comercial", "Registro definitivo": "Sí"}
             | ({"Disponibilidad comercial": "Lima, Arequipa" if codigo in ("TF", "IF", "TVP") else ""} if disp else {})
             | ({"Inicio de vigencia": MANANA, "Inicio comercialización": MANANA} if codigo in ("TF", "IF", "SMP") else {}),
             "Alcance Residencial + Comercial · SIN filas de tabla (valores por defecto del código)"
             + (" · fechas de inicio = mañana" if codigo in ("TF", "IF", "SMP") else "")
             + (" · disponibilidad Lima + Arequipa" if codigo in ("TF", "IF", "TVP")
                else " · disponibilidad vacía (= la de por defecto)" if disp else ""),
             "Registro definitivo", False),
        ]
        for n in (13, 14) if codigo in CON_DOS_MODIFICACIONES else (13,):
            filas.append((fn(codigo, n), {"Alcance de la tarifa": "Residencial", "Registro definitivo": "Sí"}
                          | ({"Disponibilidad comercial": "Lima"} if disp else {}),
                          "Reservada para modificaciones (día 2) · alcance Residencial · tablas copiadas",
                          "Registro definitivo", True))
        if codigo == "OS":                               # una sin registro definitivo
            filas[1] = (fn(codigo, 12), {"Alcance de la tarifa": "Residencial + Comercial", "Registro definitivo": "No"},
                        "«Registro definitivo» = No → queda solo registrada (temporal)", "Registrada", False)
        if codigo == "TF":
            filas += [(fn(codigo, 90), {"Ejecutar": "No"}, "«Ejecutar» = No", "Omitida", False),
                      (fn(codigo, 91), {"Alcance de la tarifa": "Rural"}, "ERROR esperado: alcance no válido",
                       "Error de datos", False)]
        total_filas += len(filas)
        for k, (nombre, cambios, prueba, esperado, con_tablas) in enumerate(filas):
            r = xc.FILA_CABECERA + 1 + k
            copiar_fila(ws_o, origen, ws, r, {"N°": k + 1, "Nombre de la tarifa": nombre, **cambios}, col)
            anotar(ws, r, cols, prueba, esperado)
            if k == 0:
                valores[nombre] = xc.texto(ws.cell(r, col["Valor del servicio"]).value)
            if con_tablas:
                for t in tablas:
                    ct = cabecera(original[t])
                    for r_t, n in filas_tablas[t]:
                        if n == f"TE-{codigo}-CM-001":
                            copiar_fila(original[t], r_t, libro[t], proximas[t], {"Nombre de la tarifa": nombre}, ct)
                            proximas[t] += 1
    libro.save(destino)
    print(f"  {destino.name}: {total_filas} filas")
    return valores


# --- 2. Empaquetados -----------------------------------------------------------------------------

def empaquetados(valores: dict) -> None:
    destino = SALIDA / "empaquetados_funcional.xlsx"
    libro = copiar_libro(ORIGEN / "empaquetados.xlsx", destino)
    ws = libro["Empaquetados"]
    col = cabecera(ws)
    origen = xc.FILA_CABECERA + 1
    plantilla = [[c.value for c in fila] for fila in ws.iter_rows(min_row=origen, max_row=origen)]
    ws_o = xc.abrir_libro(ORIGEN / "empaquetados.xlsx", True)["Empaquetados"]
    vaciar(ws)
    cols = columnas_prueba(ws)

    def menos(nombre, d):
        return str(Decimal(valores[nombre]) - d)

    b = lambda c: fn(c, 11)
    paquetes = [
        ("TE-EMP-FN-101", [b("TF"), b("SMP")], {}, {"Tarifa instalación": 10},
         "2 tarifas · valores vacíos (base − 1) · disponibilidad vacía (= intersección de sus tarifas)", "Registro definitivo"),
        ("TE-EMP-FN-102", [b("TVP"), b("TVPR"), b("PTV")], {k: menos(b(s), Decimal(2)) for k, s in
                                                         ((1, "TVP"), (2, "TVPR"), (3, "PTV"))},
         {"Disponibilidad comercial": "Lima"}, "3 tarifas · TODOS los valores escritos (base − 2)", "Registro definitivo"),
        ("TE-EMP-FN-103", [b("TF"), b("TFP"), b("OTF"), b("PTF")], {}, {"Tarifa instalación": 25},
         "4 tarifas de telefonía fija · tarifa de instalación 25 (con fijos SIRT la exige)", "Registro definitivo"),
        ("TE-EMP-FN-104", [b("SMP"), b("SMC"), b("SMO"), b("SMQ")], {1: menos(b("SMP"), Decimal(3))},
         {"Tarifa instalación": None},
         "4 móviles · valores MIXTOS (el 1.º escrito, el resto vacío → aviso base − 1) · SIN tarifa de instalación",
         "Registro definitivo"),
        ("TE-EMP-FN-105", [b("TFP"), b("TVP"), b("IF"), b("SMP"), b("AC")], {}, {"Registro definitivo": "No"},
         "5 tarifas · «Registro definitivo» = No", "Registrada"),
        ("TE-EMP-FN-106", [b("TFP"), b("TVP"), b("IF"), b("SMP"), b("AC"), b("OS")], {},
         {"Inicio de vigencia": MANANA, "Inicio comercialización": MANANA},
         "6 tarifas (máximo) · fechas de inicio = mañana", "Registro definitivo"),
        ("TE-EMP-FN-107", [b("TF"), b("PTV"), b("PIF"), b("SMO")], {},
         {"Enlace Web": "https://www.prueba-funcional.pe", "Fono Atención al cliente": "016000000",
          "Email de atención al cliente": "funcional@prueba.pe", "Otros datos generales": "Paquete de prueba funcional"},
         "Datos de contacto propios (enlace, fono, email, otros datos)", "Registro definitivo"),
        ("TE-EMP-FN-108", [b("TF"), b("SMP")], {}, {"Alcance de la tarifa": "Comercial"},
         "ERROR esperado: alcance Comercial con tarifas Residencial (se detecta antes de tocar SIRT)", "Error"),
        ("TE-EMP-FN-109", [b("TF")], {}, {}, "ERROR esperado: una sola tarifa", "Error de datos"),
        ("TE-EMP-FN-110", [b("IF"), b("SMC")], {}, {"Ejecutar": "No"}, "«Ejecutar» = No", "Omitida"),
    ]
    for k, (nombre, tarifas, vals, extra, prueba, esperado) in enumerate(paquetes):
        r = xc.FILA_CABECERA + 1 + k
        cambios = {"N°": k + 1, "Ejecutar": "Sí", "Nombre del paquete": nombre, "Registro definitivo": "Sí",
                   "Disponibilidad comercial": None, "Alcance de la tarifa": "Residencial", "Tarifa instalación": 10,
                   "Inicio de vigencia": None, "Inicio comercialización": None,
                   "Otros datos generales": prueba[:60]}
        for n in range(1, 7):
            cambios[f"Tarifa {n}"] = tarifas[n - 1] if n <= len(tarifas) else None
            cambios[f"Valor {n} en el paquete"] = vals.get(n)
        cambios.update(extra)
        copiar_fila(ws_o, origen, ws, r, cambios, col)
        anotar(ws, r, cols, prueba, esperado)
    del plantilla
    libro.save(destino)
    print(f"  {destino.name}: {len(paquetes)} filas")


# --- 3. Licitaciones -----------------------------------------------------------------------------

def licitaciones() -> None:
    destino = SALIDA / "licitaciones_funcional.xlsx"
    libro = copiar_libro(ORIGEN / "registro_licitaciones.xlsx", destino)
    ws = libro["Licitaciones"]
    col = cabecera(ws)
    ws_o = xc.abrir_libro(ORIGEN / "registro_licitaciones.xlsx", True)["Licitaciones"]
    origen = xc.FILA_CABECERA + 1
    vaciar(ws)
    cols = columnas_prueba(ws)
    servicios = ["Arrendamiento de circuitos", "Otros servicios", "Televisión de paga", "Internet Fijo",
                 "Servicio Móvil", "Telefonía Fija"]

    def con(*incluidos):
        salida = {}
        for s in servicios:
            salida[f"{s}: incluir"] = "Sí" if s in incluidos else "No"
            salida[f"{s}: detalle del servicio"] = f"Detalle de {s.lower()} (prueba funcional)" if s in incluidos else None
        return salida

    filas = [
        ("TL-LYC-FN-101", con("Arrendamiento de circuitos"), {"Ámbito": "Público", "Moneda": "Soles"},
         "1 servicio · Público · Soles · inicio vacío (= hoy)", "Registro definitivo"),
        ("TL-LYC-FN-102", con("Otros servicios"),
         {"Ámbito": "Privado", "Moneda": "Dolares", "Tarifa instalación (solo con Otros servicios)": 15},
         "1 servicio (Otros servicios) · tarifa de instalación · Privado · Dólares", "Registro definitivo"),
        ("TL-LYC-FN-103", con("Televisión de paga", "Internet Fijo", "Servicio Móvil"),
         {"Moneda": "Euro", "Inicio de vigencia": MANANA, "Duración del contrato (días)": 180},
         "3 servicios · Euro · inicio mañana · duración 180 días", "Registro definitivo"),
        ("TL-LYC-FN-104", con(*servicios), {}, "6 servicios (todos)", "Registro definitivo"),
        ("TL-LYC-FN-105", con(), {}, "ERROR esperado: ningún servicio incluido", "Error de datos"),
        ("TL-LYC-FN-106", con("Telefonía Fija"), {"Ejecutar": "No"}, "«Ejecutar» = No", "Omitida"),
    ]
    for k, (nombre, servs, extra, prueba, esperado) in enumerate(filas):
        r = xc.FILA_CABECERA + 1 + k
        cambios = {"N°": k + 1, "Ejecutar": "Sí", "Nombre de la tarifa": nombre, "Registro definitivo": "Sí",
                   "Identificador de contrato": nombre.replace("TL-LYC", "CTR"),
                   "Tarifa instalación (solo con Otros servicios)": None, **servs, **extra}
        copiar_fila(ws_o, origen, ws, r, cambios, col)
        anotar(ws, r, cols, prueba, esperado)
    libro.save(destino)
    print(f"  {destino.name}: {len(filas)} filas")


# --- 4. Promociones ------------------------------------------------------------------------------

def promociones() -> None:
    destino = SALIDA / "promociones_funcional.xlsx"
    libro = copiar_libro(ORIGEN / "promociones.xlsx", destino)
    ws = libro["Promociones"]
    col = cabecera(ws)
    ws_o = xc.abrir_libro(ORIGEN / "promociones.xlsx", True)["Promociones"]
    origen = xc.FILA_CABECERA + 1
    vaciar(ws)
    cols = columnas_prueba(ws)
    b = lambda c: fn(c, 11)
    filas = [
        ("TP-TF-FN-101", [b("TF")], {"Atributos a promocionar": "Todos", "Tipo de variación preferido": "Porcentual"},
         "Todos los atributos · preferido Porcentual", "Registro definitivo"),
        ("TP-IF-FN-102", [b("IF")], {"Atributos a promocionar": "Todos", "Tipo de variación preferido": "Cantidad",
                                     "Días de comercialización": 30},
         "Todos · preferido Cantidad · 30 días de comercialización", "Registro definitivo"),
        ("TP-SMP-FN-103", [b("SMP")], {"Atributos a promocionar": "Todos", "Tipo de variación preferido": "Ilimitado"},
         "Todos · preferido Ilimitado (donde se permita)", "Registro definitivo"),
        ("TP-TVP-FN-104", [b("TVP")], {"Atributos a promocionar": None, "Inicio comercialización": MANANA},
         "Atributos vacío (= «Los de siempre») · inicio de comercialización mañana", "Registro definitivo"),
        ("TP-SMC-FN-105", [b("SMC")], {"Atributos a promocionar": "Lista",
                                       "Lista de atributos (separados por ;)": "Datos tope descarga; Cantidad SMS nacional"},
         "Lista de 2 atributos", "Registro definitivo"),
        ("TP-PIF-FN-106", [b("PIF")], {"Atributos a promocionar": "Detalle en hoja Atributos"},
         "Detalle en hoja Atributos (Cantidad, Porcentual e Ilimitado, cada uno con su duración)", "Registro definitivo"),
        ("TP-OS-FN-107", [b("OS")], {"Valor del plan: promocionar": "No", "Atributos a promocionar": "Todos"},
         "Sin promocionar el valor del plan (solo instalación) · servicio sin atributos", "Registro definitivo"),
        ("TP-MIX-FN-108", [b("TF"), b("SMP"), b("IF")], {"Atributos a promocionar": "Todos"},
         "3 monoproductos de servicios distintos (paso 2 sin apartados en común → se continúa)", "Registro definitivo"),
        ("TP-EMP-FN-109", ["TE-EMP-FN-101"], {"Atributos a promocionar": "Todos"},
         "Promoción de un empaquetado", "Registro definitivo"),
        ("TP-MIX-FN-110", ["TE-EMP-FN-103", b("TF")], {"Atributos a promocionar": "Todos"},
         "Empaquetado + monoproducto", "Registro definitivo"),
        ("TP-ERR-FN-111", [], {}, "ERROR esperado: sin tarifa base", "Error de datos"),
        ("TP-OMI-FN-112", [b("AC")], {"Ejecutar": "No"}, "«Ejecutar» = No", "Omitida"),
    ]
    for k, (nombre, bases, extra, prueba, esperado) in enumerate(filas):
        r = xc.FILA_CABECERA + 1 + k
        cambios = {"N°": k + 1, "Ejecutar": "Sí", "Nombre de la promoción": nombre, "Registro definitivo": "Sí",
                   "Condiciones generales": prueba[:80], "Lista de atributos (separados por ;)": None,
                   "Inicio comercialización": None}
        for n in range(1, 16):
            cambios[f"Tarifa base {n}"] = bases[n - 1] if n <= len(bases) else None
        cambios.update(extra)
        copiar_fila(ws_o, origen, ws, r, cambios, col)
        anotar(ws, r, cols, prueba, esperado)
    # Hoja Atributos: solo la promoción en modo detalle
    wa = libro["Atributos"]
    ca = cabecera(wa)
    vaciar(wa)
    for k, fila in enumerate([("TP-PIF-FN-106", "PIF", "Cantidad de datos", "Ilimitado", None, "45"),
                              ("TP-PIF-FN-106", "PIF", "Velocidad de descarga contratada", "Porcentual", "20", "30"),
                              ("TP-PIF-FN-106", "PIF", "Repetidores y/o equipos similares", "Cantidad", "1", "60")]):
        for titulo, valor in zip(("Nombre de la promoción", "Servicio", "Atributo", "Tipo de variación", "Variación",
                                  "Duración (días)"), fila):
            wa.cell(xc.FILA_CABECERA + 1 + k, ca[titulo], valor.replace("-FN-", f"-{SERIE}-")
                    if isinstance(valor, str) else valor)
    libro.save(destino)
    print(f"  {destino.name}: {len(filas)} filas")


# --- 5. Acciones básicas (día 2) ----------------------------------------------------------------

def acciones() -> None:
    """Fe de erratas, renovación, cese y renombre sobre las tarifas de la serie (FN-012 y las que no
    son base). Corre el día 2: las acciones piden tarifa vigente y en comercialización."""
    destino = SALIDA / "acciones_funcional.xlsx"
    libro = copiar_libro(ORIGEN / "acciones_basicas.xlsx", destino)
    ws = libro["Acciones"]
    col = cabecera(ws)
    ws_o = xc.abrir_libro(ORIGEN / "acciones_basicas.xlsx", True)["Acciones"]
    origen = xc.FILA_CABECERA + 1
    vaciar(ws)
    cols = columnas_prueba(ws)
    en_5 = date.today() + timedelta(days=6)              # se corre mañana: queda en ~5 días
    d = lambda c: fn(c, 12)
    filas = [  # (nombre, acción, {columna: valor}, qué se prueba, esperado)
        (d("TF"), "Fe de erratas", {}, "Fe de erratas · monoproducto (con fechas de inicio de ayer)", "Aplicada"),
        ("TE-EMP-FN-102", "Fe de erratas", {}, "Fe de erratas · empaquetado", "Aplicada"),
        ("TL-LYC-FN-101", "Fe de erratas", {}, "Fe de erratas · licitación", "Aplicada"),
        ("TP-TF-FN-101", "Fe de erratas", {}, "Fe de erratas · promoción", "Aplicada"),
        ("TE-OTF-CM-002", "Fe de erratas", {},
         "Fe de erratas sobre una tarifa YA renombrada (el 20/10): se aplica sobre su nombre actual (…_REN)", "Aplicada"),
        ("TP-IF-FN-102", "Renovación", {"Renovación: días de comercialización": 45, "Renovación: duración (días)": 40},
         "Renovación con días y duración escritos · promoción de Internet fijo (varias duraciones)", "Aplicada"),
        ("TP-PIF-FN-106", "Renovación", {}, "Renovación con días y duración vacíos (valores por defecto)", "Aplicada"),
        ("TP-EMP-FN-109", "Renovación", {}, "Renovación de la promoción de un empaquetado", "Aplicada"),
        (d("TFP"), "Cese", {}, "Cese con fecha vacía (= mañana)", "Aplicada"),
        ("TE-EMP-FN-104", "Cese", {"Cese: fecha de fin de comercialización": en_5},
         "Cese con fecha escrita · empaquetado", "Aplicada"),
        (d("IF"), "Renombre", {}, "Renombre con sufijo vacío (= _REN)", "Aplicada"),
        (d("PIF"), "Renombre", {"Renombre: sufijo": "_V2"}, "Renombre con sufijo propio (_V2)", "Aplicada"),
        (d("TVP"), "Renombre", {}, "Renombre doble (1.º): → …_REN", "Aplicada"),
        (d("TVP") + "_REN", "Renombre", {"Renombre: sufijo": "_V3"},
         "Renombre doble (2.º): la fila usa el nombre ya renombrado → …_REN_V3", "Aplicada"),
        ("TP-SMP-FN-103", "Renombre", {"Renombre: sufijo": "_V2"}, "Renombre de una promoción", "Aplicada"),
        (d("SMC"), "Renovación", {}, "ERROR esperado: renovación de una tarifa que no es promoción", "Error de datos"),
        ("TP-SMC-FN-105", "Cese", {}, "ERROR esperado: cese de una promoción", "Error de datos"),
        (d("SMO"), "Cese", {"Cese: fecha de fin de comercialización": date.today()},
         "ERROR esperado: fecha de cese que no es posterior a hoy", "Error de datos"),
        (d("SMQ"), "Rectificación", {}, "ERROR esperado: acción no válida en este Excel", "Error de datos"),
        (d("AC"), "Fe de erratas", {"Ejecutar": "No"}, "«Ejecutar» = No", "Omitida"),
    ]
    for k, (nombre, accion, extra, prueba, esperado) in enumerate(filas):
        r = xc.FILA_CABECERA + 1 + k
        cambios = {"N°": k + 1, "Ejecutar": "Sí", "Nombre de la tarifa": nombre, "Acción": accion,
                   "Renombre: sufijo": None, "Cese: fecha de fin de comercialización": None,
                   "Renovación: días de comercialización": None, "Renovación: duración (días)": None, **extra}
        copiar_fila(ws_o, origen, ws, r, cambios, col)
        anotar(ws, r, cols, prueba, esperado)
    libro.save(destino)
    print(f"  {destino.name}: {len(filas)} filas (día 2)")


# --- 6. Modificaciones (día 2) ---------------------------------------------------------------------

def modificaciones() -> None:
    """Las mismas 34 filas de automatizacion/modificaciones_funcional.xlsx, pero sobre tarifas de la
    serie (FN-013 / FN-014, empaquetados y LyC que no usan las acciones), para que cada serie tenga su
    propio día 2. Las fechas se calculan desde hoy."""
    import re
    origen = ORIGEN / "modificaciones_funcional.xlsx"
    destino = SALIDA / "modificaciones_funcional.xlsx"
    libro = copiar_libro(origen, destino)
    usos: dict[str, int] = {}
    mapa: dict[str, str] = {}
    fijos = {"TE-EMP-CM-113": "TE-EMP-FN-106", "TE-EMP-CM-117": "TE-EMP-FN-107", "TE-EMP-CM-126": "TE-EMP-FN-102",
             "TL-LYC-CM-108": "TL-LYC-FN-103", "TL-LYC-CM-123": "TL-LYC-FN-104"}

    def nuevo(nombre: str, omitida: bool = False) -> str:
        if nombre in mapa:
            return mapa[nombre]
        if nombre in fijos:
            mapa[nombre] = fijos[nombre].replace("-FN-", f"-{SERIE}-")
            return mapa[nombre]
        m = re.fullmatch(r"TE-([A-Z]+)-CM-\d+", nombre)
        if not m:
            return nombre                                  # TE-PRUEBA-ERR-… (no tocan SIRT)
        codigo = m.group(1)
        if omitida:
            mapa[nombre] = fn(codigo, 99)                   # no se ejecuta: no necesita tarifa real
            return mapa[nombre]
        usos[codigo] = usos.get(codigo, 0) + 1
        mapa[nombre] = fn(codigo, 12 + usos[codigo])        # 1.ª → FN-013, 2.ª → FN-014
        return mapa[nombre]

    ws = libro["Modificaciones"]
    col = cabecera(ws)
    c_fecha = col.get("Cambio: fecha de ejecución")
    for r in range(xc.FILA_CABECERA + 1, ws.max_row + 1):
        nombre = xc.texto(ws.cell(r, col["Nombre de la tarifa"]).value)
        if not nombre:
            continue
        omitida = xc.texto(ws.cell(r, col["Ejecutar"]).value).lower() == "no"
        ws.cell(r, col["Nombre de la tarifa"]).value = nuevo(nombre, omitida)
        for titulo in ("Estado", "Código SIRT", "Fecha de ejecución", "Detalle", "ID de registro"):
            if titulo in col:
                ws.cell(r, col[titulo]).value = None
        if c_fecha and isinstance(ws.cell(r, c_fecha).value, (date,)):
            error = xc.texto(ws.cell(r, col["Qué se prueba"]).value).startswith("ERROR")
            ws.cell(r, c_fecha).value = date.today() if error else date.today() + timedelta(days=7)
    hc = libro["Campos"]
    cc = cabecera(hc)
    for r in range(xc.FILA_CABECERA + 1, hc.max_row + 1):
        nombre = xc.texto(hc.cell(r, cc["Nombre de la tarifa"]).value)
        if nombre:
            hc.cell(r, cc["Nombre de la tarifa"]).value = nuevo(nombre)
    libro.save(destino)
    print(f"  {destino.name}: tarifas de la serie (día 2)")


def main() -> None:
    import argparse
    import re
    from datetime import datetime
    global SERIE, SALIDA
    p = argparse.ArgumentParser(description="Excel de la prueba funcional (días 1 y 2) de una serie.")
    p.add_argument("--serie", default="FN",
                   help="parte del nombre de las tarifas (por defecto FN → TE-TF-FN-011). Para repetir la "
                        "prueba con tarifas nuevas: --serie FA, FB, F2…")
    p.add_argument("--carpeta", help="crea los Excel en pruebas_funcionales/<carpeta> (p. ej. Prueba4)")
    p.add_argument("--solo-dia2", action="store_true",
                   help="solo crea los Excel del día 2 (acciones y modificaciones)")
    args = p.parse_args()
    serie = args.serie.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{2,4}", serie):
        sys.exit("--serie: de 2 a 4 letras o números (p. ej. FA)")
    SERIE = serie
    if args.carpeta:
        if not re.fullmatch(r"[\w\- ]{1,40}", args.carpeta):
            sys.exit("--carpeta: solo letras, números, espacios, - y _")
        SALIDA = CARPETA / args.carpeta
    SALIDA.mkdir(parents=True, exist_ok=True)
    existentes = list(SALIDA.glob("*_funcional.xlsx"))
    if any(xc.excel_esta_abierto(f) for f in existentes):
        sys.exit(f"Cierra los Excel de {SALIDA} antes de generar.")
    if existentes:                                   # se guardan los anteriores (con sus resultados)
        respaldo = CARPETA / "anteriores" / datetime.now().strftime("%Y%m%d_%H%M%S")
        respaldo.mkdir(parents=True, exist_ok=True)
        for f in existentes:
            shutil.copyfile(f, respaldo / f.name)
        print(f"Excel anteriores (con sus resultados) guardados en {respaldo}")
    print(f"Generando en {SALIDA} (serie {SERIE}: TE-TF-{SERIE}-011…):")
    if not args.solo_dia2:
        valores = registro()
        empaquetados(valores)
        licitaciones()
        promociones()
    acciones()
    modificaciones()


if __name__ == "__main__":
    main()
