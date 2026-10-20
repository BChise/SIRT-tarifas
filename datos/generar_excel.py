"""
Genera datos/pruebas.xlsx: una hoja por servicio con sus tarifas de prueba.

Cada hoja tiene dos bloques:
  1. «Tarifas originales»: las tarifas a registrar, con TODOS sus atributos (los campos
     que la herramienta llena en SIRT, pasos 1 y 2) como columnas.
  2. «Tarifas con el cambio»: las mismas tarifas tal como deben quedar después de su
     acción (renombre, cese, fe de erratas, rectificación o cambio). Las celdas que
     cambian van resaltadas.

Uso (desde la carpeta del proyecto):
    python datos/generar_excel.py              → datos/pruebas.xlsx (7 tarifas por servicio)

La app web lee la hoja de cada servicio (bloque 1) para registrar.
"""

import importlib
import re
import sys
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))

import comun  # noqa: E402
from app import catalogo  # noqa: E402

LOTE = "DEMO"
SALIDA = PROYECTO / "datos" / "pruebas.xlsx"

# Escenario por tarifa: (acción, parámetros, alcance)
PLAN = [
    ("Ninguna", "", "Residencial"),
    ("Renombre", "sufijo=_REN", "Residencial"),
    ("Cese", "fecha=mañana", "Residencial"),
    ("Fe de erratas", "", "Comercial"),
    ("Rectificación", "delta=1", "Comercial"),
    ("Cambio", "delta=-1", "Residencial + Comercial"),
    ("Cambio", "delta=1", "Residencial + Comercial"),
]
# Empaquetados: grupos de 4 servicios mezclados por tipo (fija, internet/TV, móvil, otros),
# con la tarifa 001 (testigo, la que no cambia) de cada servicio.
GRUPOS = {
    "G1": ["TF", "IF", "TVP", "SMP"],
    "G2": ["TFP", "PIF", "TVPR", "SMC"],
    "G3": ["OTF", "PTV", "SMO", "OS"],
    "G4": ["PTF", "SMQ", "AC"],
}
# 4 paquetes testigo + 1 paquete por acción.
PAQUETES = [
    ("G1", "Ninguna", ""), ("G2", "Ninguna", ""), ("G3", "Ninguna", ""), ("G4", "Ninguna", ""),
    ("G1", "Renombre", "sufijo=_REN"), ("G2", "Cese", "fecha=mañana"), ("G3", "Fe de erratas", ""),
    ("G4", "Rectificación", "delta=1"), ("G1", "Cambio", "delta=-1"), ("G2", "Cambio", "delta=1"),
]
SIN_PROMOCION = {"OS"}          # su catálogo no tiene atributos promocionables
META = ["N°", "Nombre de la tarifa", "Lote", "Registro definitivo", "Acción", "Parámetros de la acción"]

ETIQUETAS = {
    "alcance": "Alcance", "moneda": "Tipo de moneda", "periodicidad": "Periodicidad",
    "tipo_acceso": "Tipo de acceso", "tecnologia": "Tecnología", "componentes": "Componentes del plan",
    "internet_ilimitado": "Internet Ilimitado", "tipo_linea": "Tipo de línea",
    "tipo_linea_movil": "Tipo de línea móvil", "tipo_contratacion": "Tipo de Contratación",
    "tecnologia_if": "Tecnología (IF)", "valor_servicio": "Valor del servicio",
    "duracion_dias": "Duración (días)", "decodificadores": "Decodificadores o similares",
    "tarifa_por_minuto": "Tarifa por minuto", "tarifa_mb_adicional": "Tarifa por MB adicional",
    "tarifa_movil_a_fijo": "Tarifa minuto móvil a fijo",
    "tarifa_movil_todo_destino": "Tarifa minuto móvil todo destino",
    "datos_tope_descarga_gb": "Datos tope descarga (GB)", "datos_tethering_gb": "Datos para tethering (GB)",
    "traslada_datos_gb": "Traslada datos (GB)",
    "datos_roaming_incluidos_gb": "Datos Roaming incluidos en los datos del plan (GB)",
    "datos_roaming_adicionales_gb": "Datos Roaming adicionales a los datos del plan (GB)",
    "cantidad_datos_gb": "Cantidad de datos (GB)", "repetidores": "Repetidores y/o equipos similares",
    "periodo_contrato_meses": "Periodo de contrato (meses)",
    "tecnologia_transmision": "Tecnología de transmisión",
    "velocidad_transmision": "Velocidad de transmisión", "velocidad_garantizada": "Vel. transmisión garant.",
    "condiciones_generales": "Condiciones generales", "restricciones_tecnicas": "Restricciones técnicas",
    "otras_condiciones": "Otras condiciones tarifa", "otras_restricciones": "Otras restricciones",
    "ldi_fijo": "Tarifas fijo LDI (tabla)", "ldi_movil": "Tarifas móvil LDI (tabla)",
    "velocidades": "Tecnologías y velocidades (tabla)", "apps_sin_consumo": "Apps sin consumo de datos",
}


def etiqueta(clave: str) -> str:
    if clave in ETIQUETAS:
        return ETIQUETAS[clave]
    texto = clave.replace("_gb", " (GB)").replace("_", " ")
    return texto[:1].upper() + texto[1:]


def formato(valor) -> str:
    """Texto legible para una celda."""
    if valor is None:
        return ""
    if isinstance(valor, Path):
        return valor.name
    if isinstance(valor, (tuple, set)):
        return " + ".join(formato(v) for v in valor)
    if isinstance(valor, dict):
        if "modo" in valor:                                   # condicional triple
            if valor["modo"] == "ilimitado":
                return "Ilimitado"
            partes = [f"{k}: {v}" for k, v in valor.items() if k != "modo"]
            return "Limitado · " + " · ".join(partes)
        return "; ".join(f"{k}: {formato(v)}" for k, v in valor.items())
    if isinstance(valor, list):                               # tabla: filas
        if not valor:
            return "(sin filas)"
        filas = []
        for fila in valor:
            if isinstance(fila, dict):
                filas.append(", ".join(f"{k}={formato(v)}" for k, v in fila.items()))
            else:
                filas.append(formato(fila))
        return f"{len(valor)} fila(s): " + " | ".join(filas)
    return str(valor)


def atributos(modulo, caso: dict) -> dict:
    """Atributo → valor (texto), en el orden en que se llenan: paso 1 y paso 2."""
    paso1 = {**comun.PASO1_POR_DEFECTO, **getattr(modulo, "PASO1", {})}
    salida = {}
    for clave, valor in caso["valores"].items():              # listas que se combinan
        salida[etiqueta(clave)] = formato(valor)
    if modulo.CODIGO in catalogo.CON_INSTALACION:
        salida["Tarifa de instalación"] = formato(paso1.get("tarifa_instalacion"))
    salida["Otros datos generales"] = formato(paso1.get("otros_datos"))
    for clave, valor in modulo.VALORES.items():               # valores fijos del paso 2
        if isinstance(valor, dict) and valor and all(isinstance(k, str) and k[:1].isupper() for k in valor):
            for sub, v in valor.items():                      # p. ej. textos, condicionales, bolsas…
                salida[sub] = formato(v)
        else:
            salida[etiqueta(clave)] = formato(valor)
    return salida


TEXTOS = re.compile(r"^(otros datos|condiciones|restricciones|otras condiciones|otras restricciones)", re.I)
NUMERO = re.compile(r"^-?\d+(\.\d+)?$")


def _sumar(texto: str, delta: int, por_decimal: bool) -> str:
    decimales = len(texto.split(".")[1]) if "." in texto else 0
    paso = Decimal(delta) * (Decimal(1).scaleb(-decimales) if por_decimal else 1)
    nuevo = Decimal(texto) + paso
    return f"{nuevo:.{decimales}f}" if nuevo >= 0 else texto


def con_cambio(codigo: str, nombre: str, attrs: dict, accion: str, parametros: str) -> tuple[str, dict, dict, str]:
    """(nombre resultante, atributos resultantes, columnas extra, descripción del cambio)."""
    nuevo = dict(attrs)
    extra: dict[str, str] = {}
    hoy = date.today()
    delta = int(re.search(r"delta=(-?\d+)", parametros).group(1)) if "delta=" in parametros else 0
    if accion == "Ninguna":
        return nombre, nuevo, extra, "Sin cambios (testigo)"
    if accion == "Renombre":
        sufijo = re.search(r"sufijo=(\S+)", parametros)
        return nombre + (sufijo.group(1) if sufijo else "_REN"), nuevo, extra, "Cambia solo el nombre"
    if accion == "Cese":
        extra["Fin de comercialización"] = (hoy + timedelta(days=1)).strftime("%d/%m/%Y")
        return nombre, nuevo, extra, "Fin de comercialización mañana"
    sello = {"Fe de erratas": "Fe de erratas", "Rectificación": "Rectificación", "Cambio": "Cambio"}[accion]
    excluidos = {catalogo._sin_tildes(n) for n in catalogo.CAMBIO_SIN.get(codigo, [])} if accion == "Cambio" else set()
    for campo, valor in attrs.items():
        base = catalogo._sin_tildes(re.sub(r"\s*\(.*?\)\s*$", "", campo))
        if base in excluidos or campo == "Alcance":
            continue
        if TEXTOS.match(campo) and valor:
            nuevo[campo] = f"{valor} | {sello} <fecha hora>"
        elif accion != "Fe de erratas" and NUMERO.match(valor or "") and Decimal(valor) != 0:
            principal = campo in ("Valor del servicio", "Tarifa de instalación")
            nuevo[campo] = _sumar(valor, delta, por_decimal=not principal)
    if accion == "Cambio":
        dias = 2 if delta > 0 else 1
        extra["Fecha de ejecución"] = (hoy + timedelta(days=dias)).strftime("%d/%m/%Y")
        descripcion = f"{'Incremento' if delta > 0 else 'Reducción'} de 1 en valores; textos con sello"
    elif accion == "Rectificación":
        descripcion = "Valores +1 y textos con sello"
    else:
        descripcion = "Solo textos (sello «Fe de erratas»)"
    return nombre, nuevo, extra, descripcion


def generar() -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    F = "Arial"
    borde = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=borde, right=borde, top=borde, bottom=borde)
    CAB = PatternFill("solid", fgColor="1F4E78")
    CAB2 = PatternFill("solid", fgColor="7F6000")
    EDIT = PatternFill("solid", fgColor="FFF2CC")
    CAMBIO = PatternFill("solid", fgColor="F8CBAD")

    wb = Workbook()
    leyenda = wb.active
    leyenda.title = "Leyenda"
    resumen = []

    for servicio in catalogo.servicios():
        modulo = importlib.import_module(f"servicios.{servicio['archivo']}")
        casos = comun.generar_casos(modulo.DIMENSIONES, modulo.CODIGO, LOTE)
        filas = []
        for k, (accion, parametros, alcance) in enumerate(PLAN, start=1):
            buscado = {a.strip().lower() for a in alcance.split("+")}
            caso = next((c for c in casos if {a.strip().lower() for a in
                         comun.texto_valor(c["valores"].get("alcance", "")).split("+")} == buscado), casos[0])
            nombre = f"TE-{modulo.CODIGO}-{LOTE}-{k:03d}"
            attrs = atributos(modulo, caso)
            filas.append((k, nombre, accion, parametros, attrs, con_cambio(modulo.CODIGO, nombre, attrs, accion, parametros)))
        columnas = list(filas[0][4])
        extras = sorted({c for f in filas for c in f[5][2]})

        hoja = wb.create_sheet(modulo.CODIGO)
        hoja["A1"] = f"{modulo.CODIGO} · {servicio['nombre']}"
        hoja["A1"].font = Font(name=F, bold=True, size=14)
        hoja["A2"] = ("Bloque 1: tarifas a registrar con sus atributos (los campos que llena la herramienta). "
                      "Bloque 2: cómo deben quedar después de su acción; en naranja lo que cambia.")
        hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")

        # Bloque 1
        hoja["A4"] = "TARIFAS ORIGINALES"
        hoja["A4"].font = Font(name=F, bold=True, size=11, color="1F4E78")
        cabecera = META + columnas
        for j, titulo in enumerate(cabecera, start=1):
            c = hoja.cell(5, j, titulo)
            c.font = Font(name=F, bold=True, color="FFFFFF", size=9)
            c.fill, c.border = CAB, BORDE
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, (k, nombre, accion, parametros, attrs, _) in enumerate(filas, start=6):
            valores = [k, nombre, LOTE, "Sí", accion, parametros] + [attrs.get(c, "") for c in columnas]
            for j, v in enumerate(valores, start=1):
                c = hoja.cell(i, j, v)
                c.font = Font(name=F, size=9, color="0000FF" if 2 <= j <= 6 else "000000")
                c.border = BORDE
                c.alignment = Alignment(vertical="top", wrap_text=j > 6)
                if 2 <= j <= 6:
                    c.fill = EDIT
        fin1 = 5 + len(filas)
        for rango, opciones in ((f"D6:D{fin1 + 20}", '"Sí,No"'),
                                (f"E6:E{fin1 + 20}", '"Ninguna,Renombre,Cese,Fe de erratas,Rectificación,Cambio"')):
            dv = DataValidation(type="list", formula1=opciones, allow_blank=True)
            hoja.add_data_validation(dv)
            dv.add(rango)

        # Bloque 2
        inicio2 = fin1 + 3
        hoja.cell(inicio2, 1, "TARIFAS CON EL CAMBIO (resultado esperado de cada acción)").font = \
            Font(name=F, bold=True, size=11, color="7F6000")
        cabecera2 = ["N°", "Nombre resultante", "Acción", "Cambio aplicado"] + columnas + extras
        for j, titulo in enumerate(cabecera2, start=1):
            c = hoja.cell(inicio2 + 1, j, titulo)
            c.font = Font(name=F, bold=True, color="FFFFFF", size=9)
            c.fill, c.border = CAB2, BORDE
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, (k, nombre, accion, parametros, attrs, cambio) in enumerate(filas, start=inicio2 + 2):
            nombre2, attrs2, extra2, descripcion = cambio
            valores = [k, nombre2, accion, descripcion] + [attrs2.get(c, "") for c in columnas] + [extra2.get(c, "") for c in extras]
            originales = [k, nombre, accion, descripcion] + [attrs.get(c, "") for c in columnas] + ["" for _ in extras]
            for j, (v, antes) in enumerate(zip(valores, originales), start=1):
                c = hoja.cell(i, j, v)
                c.font = Font(name=F, size=9)
                c.border = BORDE
                c.alignment = Alignment(vertical="top", wrap_text=j > 4)
                if j != 4 and v != antes and v != "":
                    c.fill = CAMBIO
        anchos = [5, 24, 8, 11, 14, 18] + [max(14, min(36, len(c) + 4)) for c in columnas] + [16] * len(extras)
        for j, ancho in enumerate(anchos, start=1):
            hoja.column_dimensions[get_column_letter(j)].width = ancho
        hoja.freeze_panes = "C6"
        resumen.append((modulo.CODIGO, servicio["nombre"], len(filas), len(columnas)))

    # --- Hoja EMP: empaquetados -------------------------------------------------
    nombre_servicio = {s["codigo"]: s["nombre"] for s in catalogo.servicios()}
    hoja = wb.create_sheet("EMP", 1)
    hoja["A1"] = "EMP · Tarifas empaquetadas"
    hoja["A1"].font = Font(name=F, bold=True, size=14)
    hoja["A2"] = ("Grupos de 3-4 servicios con la tarifa 001 (testigo) de cada uno. Los paquetes 001-004 son testigos; "
                  "del 005 en adelante, cada uno recibe una acción. Valor de cada servicio: su tarifa base − 1.")
    hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    cab = ["N°", "Nombre del paquete", "Grupo", "Tarifas del paquete", "Servicios", "Lote",
           "Registro definitivo", "Acción", "Parámetros de la acción", "Resultado esperado"]
    for j, t in enumerate(cab, start=1):
        c = hoja.cell(4, j, t)
        c.font, c.fill, c.border = Font(name=F, bold=True, color="FFFFFF", size=9), CAB, BORDE
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, (grupo, accion, parametros) in enumerate(PAQUETES, start=1):
        codigos = GRUPOS[grupo]
        tarifas = "; ".join(f"TE-{c}-{LOTE}-001" for c in codigos)
        servicios = " + ".join(codigos)
        _, _, _, esperado = con_cambio("EMP", "", {}, accion, parametros)
        if accion == "Renombre":
            esperado = f"Nombre: TE-EMP-{LOTE}-{i:03d}_REN"
        fila = [i, f"TE-EMP-{LOTE}-{i:03d}", grupo, tarifas, servicios, LOTE, "Sí", accion, parametros, esperado]
        for j, v in enumerate(fila, start=1):
            c = hoja.cell(4 + i, j, v)
            c.font = Font(name=F, size=9, color="0000FF" if j in (2, 4, 6, 7, 8, 9) else "000000")
            c.border = BORDE
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if j in (2, 4, 6, 7, 8, 9):
                c.fill = EDIT
    for col, ancho in zip("ABCDEFGHIJ", [5, 20, 7, 58, 22, 8, 11, 14, 18, 40]):
        hoja.column_dimensions[col].width = ancho
    hoja.freeze_panes = "C5"

    # --- Hoja PROMO: promociones ------------------------------------------------
    hoja = wb.create_sheet("PROMO", 2)
    hoja["A1"] = "PROMO · Tarifas promocionales"
    hoja["A1"].font = Font(name=F, bold=True, size=14)
    hoja["A2"] = ("Una promoción por la tarifa 001 de cada servicio (OS no tiene atributos promocionables) y una por cada "
                  "paquete testigo (001-004), con TODOS los atributos posibles de sus servicios. Registro definitivo sin disponibilidad.")
    hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    cab = ["N°", "Nombre de la promoción", "Tarifas base", "Servicios", "Lote", "Registro definitivo",
           "Atributos", "Variación %", "Variación cantidad", "Duración (días)"]
    for j, t in enumerate(cab, start=1):
        c = hoja.cell(4, j, t)
        c.font, c.fill, c.border = Font(name=F, bold=True, color="FFFFFF", size=9), CAB, BORDE
        c.alignment = Alignment(wrap_text=True, vertical="center")
    promos = []
    for s in catalogo.servicios():
        if s["codigo"] not in SIN_PROMOCION:
            promos.append((f"TP-{s['codigo']}-{LOTE}-001", f"TE-{s['codigo']}-{LOTE}-001", s["codigo"]))
    for i, (grupo, accion, _) in enumerate(PAQUETES, start=1):
        if accion == "Ninguna":
            codigos = [c for c in GRUPOS[grupo] if c not in SIN_PROMOCION]
            promos.append((f"TP-EMP-{LOTE}-{i:03d}", f"TE-EMP-{LOTE}-{i:03d}", "EMP (" + " + ".join(codigos) + ")"))
    for i, (nombre, base, servicios) in enumerate(promos, start=1):
        fila = [i, nombre, base, servicios, LOTE, "Sí", "todos", "10", "1", "30"]
        for j, v in enumerate(fila, start=1):
            c = hoja.cell(4 + i, j, v)
            c.font = Font(name=F, size=9, color="0000FF" if j != 1 and j != 4 else "000000")
            c.border = BORDE
            if j not in (1, 4):
                c.fill = EDIT
    for col, ancho in zip("ABCDEFGHIJ", [5, 20, 20, 26, 8, 11, 10, 10, 12, 12]):
        hoja.column_dimensions[col].width = ancho
    hoja.freeze_panes = "C5"

    textos = [
        ("Escenarios de prueba de la herramienta SIRT", True, 14),
        ("", False, 10),
        ("Una hoja por servicio (código en la pestaña). Cada hoja tiene dos bloques:", False, 10),
        ("1. TARIFAS ORIGINALES: las tarifas a registrar. Columnas amarillas = datos del escenario (nombre, lote,", False, 10),
        ("   registro definitivo, acción y parámetros). El resto de columnas son los ATRIBUTOS: los campos que la", False, 10),
        ("   herramienta llena en SIRT en los pasos 1 y 2 (listas, valores, textos, condicionales y tablas).", False, 10),
        ("2. TARIFAS CON EL CAMBIO: cómo debe quedar cada tarifa después de su acción. En naranja, lo que cambia.", False, 10),
        ("", False, 10),
        ("Acciones: Ninguna (testigo) · Renombre · Cese · Fe de erratas · Rectificación · Cambio.", False, 10),
        ("Parámetros: sufijo=_REN · fecha=mañana · delta=1 (sube) / delta=-1 (baja).", False, 10),
        ("Fe de erratas cambia solo textos; rectificación y cambio también los números (+1 / −1 en la última cifra).", False, 10),
        ("Cambio no toca los campos que SIRT no muestra en esa acción; su fecha de ejecución es mañana (baja) u hoy + 2 (sube).", False, 10),
        ("Las acciones exigen que la tarifa esté vigente: el día siguiente a su registro definitivo.", False, 10),
        ("Hoja EMP: empaquetados (grupos de 3-4 servicios con la tarifa 001). Hoja PROMO: promociones de monoproductos y paquetes.", False, 10),
        ("", False, 10),
        ("Servicios", True, 11),
    ]
    for i, (t, negrita, tam) in enumerate(textos, start=1):
        leyenda.cell(i, 1, t).font = Font(name=F, bold=negrita, size=tam)
    fila = len(textos) + 1
    for j, titulo in enumerate(["Código", "Servicio", "Tarifas", "Atributos"], start=1):
        c = leyenda.cell(fila, j, titulo)
        c.font, c.fill = Font(name=F, bold=True, color="FFFFFF"), CAB
    for codigo, nombre, n, a in resumen:
        fila += 1
        for j, v in enumerate([codigo, nombre, n, a], start=1):
            leyenda.cell(fila, j, v).font = Font(name=F, size=10)
    leyenda.column_dimensions["A"].width = 12
    leyenda.column_dimensions["B"].width = 60
    SALIDA.parent.mkdir(exist_ok=True)
    wb.save(SALIDA)
    print(f"Generado: {SALIDA}  ({len(resumen)} hojas de servicio)")


if __name__ == "__main__":
    generar()
