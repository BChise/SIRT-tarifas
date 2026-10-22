"""
Genera automatizacion/promociones.xlsx: el Excel para registrar TARIFAS PROMOCIONALES de
monoproductos y/o empaquetados (lo lee promociones_excel.py). Hasta 15 tarifas por promoción.

Hojas:
- Promociones: una fila por promoción (tarifas base, comercialización, valor del plan,
  tarifa de instalación y qué atributos promocionar con sus valores por defecto).
- Atributos (opcional): detalle por atributo cuando se quiere controlar cada uno
  (tipo de variación, variación y duración). Solo para filas con «Detalle en hoja Atributos».
- Catálogo de atributos: consulta. Qué atributos se pueden promocionar en cada servicio y
  con qué tipos de variación (del JSON del paso 2 de cada servicio en SIRT).

Uso (desde la carpeta del proyecto):
    python automatizacion/generar_plantilla_promociones.py
"""

import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))

SALIDA = PROYECTO / "automatizacion" / "promociones.xlsx"
FILAS_CON_LISTAS = 1000
MAX_TARIFAS = 15
MODOS = ["Los de siempre", "Todos", "Lista", "Detalle en hoja Atributos"]
TIPOS_VARIACION = ["Porcentual", "Cantidad", "Ilimitado"]

# (título, tipo, lista, ancho, sección)
COLUMNAS = [
    ("N°", "meta", None, 5, ""), ("Ejecutar", "meta", "Sí/No", 9, ""),
    ("Nombre de la promoción", "meta", None, 20, ""), ("Registro definitivo", "meta", "Sí/No", 11, ""),
] + [(f"Tarifa base {k}", "obligatorio" if k == 1 else "opcional", None, 20,
      "Tarifas a promocionar (monoproducto o empaquetado)") for k in range(1, MAX_TARIFAS + 1)] + [
    ("Inicio comercialización", "obligatorio", None, 13, "Datos generales"),
    ("Días de comercialización", "obligatorio", None, 12, "Datos generales"),
    ("Condiciones generales", "obligatorio", None, 26, "Datos generales"),
    ("Otras condiciones tarifa", "opcional", None, 22, "Datos generales"),
    ("Valor del plan: promocionar", "obligatorio", "Sí/No", 12, "Valor del plan"),
    ("Valor del plan: tipo de variación", "opcional", "Tipo plan", 13, "Valor del plan"),
    ("Valor del plan: variación", "opcional", None, 10, "Valor del plan"),
    ("Valor del plan: duración (días)", "opcional", None, 11, "Valor del plan"),
    ("Tarifa instalación: promocionar", "opcional", "Sí/No", 12, "Tarifa de instalación (si el servicio la tiene)"),
    ("Tarifa instalación: tipo de variación", "opcional", "Tipo plan", 13, "Tarifa de instalación (si el servicio la tiene)"),
    ("Tarifa instalación: variación", "opcional", None, 10, "Tarifa de instalación (si el servicio la tiene)"),
    ("Atributos a promocionar", "obligatorio", "Modo", 18, "Atributos del servicio (paso 2)"),
    ("Lista de atributos (separados por ;)", "opcional", None, 32, "Atributos del servicio (paso 2)"),
    ("Tipo de variación preferido", "opcional", "Tipo", 13, "Atributos del servicio (paso 2)"),
    ("Variación % por defecto", "opcional", None, 10, "Atributos del servicio (paso 2)"),
    ("Variación en cantidad por defecto", "opcional", None, 11, "Atributos del servicio (paso 2)"),
    ("Duración por defecto (días)", "opcional", None, 10, "Atributos del servicio (paso 2)"),
    ("Estado", "script", None, 13, "Resultado"), ("Código SIRT", "script", None, 18, "Resultado"),
    ("Fecha de registro", "script", None, 17, "Resultado"), ("Detalle", "script", None, 50, "Resultado"),
]

MUESTRAS = [
    {"nombre": "TP-IF-CM-001", "tarifas": ["TE-IF-CM-003"], "modo": "Los de siempre",
     "condiciones": "Promoción de velocidad por 30 días", "plan": ("Sí", "Porcentual", "10", "30"),
     "instalacion": ("Sí", "Porcentual", "10")},
    {"nombre": "TP-SMP-CM-001", "tarifas": ["TE-SMP-CM-001", "TE-SMP-CM-002"], "modo": "Todos",
     "condiciones": "Promoción de datos para dos planes", "plan": ("Sí", "Porcentual", "10", "30"),
     "instalacion": ("", "", "")},
    {"nombre": "TP-EMP-CM-001", "tarifas": ["TE-EMP-CM-101"], "modo": "Lista",
     "lista": "Minutos todo destino nacional; Datos para tethering",
     "condiciones": "Promoción del dúo fijo + móvil", "plan": ("No", "", "", ""), "instalacion": ("", "", "")},
    {"nombre": "TP-TF-CM-001", "tarifas": ["TE-TF-CM-001"], "modo": "Detalle en hoja Atributos",
     "condiciones": "Promoción de minutos con valores a medida", "plan": ("Sí", "Cantidad", "5", "60"),
     "instalacion": ("", "", "")},
]
DETALLE_ATRIBUTOS = [  # (promoción, servicio, atributo, tipo, variación, duración)
    ("TP-TF-CM-001", "TF", "Minutos todo destino nacional", "Cantidad", "100", "60"),
    ("TP-TF-CM-001", "TF", "Minutos red fija on net", "Ilimitado", "", "60"),
    ("TP-TF-CM-001", "TF", "Minutos móvil off net", "Porcentual", "20", "30"),
]


def catalogo() -> list[tuple]:
    """(código, servicio, atributo, tipos de variación) de cada servicio."""
    import promocionar
    from app import catalogo as cat
    salida = []
    for s in cat.servicios():
        lista = promocionar.catalogo_atributos(s["archivo"])
        if lista is None:
            lista = [{"nombre": n, "opciones": []} for n in promocionar.ATRIBUTOS_PROBADOS.get(s["codigo"], [])]
        for a in lista:
            tipos = " / ".join(a.get("opciones") or []) or ("Archivo Excel (canales)" if "Canales" in a["nombre"] else "—")
            salida.append((s["codigo"], s["nombre"], a["nombre"], tipos))
        if not lista:
            salida.append((s["codigo"], s["nombre"], "(sin atributos promocionables: solo valor del plan e instalación)", "—"))
    return salida


def generar() -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    F = "Arial"
    lado = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=lado, right=lado, top=lado, bottom=lado)
    RELLENO = {"meta": ("375623", "FFFFFF"), "obligatorio": ("1F4E78", "FFFFFF"), "opcional": ("BDD7EE", "1F1F1F"),
               "script": ("808080", "FFFFFF")}
    CELDA_SCRIPT = PatternFill("solid", fgColor="EDEDED")
    SECCION = PatternFill("solid", fgColor="DDEBF7")

    def cabeza(c, tipo):
        fondo, color = RELLENO[tipo]
        c.fill, c.border = PatternFill("solid", fgColor=fondo), BORDE
        c.font = Font(name=F, bold=True, size=9, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    def celda(hoja, f, c, v, centro=False):
        x = hoja.cell(f, c, v)
        x.font, x.border = Font(name=F, size=9), BORDE
        x.alignment = Alignment(vertical="top", wrap_text=True, horizontal="center" if centro else None)
        return x

    wb = Workbook()
    guia = wb.active
    guia.title = "Instrucciones"
    hoja = wb.create_sheet("Promociones")
    hoja_attr = wb.create_sheet("Atributos")
    hoja_cat = wb.create_sheet("Catálogo de atributos")
    listas = wb.create_sheet("Listas")
    listas.sheet_state = "hidden"

    cat = catalogo()
    codigos = sorted({c[0] for c in cat})
    atributos = sorted({c[2] for c in cat if not c[2].startswith("(")})
    opciones = {"Sí/No": ["Sí", "No"], "Modo": MODOS, "Tipo plan": ["Porcentual", "Cantidad"],
                "Tipo": TIPOS_VARIACION, "Servicio": codigos, "Atributo": atributos}
    rangos = {}
    for k, (nombre, valores) in enumerate(opciones.items(), start=1):
        letra = get_column_letter(k)
        for f, v in enumerate(valores, start=1):
            listas[f"{letra}{f}"] = v
        rangos[nombre] = f"=Listas!${letra}$1:${letra}${len(valores)}"

    def lista(h, col, nombre, aviso=False):
        dv = DataValidation(type="list", formula1=rangos[nombre], allow_blank=True)
        dv.showErrorMessage = True
        if aviso:
            dv.errorStyle, dv.error = "warning", "No está en el catálogo; revisa el nombre exacto en «Catálogo de atributos»."
        else:
            dv.error = "Elige una opción de la lista."
        h.add_data_validation(dv)
        letra = get_column_letter(col)
        dv.add(f"{letra}6:{letra}{5 + FILAS_CON_LISTAS}")

    # --- Hoja Promociones ------------------------------------------------------------
    hoja["A1"] = "Tarifas promocionales"
    hoja["A1"].font = Font(name=F, bold=True, size=14)
    hoja["A2"] = ("Una fila por promoción. Tarifas base: nombre exacto de tarifas (monoproducto o empaquetado) con "
                  "registro definitivo. Qué hace cada columna: hoja «Instrucciones».")
    hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    for j, (texto, clave) in enumerate([("Obligatorio", "obligatorio"), ("Opcional", "opcional"),
                                        ("Lo llena el script", "script")], start=4):
        cabeza(hoja.cell(3, j, texto), clave)
    secciones = [c[4] for c in COLUMNAS]
    for j, s in enumerate(secciones, start=1):
        c = hoja.cell(4, j, s if (s and (j == 1 or s != secciones[j - 2])) else None)
        c.fill, c.border = SECCION, BORDE
        c.font = Font(name=F, bold=True, size=8, color="1F4E78")
    col = {}
    for j, (titulo, tipo, nombre_lista, ancho, _) in enumerate(COLUMNAS, start=1):
        cabeza(hoja.cell(5, j, titulo), tipo)
        hoja.column_dimensions[get_column_letter(j)].width = ancho
        col[titulo] = j
        if nombre_lista:
            lista(hoja, j, nombre_lista)
    hoja.row_dimensions[5].height = 54

    for i, m in enumerate(MUESTRAS, start=6):
        v = {"N°": i - 5, "Ejecutar": "Sí", "Nombre de la promoción": m["nombre"], "Registro definitivo": "Sí",
             "Inicio comercialización": "", "Días de comercialización": 60,
             "Condiciones generales": m["condiciones"], "Otras condiciones tarifa": "",
             "Valor del plan: promocionar": m["plan"][0], "Valor del plan: tipo de variación": m["plan"][1],
             "Valor del plan: variación": m["plan"][2], "Valor del plan: duración (días)": m["plan"][3],
             "Tarifa instalación: promocionar": m["instalacion"][0],
             "Tarifa instalación: tipo de variación": m["instalacion"][1],
             "Tarifa instalación: variación": m["instalacion"][2],
             "Atributos a promocionar": m["modo"], "Lista de atributos (separados por ;)": m.get("lista", ""),
             "Variación % por defecto": "10", "Variación en cantidad por defecto": "1",
             "Duración por defecto (días)": "30"}
        for k, t in enumerate(m["tarifas"], start=1):
            v[f"Tarifa base {k}"] = t
        for titulo, (j) in col.items():
            x = celda(hoja, i, j, v.get(titulo, ""), centro=titulo in ("N°", "Ejecutar", "Registro definitivo"))
            if titulo in ("Estado", "Código SIRT", "Fecha de registro", "Detalle"):
                x.fill = CELDA_SCRIPT
    for i in range(6, 6 + FILAS_CON_LISTAS):
        hoja.cell(i, col["Inicio comercialización"]).number_format = "DD/MM/YYYY"
    for titulo, maximo in (("Valor del plan: duración (días)", 180), ("Duración por defecto (días)", 180)):
        dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(maximo), allow_blank=True)
        dv.showErrorMessage, dv.error = True, f"Duración en días: de 1 a {maximo} (regla de SIRT)."
        hoja.add_data_validation(dv)
        letra = get_column_letter(col[titulo])
        dv.add(f"{letra}6:{letra}{5 + FILAS_CON_LISTAS}")
    hoja.freeze_panes = "D6"
    hoja.auto_filter.ref = f"A5:{get_column_letter(len(COLUMNAS))}{5 + len(MUESTRAS)}"

    # --- Hoja Atributos (opcional) ---------------------------------------------------
    hoja_attr["A1"] = "Detalle por atributo (opcional)"
    hoja_attr["A1"].font = Font(name=F, bold=True, size=14)
    hoja_attr["A2"] = ("Solo para promociones con «Atributos a promocionar» = «Detalle en hoja Atributos». "
                       "Una fila por atributo; nombres exactos en «Catálogo de atributos».")
    hoja_attr["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    cab = [("Nombre de la promoción", "meta", 20), ("Servicio", "obligatorio", 10), ("Atributo", "obligatorio", 40),
           ("Tipo de variación", "obligatorio", 14), ("Variación", "opcional", 10), ("Duración (días)", "obligatorio", 11)]
    for j, (t, tipo, ancho) in enumerate(cab, start=1):
        cabeza(hoja_attr.cell(5, j, t), tipo)
        hoja_attr.column_dimensions[get_column_letter(j)].width = ancho
    for i, fila in enumerate(DETALLE_ATRIBUTOS, start=6):
        for j, v in enumerate(fila, start=1):
            celda(hoja_attr, i, j, v, centro=j in (2, 4, 5, 6))
    lista(hoja_attr, 2, "Servicio")
    lista(hoja_attr, 3, "Atributo", aviso=True)
    lista(hoja_attr, 4, "Tipo")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="180", allow_blank=True)
    dv.showErrorMessage, dv.error = True, "Duración en días: de 1 a 180 (regla de SIRT)."
    hoja_attr.add_data_validation(dv)
    dv.add(f"F6:F{5 + FILAS_CON_LISTAS}")
    hoja_attr.freeze_panes = "B6"

    # --- Hoja Catálogo (consulta) ----------------------------------------------------
    hoja_cat["A1"] = "Catálogo de atributos promocionables (consulta)"
    hoja_cat["A1"].font = Font(name=F, bold=True, size=14)
    hoja_cat["A2"] = "Del JSON del paso 2 de cada servicio en SIRT. Úsalo para escribir los nombres exactos."
    hoja_cat["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    for j, (t, ancho) in enumerate([("Código", 9), ("Servicio", 48), ("Atributo", 48),
                                    ("Tipos de variación permitidos", 30)], start=1):
        cabeza(hoja_cat.cell(4, j, t), "meta")
        hoja_cat.column_dimensions[get_column_letter(j)].width = ancho
    for i, fila in enumerate(cat, start=5):
        for j, v in enumerate(fila, start=1):
            celda(hoja_cat, i, j, v)
    hoja_cat.freeze_panes = "A5"
    hoja_cat.auto_filter.ref = f"A4:D{4 + len(cat)}"

    textos = [
        ("Registro de tarifas promocionales", True, 14),
        ("Ubicación: automatizacion/promociones.xlsx. Sirve para monoproductos y empaquetados.", False, 9),
        ("", False, 10),
        ("Cómo se usa", True, 11),
        ("1. Hoja «Promociones»: una fila por promoción. «Ejecutar» = No: la fila se deja sin tocar.", False, 10),
        ("2. Tarifas base: de 1 a 15, con nombre exacto y registro definitivo. Pueden ser monoproductos o empaquetados;", False, 10),
        ("   varias del mismo servicio también se pueden (p. ej. dos planes móviles en una sola promoción).", False, 10),
        ("3. Comercialización: «Inicio» vacío = día del registro; fin = inicio + «Días de comercialización».", False, 10),
        ("4. Valor del plan y tarifa de instalación: «promocionar» Sí/No, tipo de variación, variación y duración.", False, 10),
        ("   La tarifa de instalación solo aplica si alguno de los servicios la tiene.", False, 10),
        ("5. «Atributos a promocionar» (paso 2, una pestaña por servicio):", False, 10),
        ("   • Los de siempre: los atributos probados de cada servicio (los de las grabaciones).", False, 10),
        ("   • Todos: todos los promocionables del servicio (ver «Catálogo de atributos»).", False, 10),
        ("   • Lista: solo los escritos en «Lista de atributos», separados por « ; ».", False, 10),
        ("   • Detalle en hoja Atributos: cada atributo con su tipo, variación y duración propios (hoja «Atributos»).", False, 10),
        ("   En los tres primeros casos cada atributo usa los valores por defecto de la fila: variación % (si es", False, 10),
        ("   Porcentual), variación en cantidad (si es Cantidad) y duración. Si solo admite Ilimitado, no lleva variación.", False, 10),
        ("   «Tipo de variación preferido» (opcional): Porcentual, Cantidad o Ilimitado. Cada atributo usa ese tipo si", False, 10),
        ("   SIRT lo permite; si no, el orden normal Porcentual → Cantidad → Ilimitado. Vacío = orden normal.", False, 10),
        ("6. Guarda, cierra el Excel y corre: python promociones_excel.py (--solo-listar para revisar antes).", False, 10),
        ("   Las columnas grises mostrarán el resultado de cada promoción.", False, 10),
        ("", False, 10),
        ("Reglas de SIRT", True, 11),
        ("• Duraciones: de 1 a 180 días. • Las tarifas base deben tener registro definitivo.", False, 10),
        ("• Una promoción no se puede cesar (sí renovar y corregir con fe de erratas).", False, 10),
        ("• Registro definitivo de la promoción: sin disponibilidad comercial.", False, 10),
        ("", False, 10),
        ("Colores", True, 11),
        ("Verde oscuro: datos de la fila. Azul oscuro: obligatorio. Azul claro: opcional. Gris: lo llena el script.", False, 10),
    ]
    for i, (t, negrita, tam) in enumerate(textos, start=1):
        guia.cell(i, 1, t).font = Font(name=F, bold=negrita, size=tam)
    guia.column_dimensions["A"].width = 120

    SALIDA.parent.mkdir(exist_ok=True)
    wb.save(SALIDA)
    print(f"Generado: {SALIDA} ({len(MUESTRAS)} promociones de muestra, {len(cat)} atributos en el catálogo)")


if __name__ == "__main__":
    generar()
