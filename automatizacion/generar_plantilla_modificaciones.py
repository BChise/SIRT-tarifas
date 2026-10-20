"""
Genera automatizacion/modificaciones.xlsx (BORRADOR, sin conectar con los scripts): el
Excel para RECTIFICACIÓN y CAMBIO DE TARIFA de monoproductos, empaquetados y licitaciones.

Se indica QUÉ hacer (reglas), no el valor nuevo de cada campo: el script toma lo que
la tarifa tiene en SIRT y lo modifica según las opciones de la fila (lo mismo que hoy
hace acciones.py rectificar / cambio, ya probado).

Uso (desde la carpeta del proyecto):
    python automatizacion/generar_plantilla_modificaciones.py
"""

from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
SALIDA = PROYECTO / "automatizacion" / "modificaciones.xlsx"
FILAS_CON_LISTAS = 1000

ACCIONES = ["Rectificación", "Cambio"]
CAMPOS = ["Por defecto", "Todos los campos", "Solo la lista"]
LISTAS = {"Sí/No": ["Sí", "No"], "Acción": ACCIONES, "Campos": CAMPOS}

# (título, tipo, lista, ancho, sección)
COLUMNAS = [
    ("N°", "meta", None, 5, ""), ("Ejecutar", "meta", "Sí/No", 9, ""),
    ("Nombre de la tarifa", "meta", None, 22, ""),
    ("Acción", "obligatorio", "Acción", 14, "Qué hacer"),
    ("Delta (vacío: +1 rectificación, −1 cambio)", "opcional", None, 13, "Qué hacer"),
    ("Campos a modificar", "obligatorio", "Campos", 16, "Qué hacer"),
    ("Lista de campos (separados por ;)", "opcional", None, 30, "Qué hacer"),
    ("Agregar una fila a cada tabla", "opcional", "Sí/No", 11, "Tablas y condicionales (paso 2)"),
    ("Pasar «ilimitado» a limitado", "opcional", "Sí/No", 11, "Tablas y condicionales (paso 2)"),
    ("Cambiar canales (TV de paga)", "opcional", "Sí/No", 11, "Tablas y condicionales (paso 2)"),
    ("Cambio: fecha de ejecución", "opcional", None, 13, "Solo cambio"),
    ("Texto del sello (vacío: «Rectificación» / «Cambio»)", "opcional", None, 18, "Textos"),
    ("Estado", "script", None, 13, "Resultado"), ("Código SIRT", "script", None, 18, "Resultado"),
    ("Fecha de ejecución", "script", None, 17, "Resultado"), ("Detalle", "script", None, 60, "Resultado"),
]

MUESTRAS = [  # (tarifa, acción, delta, campos, lista, agregar, ilimitado, canales, fecha)
    ("TE-IF-CM-004", "Rectificación", "", "Por defecto", "", "No", "No", "", ""),
    ("TE-SMP-CM-003", "Cambio", "", "Todos los campos", "", "Sí", "Sí", "", ""),
    ("TE-TVP-CM-005", "Rectificación", "", "Todos los campos", "", "No", "No", "Sí", ""),
    ("TE-EMP-CM-002", "Cambio", "1", "Por defecto", "", "No", "No", "", ""),
    ("TE-TF-CM-005", "Rectificación", "", "Solo la lista", "Otros datos generales; Valor del servicio", "No", "No", "", ""),
    ("TL-LYC-CM-001", "Rectificación", "", "Por defecto", "", "No", "No", "", ""),
]

DESCRIPCION = [
    ("Qué hace cada opción", True, 14),
    ("", False, 10),
    ("ACCIÓN", True, 12),
    ("• Rectificación: corrige la tarifa vigente; se guarda de inmediato.", False, 10),
    ("• Cambio: programa una nueva versión de la tarifa para una fecha de ejecución. Solo puede haber UN cambio", False, 10),
    ("  pendiente por tarifa: si ya hay uno programado, la fila queda «No aplica» hasta que se ejecute.", False, 10),
    ("", False, 10),
    ("DELTA (cuánto suben o bajan los números)", True, 12),
    ("Vacío: +1 en rectificación y −1 en cambio. Se puede poner otro número (p. ej. 2, −1, 0.5).", False, 10),
    ("Nunca baja de 0, los porcentajes no pasan de 100 y los «topes» se cambian en el orden que exige SIRT.", False, 10),
    ("Empaquetados que suben: el último servicio no sube (regla: suma de servicios < suma de las tarifas base).", False, 10),
    ("", False, 10),
    ("CAMPOS A MODIFICAR", True, 12),
    ("• Por defecto: textos (con el sello «| Rectificación/Cambio fecha hora»), valor del servicio, tarifa de", False, 10),
    ("  instalación y listas como «Tipo de Contratación».", False, 10),
    ("• Todos los campos: además TODOS los números del paso 2 (datos, velocidades, minutos, tarifas…), «Apps sin", False, 10),
    ("  consumo de datos» y los números de cada fila de las tablas (lápiz → Guardar), sin tocar tecnología, países", False, 10),
    ("  ni plataformas (SIRT no permite repetirlos).", False, 10),
    ("• Solo la lista: solo los campos escritos en «Lista de campos», con su nombre como aparece en SIRT.", False, 10),
    ("", False, 10),
    ("TABLAS Y CONDICIONALES (paso 2)", True, 12),
    ("• Agregar una fila a cada tabla: una fila nueva con un país / tecnología / plataforma que la tabla aún no tenga.", False, 10),
    ("• Pasar «ilimitado» a limitado: en minutos y SMS marcados como ilimitados, desmarca la casilla y llena", False, 10),
    ("  cantidad y tarifa adicional (10 y 15).", False, 10),
    ("• Cambiar canales: sube el otro Excel de canales de automatizacion/canales/ (3 ↔ 5 canales). Vacío = Sí.", False, 10),
    ("", False, 10),
    ("CAMBIO: FECHA DE EJECUCIÓN", True, 12),
    ("Vacío: hoy + 3 días (SIRT pide al menos 2 si el cambio perjudica al usuario; hoy siempre está bloqueado).", False, 10),
    ("Si el día está bloqueado en el calendario, usa el siguiente permitido (hasta 7 días).", False, 10),
    ("", False, 10),
    ("POR TIPO DE TARIFA", True, 12),
    ("• Monoproductos y empaquetados: todo lo anterior (en empaquetados, en cada pestaña de servicio).", False, 10),
    ("• Licitación y Contratos (solo rectificación): además mueve inicio y fin de vigencia (+1 día, con el", False, 10),
    ("  calendario), cambia Ámbito y Moneda a la otra opción y suma el delta al Valor de contrato.", False, 10),
    ("• Promociones: la rectificación / cambio de promociones aún no está grabada; esas filas quedarán en error.", False, 10),
    ("", False, 10),
    ("REGLAS GENERALES", True, 12),
    ("• La tarifa debe estar EN COMERCIALIZACIÓN y VIGENTE (licitaciones: vigente). Si no, «No aplica».", False, 10),
    ("• Antes de guardar se comprueba que todos los cambios aparezcan en el resumen de SIRT.", False, 10),
    ("• Solo se procesan las filas con «Ejecutar» = Sí. El Excel debe estar cerrado al correr el comando.", False, 10),
]


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

    wb = Workbook()
    hoja = wb.active
    hoja.title = "Modificaciones"
    guia = wb.create_sheet("Qué hace cada opción")
    listas = wb.create_sheet("Listas")
    listas.sheet_state = "hidden"
    rangos = {}
    for k, (nombre, valores) in enumerate(LISTAS.items(), start=1):
        letra = get_column_letter(k)
        for f, v in enumerate(valores, start=1):
            listas[f"{letra}{f}"] = v
        rangos[nombre] = f"=Listas!${letra}$1:${letra}${len(valores)}"

    hoja["A1"] = "Rectificación y cambio de tarifa"
    hoja["A1"].font = Font(name=F, bold=True, size=14)
    hoja["A2"] = ("Una fila por acción. Se indica qué hacer, no el valor nuevo: el script toma lo que la tarifa tiene "
                  "en SIRT y lo modifica. Detalle en la hoja «Qué hace cada opción».")
    hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    for j, (texto, clave) in enumerate([("Obligatorio", "obligatorio"), ("Opcional", "opcional"),
                                        ("Lo llena el script", "script")], start=4):
        cabeza(hoja.cell(3, j, texto), clave)
    secciones = [c[4] for c in COLUMNAS]
    for j, s in enumerate(secciones, start=1):
        c = hoja.cell(4, j, s if (s and (j == 1 or s != secciones[j - 2])) else None)
        c.fill, c.border = SECCION, BORDE
        c.font = Font(name=F, bold=True, size=8, color="1F4E78")
    for j, (titulo, tipo, nombre_lista, ancho, _) in enumerate(COLUMNAS, start=1):
        cabeza(hoja.cell(5, j, titulo), tipo)
        hoja.column_dimensions[get_column_letter(j)].width = ancho
        if nombre_lista:
            dv = DataValidation(type="list", formula1=rangos[nombre_lista], allow_blank=True)
            dv.showErrorMessage, dv.error = True, "Elige una opción de la lista."
            hoja.add_data_validation(dv)
            letra = get_column_letter(j)
            dv.add(f"{letra}6:{letra}{5 + FILAS_CON_LISTAS}")
    hoja.row_dimensions[5].height = 58

    for i, (tarifa, accion, delta, campos, lista_c, agregar, ilimitado, canales, fecha) in enumerate(MUESTRAS, start=6):
        fila = [i - 5, "Sí", tarifa, accion, delta, campos, lista_c, agregar, ilimitado, canales, fecha, "",
                "", "", "", ""]
        for j, v in enumerate(fila, start=1):
            c = hoja.cell(i, j, v)
            c.font, c.border = Font(name=F, size=9), BORDE
            c.alignment = Alignment(vertical="top", wrap_text=True,
                                    horizontal="center" if j in (1, 2, 5, 8, 9, 10) else None)
            if COLUMNAS[j - 1][1] == "script":
                c.fill = CELDA_SCRIPT
    for i in range(6, 6 + FILAS_CON_LISTAS):
        hoja.cell(i, 11).number_format = "DD/MM/YYYY"
    delta = DataValidation(type="decimal", operator="between", formula1="-1000", formula2="1000", allow_blank=True)
    delta.showErrorMessage, delta.error = True, "El delta es un número (p. ej. 1, -1, 0.5)."
    hoja.add_data_validation(delta)
    delta.add(f"E6:E{5 + FILAS_CON_LISTAS}")
    fecha = DataValidation(type="date", operator="greaterThan", formula1="TODAY()", allow_blank=True)
    fecha.showErrorMessage, fecha.error = True, "La fecha de ejecución debe ser posterior a hoy (vacío = hoy + 3)."
    hoja.add_data_validation(fecha)
    fecha.add(f"K6:K{5 + FILAS_CON_LISTAS}")
    hoja.freeze_panes = "D6"
    hoja.auto_filter.ref = f"A5:{get_column_letter(len(COLUMNAS))}{5 + len(MUESTRAS)}"

    for i, (texto, negrita, tam) in enumerate(DESCRIPCION, start=1):
        guia.cell(i, 1, texto).font = Font(name=F, bold=negrita, size=tam,
                                          color="1F4E78" if negrita and tam == 12 else "000000")
    guia.column_dimensions["A"].width = 115

    SALIDA.parent.mkdir(exist_ok=True)
    wb.save(SALIDA)
    print(f"Generado: {SALIDA} ({len(MUESTRAS)} filas de muestra)")


if __name__ == "__main__":
    generar()
