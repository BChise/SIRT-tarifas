"""
Genera automatizacion/acciones_basicas.xlsx: el Excel para aplicar CESE, RENOMBRE y
FE DE ERRATAS a tarifas ya registradas. Lo usa el comando:

    python acciones_excel.py            (ver acciones_excel.py)

Uso (desde la carpeta del proyecto):
    python automatizacion/generar_plantilla_acciones.py
"""

from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
SALIDA = PROYECTO / "automatizacion" / "acciones_basicas.xlsx"
FILAS_CON_LISTAS = 1000
ACCIONES = ["Fe de erratas", "Renovación", "Cese", "Renombre"]

COLUMNAS = [  # (título, tipo: meta | obligatorio | opcional | script, ancho)
    ("N°", "meta", 5), ("Ejecutar", "meta", 9), ("Nombre de la tarifa", "meta", 26),
    ("Acción", "obligatorio", 15),
    ("Renombre: sufijo", "opcional", 16),
    ("Cese: fecha de fin de comercialización", "opcional", 18),
    ("Renovación: días de comercialización", "opcional", 14),
    ("Renovación: duración (días)", "opcional", 12),
    ("Estado", "script", 13), ("Código SIRT", "script", 18), ("Fecha de ejecución", "script", 17),
    ("Detalle", "script", 60),
]
MUESTRAS = [
    ("TE-TF-CM-004", "Fe de erratas", "", ""),
    ("TP-TF-PROMO-017", "Renovación", "", ""),
    ("TE-TFP-CM-004", "Renombre", "_REN", ""),
    ("TE-IF-CM-005", "Renombre", "_V2", ""),
    ("TE-EMP-CM-001", "Fe de erratas", "", ""),
    ("TP-IF-PROMO-021", "Fe de erratas", "", ""),
    ("TE-TVP-CM-004", "Cese", "", ""),
    ("TE-SMP-CM-005", "Fe de erratas", "", ""),
]

DESCRIPCION = [
    ("Qué hace cada acción", True, 14),
    ("", False, 10),
    ("FE DE ERRATAS", True, 12),
    ("Corrige TODOS los textos que SIRT permite cambiar en una fe de erratas, en los 2 pasos y en cada", False, 10),
    ("pestaña de servicio (condiciones, restricciones, otros datos, detalle del servicio, nombre de empresa…).", False, 10),
    ("A cada texto le agrega al final el sello «| Fe de erratas dd/mm/aaaa hh:mm» (si ya tenía uno de fe de erratas,", False, 10),
    ("lo reemplaza; no se acumulan). En licitaciones, el «Documento de Empresa» sube 1 en su última cifra.", False, 10),
    ("No cambia números, listas ni tablas (la fe de erratas no lo permite). Antes de guardar comprueba que todo", False, 10),
    ("aparezca en el resumen de SIRT; si algo no cuadra, no guarda.", False, 10),
    ("", False, 10),
    ("RENOMBRE", True, 12),
    ("Agrega un sufijo al final del nombre de la tarifa. En la columna «Renombre: sufijo»:", False, 10),
    ("  • vacío → «_REN» (TE-TF-CM-001 → TE-TF-CM-001_REN);", False, 10),
    ("  • un texto → se agrega tal cual al final (con «_V2» → TE-TF-CM-001_V2).", False, 10),
    ("Después de guardar busca la tarifa con el nombre nuevo para confirmar el cambio.", False, 10),
    ("", False, 10),
    ("RENOVACIÓN (solo tarifas promocionales)", True, 12),
    ("Renueva la promoción por un nuevo periodo, en los 3 pasos de SIRT:", False, 10),
    ("  • Comercialización: el calendario abre en el día siguiente al fin actual; el inicio es ese día + 1 y el fin,", False, 10),
    ("    inicio + «Renovación: días de comercialización» (vacío = 19).", False, 10),
    ("    Ejemplo: fin actual 14/12/2026 → inicio 16/12/2026 → fin 04/01/2027.", False, 10),
    ("  • Duraciones: la del valor del plan (y de la tarifa de instalación, si la tiene) y la de CADA atributo", False, 10),
    ("    promocionado en cada pestaña de servicio pasan a «Renovación: duración» (vacío = 25, o 31 si ya era 25).", False, 10),
    ("  • Antes de cambiar nada, lee de SIRT qué atributos se promocionaron y comprueba que en pantalla estén", False, 10),
    ("    todos; si no cuadra, no guarda. Verifica las fechas en el resumen y guarda («Guardar cambios»).", False, 10),
    ("", False, 10),
    ("CESE DE COMERCIALIZACIÓN", True, 12),
    ("Programa el fin de comercialización de la tarifa. Fecha vacía → mañana (hoy siempre está bloqueado).", False, 10),
    ("Solo monoproductos y empaquetados: una PROMOCIÓN no se puede cesar (la fila queda en error y no se toca SIRT).", False, 10),
    ("Si el día elegido está bloqueado en el calendario de SIRT, usa el siguiente día permitido (hasta 7 días)", False, 10),
    ("y lo anota en «Detalle». La fecha debe ser posterior a hoy.", False, 10),
    ("", False, 10),
    ("REGLAS GENERALES", True, 12),
    ("• Sirve para monoproductos, empaquetados y promociones (el cese, solo para los dos primeros; la renovación,", False, 10),
    ("  solo para promociones).", False, 10),
    ("• Cada acción usa solo sus columnas (sufijo → renombre; fecha → cese; días y duración → renovación). Lo que", False, 10),
    ("  no corresponda a la acción de la fila se ignora y se avisa en «Detalle».", False, 10),
    ("• Las tres acciones exigen que la tarifa esté EN COMERCIALIZACIÓN y VIGENTE (el día siguiente a su registro", False, 10),
    ("  definitivo). Si no lo está, no se toca SIRT y el Estado queda «No aplica» con el motivo.", False, 10),
    ("  (Licitación y Contratos: basta que esté vigente.)", False, 10),
    ("• Una misma tarifa puede tener varias filas (p. ej. fe de erratas y renombre). El script hace primero la", False, 10),
    ("  fe de erratas, luego la renovación, el cese y al final el renombre (para no perder la tarifa al cambiarle", False, 10),
    ("  el nombre).", False, 10),
    ("• Si una acción ya se aplicó con éxito a esa tarifa, no se repite (Estado «Ya aplicada»).", False, 10),
    ("• Solo se procesan las filas con «Ejecutar» = Sí. El Excel debe estar cerrado al correr el comando.", False, 10),
    ("• Al terminar, cada fila muestra su resultado en las columnas grises y se guarda una copia del Excel en", False, 10),
    ("  resultados/ACCIONES/.", False, 10),
    ("", False, 10),
    ("Comando (desde la carpeta del proyecto, con Edge abierto y la sesión de SIRT iniciada)", True, 12),
    ("  python acciones_excel.py --solo-listar     → muestra lo que haría, sin tocar SIRT", False, 10),
    ("  python acciones_excel.py                   → aplica las acciones (pide confirmación)", False, 10),
    ("  python acciones_excel.py --excel otra_copia.xlsx --si", False, 10),
]


def generar() -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    F = "Arial"
    lado = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=lado, right=lado, top=lado, bottom=lado)
    RELLENO = {"meta": (PatternFill("solid", fgColor="375623"), "FFFFFF"),
               "obligatorio": (PatternFill("solid", fgColor="1F4E78"), "FFFFFF"),
               "opcional": (PatternFill("solid", fgColor="BDD7EE"), "1F1F1F"),
               "script": (PatternFill("solid", fgColor="808080"), "FFFFFF")}
    CELDA_SCRIPT = PatternFill("solid", fgColor="EDEDED")

    wb = Workbook()
    hoja = wb.active
    hoja.title = "Acciones"
    guia = wb.create_sheet("Qué hace cada acción")
    listas = wb.create_sheet("Listas")
    listas.sheet_state = "hidden"

    hoja["A1"] = "Fe de erratas, renovación, cese y renombre"
    hoja["A1"].font = Font(name=F, bold=True, size=14)
    hoja["A2"] = ("Una fila por acción. Elige la acción en la lista; lo que hace cada una está en la hoja "
                  "«Qué hace cada acción».")
    hoja["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    for j, (texto, clave) in enumerate([("Obligatorio", "obligatorio"), ("Opcional", "opcional"),
                                        ("Lo llena el script", "script")], start=4):
        relleno, color = RELLENO[clave]
        c = hoja.cell(3, j, texto)
        c.fill, c.border = relleno, BORDE
        c.font = Font(name=F, bold=True, size=9, color=color)
        c.alignment = Alignment(horizontal="center")

    for j, (titulo, clave, ancho) in enumerate(COLUMNAS, start=1):
        relleno, color = RELLENO[clave]
        c = hoja.cell(5, j, titulo)
        c.fill, c.border = relleno, BORDE
        c.font = Font(name=F, bold=True, size=9, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        hoja.column_dimensions[get_column_letter(j)].width = ancho
    hoja.row_dimensions[5].height = 40

    for i, (nombre, accion, renombre, fecha) in enumerate(MUESTRAS, start=6):
        fila = [i - 5, "Sí", nombre, accion, renombre, fecha, "", "", "", "", "", ""]
        for j, v in enumerate(fila, start=1):
            c = hoja.cell(i, j, v)
            c.font = Font(name=F, size=9)
            c.border = BORDE
            c.alignment = Alignment(vertical="top", wrap_text=j >= 12,
                                    horizontal="center" if j in (1, 2) else None)
            if COLUMNAS[j - 1][1] == "script":
                c.fill = CELDA_SCRIPT
    # La columna de fecha acepta fechas de Excel o texto dd/mm/aaaa.
    for i in range(6, 6 + FILAS_CON_LISTAS):
        hoja.cell(i, 6).number_format = "DD/MM/YYYY"

    for k, v in enumerate(["Sí", "No"], start=1):
        listas[f"A{k}"] = v
    for k, v in enumerate(ACCIONES, start=1):
        listas[f"B{k}"] = v
    for col, rango in ((2, "$A$1:$A$2"), (4, f"$B$1:$B${len(ACCIONES)}")):
        dv = DataValidation(type="list", formula1=f"=Listas!{rango}", allow_blank=True)
        dv.showErrorMessage = True
        dv.errorTitle, dv.error = "Valor no válido", "Elige una opción de la lista."
        hoja.add_data_validation(dv)
        letra = get_column_letter(col)
        dv.add(f"{letra}6:{letra}{5 + FILAS_CON_LISTAS}")
    fecha = DataValidation(type="date", operator="greaterThan", formula1="TODAY()", allow_blank=True)
    fecha.showErrorMessage = True
    fecha.errorTitle, fecha.error = "Fecha no válida", "La fecha de cese debe ser posterior a hoy (vacío = mañana)."
    hoja.add_data_validation(fecha)
    fecha.add(f"F6:F{5 + FILAS_CON_LISTAS}")

    hoja.freeze_panes = "D6"
    hoja.auto_filter.ref = f"A5:L{5 + len(MUESTRAS)}"
    for col, maximo, texto in ((7, 3650, "Días de comercialización: número entero (vacío = 19)."),
                               (8, 180, "Duración: de 1 a 180 días (vacío = 25, o 31 si ya era 25).")):
        dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(maximo), allow_blank=True)
        dv.showErrorMessage, dv.error = True, texto
        hoja.add_data_validation(dv)
        dv.add(f"{get_column_letter(col)}6:{get_column_letter(col)}{5 + FILAS_CON_LISTAS}")

    for i, (texto, negrita, tam) in enumerate(DESCRIPCION, start=1):
        guia.cell(i, 1, texto).font = Font(name=F, bold=negrita, size=tam,
                                          color="1F4E78" if negrita and tam == 12 else "000000")
    guia.column_dimensions["A"].width = 115

    SALIDA.parent.mkdir(exist_ok=True)
    wb.save(SALIDA)
    print(f"Generado: {SALIDA} ({len(MUESTRAS)} filas de muestra)")


if __name__ == "__main__":
    generar()
