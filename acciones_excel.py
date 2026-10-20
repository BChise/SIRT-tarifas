"""
Aplica CESE, RENOMBRE y FE DE ERRATAS según automatizacion/acciones_basicas.xlsx.

Cada fila del Excel (hoja «Acciones») = una acción sobre una tarifa. El script:
1. Lee y valida las filas con «Ejecutar» = Sí (acción de la lista, fecha posterior a hoy…).
2. Las ordena: primero fe de erratas, luego cese y al final renombre (para no perder la
   tarifa al cambiarle el nombre si tiene varias acciones).
3. Aplica cada una con acciones.py (el mismo código ya probado) y escribe el resultado en
   las columnas grises de su fila: Estado, Código SIRT, Fecha de ejecución y Detalle.
4. Guarda una copia del Excel con los resultados en resultados/ACCIONES/.

Uso (desde la carpeta del proyecto, con Edge abierto y la sesión de SIRT iniciada):
    python acciones_excel.py --solo-listar        muestra lo que haría, sin tocar SIRT
    python acciones_excel.py                      aplica las acciones (pide confirmación)
    python acciones_excel.py --si --diagnostico
    python acciones_excel.py --excel otra_copia.xlsx

El Excel debe estar CERRADO (en Windows, Excel bloquea el archivo abierto).
"""

import argparse
import csv
import re
import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent
EXCEL = PROYECTO / "automatizacion" / "acciones_basicas.xlsx"
HOJA = "Acciones"
FILA_CABECERA = 5
CSV_ACCIONES = PROYECTO / "resultados" / "ACCIONES" / "acciones.csv"

ACCIONES = {"fe de erratas": "fe_erratas", "renovación": "renovar", "renovacion": "renovar",
            "cese": "cese", "renombre": "renombre"}
ORDEN = {"fe_erratas": 0, "renovar": 1, "cese": 2, "renombre": 3}
ESTADOS = {"EXITOSO": "Aplicada", "ERROR": "Error", "NO_APLICA": "No aplica"}
COLUMNAS = {"ejecutar": "Ejecutar", "nombre": "Nombre de la tarifa", "accion": "Acción",
            "renombre": "Renombre: sufijo", "fecha": "Cese: fecha de fin de comercialización",
            "dias_ren": "Renovación: días de comercialización", "duracion_ren": "Renovación: duración (días)",
            "estado": "Estado", "codigo": "Código SIRT", "fecha_ejec": "Fecha de ejecución", "detalle": "Detalle"}


def leer_argumentos():
    p = argparse.ArgumentParser(description="Cese, renombre y fe de erratas desde un Excel.")
    p.add_argument("--excel", default=str(EXCEL))
    p.add_argument("--solo-listar", action="store_true", help="muestra lo que haría, sin tocar SIRT")
    p.add_argument("--si", action="store_true", help="no pedir confirmación")
    p.add_argument("--diagnostico", action="store_true", help="capturas por paso (como en acciones.py)")
    return p.parse_args()


def _texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


def leer_fecha(valor) -> date | None:
    if valor in (None, ""):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = _texto(valor)
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    raise ValueError(f"fecha «{texto}» no reconocida (usa dd/mm/aaaa)")


def excel_esta_abierto(ruta: Path) -> bool:
    """Windows: Excel bloquea el archivo. Linux/LibreOffice: deja un archivo .~lock."""
    if (ruta.parent / f".~lock.{ruta.name}#").exists() or (ruta.parent / f"~${ruta.name}").exists():
        return True
    try:
        with ruta.open("r+b"):
            return False
    except PermissionError:
        return True


def leer_filas(hoja) -> tuple[dict, list[dict]]:
    cabecera = {_texto(c.value): c.column for c in hoja[FILA_CABECERA] if c.value}
    faltan = [t for t in COLUMNAS.values() if t not in cabecera]
    if faltan:
        sys.exit(f"Faltan columnas en la hoja «{HOJA}»: {', '.join(faltan)}")
    col = {k: cabecera[t] for k, t in COLUMNAS.items()}
    filas = []
    for r in range(FILA_CABECERA + 1, hoja.max_row + 1):
        nombre = _texto(hoja.cell(r, col["nombre"]).value)
        if not nombre:
            continue
        filas.append({"fila": r, "nombre": nombre,
                      "ejecutar": _texto(hoja.cell(r, col["ejecutar"]).value).lower() or "sí",   # vacío = Sí
                      "accion_texto": _texto(hoja.cell(r, col["accion"]).value),
                      "renombre": _texto(hoja.cell(r, col["renombre"]).value),
                      "fecha": hoja.cell(r, col["fecha"]).value,
                      "dias_ren": hoja.cell(r, col["dias_ren"]).value,
                      "duracion_ren": hoja.cell(r, col["duracion_ren"]).value})
    return col, filas


def validar(fila: dict) -> str:
    """Completa la fila (acción y parámetros) y devuelve un mensaje de error o ''."""
    if fila["ejecutar"] not in ("sí", "si", "no"):
        return f"«Ejecutar» debe ser Sí o No (tiene «{fila['ejecutar']}»)"
    accion = ACCIONES.get(fila["accion_texto"].lower())
    if not accion:
        return f"acción «{fila['accion_texto']}» no válida (Fe de erratas, Renovación, Cese o Renombre)"
    fila["accion"] = accion
    fila["extra"] = []
    # Cada acción usa solo su columna; lo demás se ignora (y se avisa en «Detalle»).
    fila["avisos"] = []
    if accion != "renombre" and fila["renombre"]:
        fila["avisos"].append(f"se ignoró el sufijo «{fila['renombre']}» (solo aplica al renombre)")
    if accion != "cese" and fila["fecha"] not in (None, ""):
        fila["avisos"].append("se ignoró la fecha de cese (solo aplica al cese)")
    if accion != "renovar" and any(fila[k] not in (None, "") for k in ("dias_ren", "duracion_ren")):
        fila["avisos"].append("se ignoraron los datos de renovación (solo aplican a la renovación)")
    if accion == "renovar":
        if not es_promocion(fila["nombre"]):
            return "la renovación solo aplica a tarifas promocionales (TP-…)"
        for clave, opcion, minimo, maximo in (("dias_ren", "--dias-renovacion", 1, 3650),
                                              ("duracion_ren", "--duracion-renovacion", 1, 180)):
            valor = fila[clave]
            if valor in (None, ""):
                continue
            try:
                numero = int(float(str(valor).strip()))
            except ValueError:
                return f"«{COLUMNAS[clave]}» debe ser un número de días (tiene «{valor}»)"
            if not minimo <= numero <= maximo:
                return f"«{COLUMNAS[clave]}» debe estar entre {minimo} y {maximo} días"
            fila["extra"] += [opcion, str(numero)]
    if accion == "renombre":
        sufijo = fila["renombre"] or "_REN"
        fila["extra"] = ["--sufijo", sufijo]
        fila["nombre_nuevo"] = fila["nombre"] + sufijo
    if accion == "cese" and es_promocion(fila["nombre"]):
        return "una promoción no se puede cesar (solo monoproductos y empaquetados)"
    if accion == "cese":
        try:
            fecha = leer_fecha(fila["fecha"]) or date.today() + timedelta(days=1)
        except ValueError as error:
            return str(error)
        if fecha <= date.today():
            return f"la fecha de cese {fecha:%d/%m/%Y} debe ser posterior a hoy"
        fila["extra"] = ["--fecha", fecha.isoformat()]
        fila["fecha_cese"] = fecha
    return ""


def es_promocion(nombre: str) -> bool:
    """Por el nombre (TP-…) o por los resultados de promocionar.py (resultados/PRM)."""
    if nombre.upper().startswith("TP-"):
        return True
    for ruta in (PROYECTO / "resultados" / "PRM").glob("resultados_*.csv"):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            if any(f.get("nombre") == nombre for f in csv.DictReader(entrada, delimiter=";")):
                return True
    return False


def resultado_en_csv(accion: str, nombre: str, desde: str) -> dict | None:
    if not CSV_ACCIONES.exists():
        return None
    with CSV_ACCIONES.open("r", newline="", encoding="utf-8-sig") as entrada:
        filas = [f for f in csv.DictReader(entrada, delimiter=";")
                 if f.get("accion") == accion and f.get("nombre") == nombre and (f.get("fecha_hora") or "") >= desde]
    return filas[-1] if filas else None


def describir(f: dict) -> str:
    if f["accion"] == "renombre":
        return f"renombre → {f['nombre_nuevo']}"
    if f["accion"] == "cese":
        return f"cese → fin de comercialización {f['fecha_cese']:%d/%m/%Y}"
    if f["accion"] == "renovar":
        dias = next((f["extra"][k + 1] for k, x in enumerate(f["extra"]) if x == "--dias-renovacion"), "19")
        dur = next((f["extra"][k + 1] for k, x in enumerate(f["extra"]) if x == "--duracion-renovacion"),
                   "25 (o 31 si ya era 25)")
        return f"renovación → comercialización {dias} días; duraciones {dur}"
    return "fe de erratas → todos los textos con «| Fe de erratas <fecha hora>»"


def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    if not ruta.exists():
        sys.exit(f"No existe {ruta}")
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill

    if not args.solo_listar and excel_esta_abierto(ruta):
        sys.exit(f"El Excel está abierto: ciérralo y vuelve a correr el comando.\n  {ruta}")
    wb = load_workbook(ruta)
    if HOJA not in wb.sheetnames:
        sys.exit(f"El Excel no tiene la hoja «{HOJA}».")
    hoja = wb[HOJA]
    col, filas = leer_filas(hoja)

    def escribir(fila: dict, estado: str, codigo: str = "", fecha: str = "", detalle: str = "") -> None:
        colores = {"Aplicada": "C6EFCE", "Error": "FFC7CE", "Error de datos": "FFC7CE",
                   "No aplica": "FFEB9C", "Ya aplicada": "DDEBF7", "Omitida": "EDEDED"}
        for clave, valor in (("estado", estado), ("codigo", codigo), ("fecha_ejec", fecha), ("detalle", detalle)):
            celda = hoja.cell(fila["fila"], col[clave], valor)
            celda.font = Font(name="Arial", size=9, bold=clave == "estado")
        hoja.cell(fila["fila"], col["estado"]).fill = PatternFill("solid", fgColor=colores.get(estado, "EDEDED"))

    a_correr, con_error = [], []
    for f in filas:
        problema = validar(f)
        if problema:
            con_error.append((f, problema))
        elif f["ejecutar"] == "no":
            escribir(f, "Omitida")
        else:
            a_correr.append(f)
    a_correr.sort(key=lambda f: (ORDEN[f["accion"]], f["fila"]))

    print(f"Excel: {ruta}")
    print(f"Filas: {len(filas)} · a ejecutar: {len(a_correr)} · con errores de datos: {len(con_error)}")
    for f, problema in con_error:
        print(f"  ✖ fila {f['fila']} ({f['nombre']}): {problema}")
        escribir(f, "Error de datos", detalle=problema)
    for k, f in enumerate(a_correr, start=1):
        print(f"  {k:>2}. fila {f['fila']}: {f['nombre']}  →  {describir(f)}"
              + (f"   (aviso: {'; '.join(f['avisos'])})" if f.get("avisos") else ""))

    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT ni el Excel.")
        return
    if not a_correr:
        wb.save(ruta)
        print("No hay acciones para ejecutar.")
        return
    if not args.si:
        print(f"\nATENCIÓN: se aplicarán {len(a_correr)} acción(es) en SIRT (probablemente irreversibles).")
        if input("Escribe 'si' para continuar: ").strip().lower() not in ("si", "sí"):
            print("Cancelado.")
            return

    resumen = {}
    for k, f in enumerate(a_correr, start=1):
        print("\n" + "=" * 70)
        print(f"[{k}/{len(a_correr)}] fila {f['fila']}: {f['nombre']} → {describir(f)}")
        desde = (datetime.now() - timedelta(seconds=2)).strftime("%Y-%m-%d %H:%M:%S")
        comando = [sys.executable, "-u", str(PROYECTO / "acciones.py"), f["accion"],
                   "--tarifas", f["nombre"], "--si", *f["extra"]] + (["--diagnostico"] if args.diagnostico else [])
        proceso = subprocess.run(comando, cwd=PROYECTO, capture_output=True, text=True,
                                 encoding="utf-8", errors="replace")
        salida = (proceso.stdout or "") + (proceso.stderr or "")
        for linea in salida.splitlines():
            if linea.strip() and not linea.startswith("=" * 20):
                print("   " + linea)
        registro = resultado_en_csv(f["accion"], f["nombre"], desde)
        if registro:
            estado = ESTADOS.get(registro.get("resultado", ""), registro.get("resultado", "Error"))
            detalle = registro.get("detalle", "")
            if f["accion"] == "renombre" and estado == "Aplicada":
                detalle = f"Nuevo nombre: {f['nombre_nuevo']} | {detalle}"
            if f.get("avisos"):
                detalle = " | ".join(f["avisos"]) + " | " + detalle
            escribir(f, estado, registro.get("codigo_sirt", ""), registro.get("fecha_hora", ""), detalle[:1000])
        elif "se omite" in salida:
            escribir(f, "Ya aplicada", detalle="Esta acción ya figura como EXITOSA para la tarifa (no se repitió).")
            estado = "Ya aplicada"
        else:
            ultimas = " | ".join(l.strip() for l in salida.splitlines() if l.strip())[-500:]
            escribir(f, "Error", detalle=f"acciones.py terminó sin registrar resultado: {ultimas}")
            estado = "Error"
        resumen[estado] = resumen.get(estado, 0) + 1
        try:
            wb.save(ruta)                                     # se guarda tras cada fila
        except PermissionError:
            print("   Aviso: no se pudo guardar el Excel (¿se abrió?); se guardará al final en una copia.")

    marca = datetime.now().strftime("%Y%m%d_%H%M")
    copia = PROYECTO / "resultados" / "ACCIONES" / f"acciones_excel_{marca}.xlsx"
    copia.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(ruta)
        shutil.copyfile(ruta, copia)
    except PermissionError:
        wb.save(copia)
        print(f"No se pudo escribir en {ruta}; los resultados quedaron solo en la copia.")
    print("\n" + "=" * 70)
    print("Resumen: " + " · ".join(f"{k}: {v}" for k, v in resumen.items()))
    print(f"Excel actualizado: {ruta}")
    print(f"Copia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
