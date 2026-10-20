"""
Aplica RECTIFICACIONES y CAMBIOS DE TARIFA según automatizacion/modificaciones.xlsx.

Cada fila = una modificación sobre una tarifa. Para cada fila con «Ejecutar» = Sí (o vacío) valida
los datos y la aplica con acciones.py (rectificar / cambio: el mismo código ya probado), con las
opciones de la fila: delta, campos, agregar filas, ilimitado → limitado, canales, fecha de ejecución
y sello. Escribe Estado, Código SIRT, Fecha y Detalle y guarda una copia en resultados/MODIFICACIONES_EXCEL/.

Uso:
    python modificaciones_excel.py --solo-listar
    python modificaciones_excel.py --hasta-resumen          llega al resumen de SIRT y NO guarda (para probar)
    python modificaciones_excel.py --nombres TE-IF-CM-004
    python modificaciones_excel.py --si --diagnostico
"""

import argparse
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "modificaciones.xlsx"
HOJA = "Modificaciones"
CSV_ACCIONES = xc.RESULTADOS / "ACCIONES" / "acciones.csv"
ACCIONES = {"rectificación": "rectificar", "rectificacion": "rectificar", "cambio": "cambio"}
ESTADOS = {"EXITOSO": "Aplicada", "ERROR": "Error", "NO_APLICA": "No aplica", "EN_RESUMEN": "Hasta resumen"}
C = {"accion": "Acción", "delta": "Delta (vacío: +1 rectificación, −1 cambio)", "campos": "Campos a modificar",
     "lista": "Lista de campos (separados por ;)", "filas": "Agregar una fila a cada tabla",
     "ilimitado": "Pasar «ilimitado» a limitado", "canales": "Cambiar canales (TV de paga)",
     "fecha": "Cambio: fecha de ejecución", "sello": "Texto del sello (vacío: «Rectificación» / «Cambio»)"}


def leer_argumentos():
    p = argparse.ArgumentParser(description="Rectificaciones y cambios desde el Excel.")
    p.add_argument("--excel", default=str(EXCEL))
    p.add_argument("--nombres", nargs="+")
    p.add_argument("--limite", type=int, default=0)
    p.add_argument("--repetir", action="store_true", help="aplica aunque la tarifa ya tenga esa acción EXITOSA")
    p.add_argument("--hasta-resumen", action="store_true", help="llega al resumen y no guarda")
    p.add_argument("--solo-listar", action="store_true")
    p.add_argument("--si", action="store_true")
    p.add_argument("--diagnostico", action="store_true")
    return p.parse_args()


def construir(f: dict) -> dict:
    nombre = texto(f["Nombre de la tarifa"])
    accion = ACCIONES.get(texto(f.get(C["accion"])).lower())
    if not accion:
        raise DatoInvalido(f"acción «{texto(f.get(C['accion']))}» no válida (Rectificación o Cambio)")
    if nombre.upper().startswith("TP-"):
        raise DatoInvalido("la rectificación / cambio de promociones aún no está grabada")
    if nombre.upper().startswith("TL-") and accion == "cambio":
        raise DatoInvalido("las licitaciones solo admiten rectificación")
    extra, avisos = [], []
    delta = texto(f.get(C["delta"])).replace(",", ".").replace("−", "-")
    if delta:
        try:
            Decimal(delta)
        except InvalidOperation:
            raise DatoInvalido(f"«Delta» debe ser un número (tiene «{delta}»)")
        extra += ["--delta", delta]
    campos = texto(f.get(C["campos"])).lower() or "por defecto"
    if campos == "todos los campos":
        extra.append("--todo")
    elif campos == "solo la lista":
        lista = texto(f.get(C["lista"]))
        if not lista:
            raise DatoInvalido("«Solo la lista» sin «Lista de campos»")
        extra += ["--campos", lista]
    elif campos != "por defecto":
        raise DatoInvalido(f"«Campos a modificar» «{texto(f.get(C['campos']))}» no válido")
    if campos != "solo la lista" and texto(f.get(C["lista"])):
        avisos.append("se ignoró la lista de campos (solo aplica con «Solo la lista»)")
    if xc.es_si(f.get(C["filas"])):
        extra.append("--agregar-filas")
    if xc.es_si(f.get(C["ilimitado"])):
        extra.append("--quitar-ilimitado")
    if texto(f.get(C["canales"])).lower() == "no":
        extra.append("--sin-canales")
    fecha = xc.leer_fecha(f.get(C["fecha"]), C["fecha"])
    if fecha and accion != "cambio":
        avisos.append("se ignoró la fecha de ejecución (solo aplica al cambio)")
    elif fecha:
        if fecha <= date.today():
            raise DatoInvalido("la fecha de ejecución debe ser posterior a hoy")
        if fecha < date.today() + timedelta(days=2):
            avisos.append("fecha con menos de 2 días: SIRT puede rechazarla si el cambio perjudica al usuario")
        extra += ["--fecha", fecha.isoformat()]
    if texto(f.get(C["sello"])):
        extra += ["--texto", texto(f[C["sello"]])]
    return {"nombre": nombre, "accion": accion, "extra": extra, "avisos": avisos}


def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    libro = xc.abrir_libro(ruta, args.solo_listar)
    hoja = xc.Hoja(libro[HOJA], obligatorias=["Nombre de la tarifa", *C.values()])

    trabajos, errores, omitidas = [], [], []
    for r, f in hoja.filas("Nombre de la tarifa"):
        nombre = texto(f["Nombre de la tarifa"])
        if args.nombres and nombre not in args.nombres:
            continue
        ej = xc.ejecutar_fila(f)
        if ej == "no":
            omitidas.append(r)
            continue
        if ej != "si":
            errores.append((r, nombre, f"«Ejecutar» debe ser Sí o No (tiene «{ej}»)"))
            continue
        try:
            d = construir(f)
        except DatoInvalido as problema:
            errores.append((r, nombre, str(problema)))
            continue
        d["fila"] = r
        trabajos.append(d)
    if args.limite:
        trabajos = trabajos[: args.limite]

    print(f"Excel: {ruta}")
    print(f"A aplicar: {len(trabajos)} · omitidas: {len(omitidas)} · con errores de datos: {len(errores)}"
          + ("   (--hasta-resumen: no se guardará nada en SIRT)" if args.hasta_resumen else ""))
    for r, nombre, problema in errores:
        print(f"  ✖ fila {r} ({nombre}): {problema}")
    for k, d in enumerate(trabajos, start=1):
        print(f"  {k:>2}. fila {d['fila']}: {d['nombre']}  →  {d['accion']} {' '.join(d['extra'])}")
        for a in d["avisos"]:
            print(f"        aviso: {a}")
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT ni el Excel.")
        return

    for r in omitidas:
        hoja.escribir(r, "Omitida")
    for r, _, problema in errores:
        hoja.escribir(r, "Error de datos", detalle=problema)
    if not trabajos:
        xc.guardar(libro, ruta)
        print("\nNo hay modificaciones para aplicar.")
        return
    if not args.hasta_resumen and not xc.confirmar(
            f"se aplicarán {len(trabajos)} rectificación(es)/cambio(s) en SIRT (no se pueden deshacer)", args.si):
        print("Cancelado.")
        return

    resumen = {}
    for k, d in enumerate(trabajos, start=1):
        print("\n" + "=" * 70)
        print(f"[{k}/{len(trabajos)}] fila {d['fila']}: {d['nombre']} → {d['accion']}")
        desde = xc.marca_tiempo()
        salida = xc.correr(["acciones.py", d["accion"], "--tarifas", d["nombre"], "--si", *d["extra"]]
                           + (["--hasta-resumen"] if args.hasta_resumen else [])
                           + (["--repetir"] if args.repetir or args.hasta_resumen else [])
                           + (["--diagnostico"] if args.diagnostico else []))
        fila = None
        if CSV_ACCIONES.exists():
            import csv
            with CSV_ACCIONES.open("r", newline="", encoding="utf-8-sig") as entrada:
                filas = [x for x in csv.DictReader(entrada, delimiter=";") if x.get("accion") == d["accion"]
                         and x.get("nombre") == d["nombre"] and (x.get("fecha_hora") or "") >= desde]
            fila = filas[-1] if filas else None
        previo = " | ".join(d["avisos"])
        if fila:
            estado = ESTADOS.get(fila.get("resultado", ""), fila.get("resultado", "Error"))
            hoja.escribir(d["fila"], estado, fila.get("codigo_sirt", ""), fila.get("fecha_hora", ""),
                          (previo + " | " if previo else "") + fila.get("detalle", ""))
        elif "se omite" in salida:
            estado = "Ya aplicada"
            hoja.escribir(d["fila"], estado, detalle="Esta acción ya figura como EXITOSA para la tarifa "
                                                     "(para repetirla: --repetir).")
        else:
            estado = "Error"
            hoja.escribir(d["fila"], estado, detalle="acciones.py terminó sin registrar resultado: " + xc.resumen_salida(salida))
        resumen[estado] = resumen.get(estado, 0) + 1
        xc.guardar(libro, ruta)

    copia = xc.guardar_con_copia(libro, ruta, "MODIFICACIONES_EXCEL", "modificaciones_excel")
    print("\n" + "=" * 70)
    print("Resumen: " + " · ".join(f"{k}: {v}" for k, v in resumen.items()))
    print(f"Excel actualizado: {ruta}\nCopia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
