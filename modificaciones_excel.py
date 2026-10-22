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
import json
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "modificaciones.xlsx"
HOJA = "Modificaciones"
HOJA_CAMPOS = "Campos"
SENTIDOS = {"beneficioso": "beneficioso", "perjudicial": "perjudicial"}
CARPETA_DETALLE = xc.RESULTADOS / "ACCIONES" / "detalle_excel"
CSV_ACCIONES = xc.RESULTADOS / "ACCIONES" / "acciones.csv"
ACCIONES = {"rectificación": "rectificar", "rectificacion": "rectificar", "cambio": "cambio"}
ESTADOS = {"EXITOSO": "Aplicada", "ERROR": "Error", "NO_APLICA": "No aplica", "EN_RESUMEN": "Hasta resumen"}
C = {"accion": "Acción", "delta": "Delta (vacío: +1 rectificación, −1 cambio)", "campos": "Campos a modificar",
     "lista": "Lista de campos (separados por ;)", "filas": "Agregar una fila a cada tabla",
     "ilimitado": "Pasar «ilimitado» a limitado", "canales": "Cambiar canales (TV de paga)",
     "fecha": "Cambio: fecha de ejecución", "sello": "Texto del sello (vacío: «Rectificación» / «Cambio»)"}
C_SENTIDO = "Cambio: sentido (Beneficioso / Perjudicial)"     # opcional (columna agregada el 20/10)


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


def leer_campos(libro) -> dict:
    """{tarifa: {campo: {"valor": X} o {"delta": N}}} de la hoja «Campos» (modo «Detalle en hoja Campos»)."""
    if HOJA_CAMPOS not in libro.sheetnames:
        return {}
    hoja = xc.Hoja(libro[HOJA_CAMPOS], obligatorias=("Nombre de la tarifa",))
    titulo_campo = next((t for t in hoja.col if t.startswith("Campo")), None)
    if titulo_campo is None:
        raise DatoInvalido("la hoja «Campos» no tiene la columna «Campo»")
    salida = {}
    for r, f in hoja.filas("Nombre de la tarifa"):
        campo = texto(f.get(titulo_campo))
        if not campo:
            raise DatoInvalido(f"hoja «Campos», fila {r}: falta el nombre del campo")
        valor, delta = texto(f.get("Valor nuevo")), texto(f.get("Delta"))
        if valor and delta:
            raise DatoInvalido(f"hoja «Campos», fila {r} ({campo}): pon «Valor nuevo» o «Delta», no los dos")
        if not valor and not delta:
            raise DatoInvalido(f"hoja «Campos», fila {r} ({campo}): falta «Valor nuevo» o «Delta»")
        if valor:
            try:                                      # número (valor de un campo) …
                regla = {"valor": xc.numero(valor, f"{campo}: valor nuevo", minimo=0)}
            except DatoInvalido:                      # … o texto: opción de una lista (p. ej. Tecnología = 4.5G)
                regla = {"valor": valor}
        else:
            regla = {"delta": texto(delta).replace(",", ".").replace("−", "-")}
        if delta:
            try:
                Decimal(regla["delta"])
            except InvalidOperation:
                raise DatoInvalido(f"hoja «Campos», fila {r} ({campo}): el delta debe ser un número")
        salida.setdefault(texto(f["Nombre de la tarifa"]), {})[campo] = regla
    return salida


def construir(f: dict, campos_detalle: dict | None = None) -> dict:
    nombre = texto(f["Nombre de la tarifa"])
    accion = ACCIONES.get(texto(f.get(C["accion"])).lower())
    if not accion:
        raise DatoInvalido(f"acción «{texto(f.get(C['accion']))}» no válida (Rectificación o Cambio)")
    if nombre.upper().startswith("TP-"):
        raise DatoInvalido("SIRT no permite rectificar ni cambiar promociones")
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
    elif campos == "detalle en hoja campos":
        propios = (campos_detalle or {}).get(nombre)
        if not propios:
            raise DatoInvalido("«Detalle en hoja Campos» pero la hoja «Campos» no tiene filas de esta tarifa")
        CARPETA_DETALLE.mkdir(parents=True, exist_ok=True)
        archivo = CARPETA_DETALLE / f"{nombre}_{accion}.json"
        archivo.write_text(json.dumps(propios, ensure_ascii=False, indent=2), encoding="utf-8")
        extra += ["--detalle", str(archivo)]
    elif campos != "por defecto":
        raise DatoInvalido(f"«Campos a modificar» «{texto(f.get(C['campos']))}» no válido")
    if campos != "solo la lista" and texto(f.get(C["lista"])):
        avisos.append("se ignoró la lista de campos (solo aplica con «Solo la lista»)")
    sentido = texto(f.get(C_SENTIDO)).lower()
    if not delta and campos != "detalle en hoja campos":
        if accion == "rectificar":
            avisos.append("Delta vacío: se aplicó +1")
        elif not sentido:
            avisos.append("Delta y sentido vacíos: se aplicó −1 a todos los campos elegidos "
                          "(escenario mixto: bajan los precios y también los beneficios)")
    if sentido:
        if sentido not in SENTIDOS:
            raise DatoInvalido(f"«{C_SENTIDO}» debe ser Beneficioso o Perjudicial (tiene «{texto(f.get(C_SENTIDO))}»)")
        if accion != "cambio":
            avisos.append("se ignoró el sentido (solo aplica al cambio)")
        else:
            extra += ["--sentido", SENTIDOS[sentido]]
            if sentido == "beneficioso" and xc.es_si(f.get(C["ilimitado"])):
                avisos.append("«ilimitado → limitado» perjudica al usuario aunque el cambio sea beneficioso")
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

    try:
        campos_detalle, error_campos = leer_campos(libro), ""
    except DatoInvalido as problema:
        campos_detalle, error_campos = {}, str(problema)
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
            if error_campos and "detalle" in texto(f.get(C["campos"])).lower():
                raise DatoInvalido(error_campos)
            d = construir(f, campos_detalle)
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
    print(f"Tiempo total: {xc.tiempo_total(len(trabajos))}")
    print(f"Excel actualizado: {ruta}\nCopia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
