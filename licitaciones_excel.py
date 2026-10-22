"""
Registra tarifas de LICITACIÓN Y CONTRATOS según automatizacion/registro_licitaciones.xlsx.

Cada fila = una tarifa (1 a 6 servicios con «<servicio>: incluir» = Sí y su detalle). Para cada
fila con «Ejecutar» = Sí (o vacío): valida, la registra con licitacion.py (el mismo código ya
probado; fin de vigencia = inicio + duración − 1), y si «Registro definitivo» = Sí hace el registro
definitivo (sin disponibilidad comercial). Escribe Estado, Código SIRT, Fecha y Detalle en la fila y
guarda una copia en resultados/LICITACIONES_EXCEL/.

Uso:
    python licitaciones_excel.py --solo-listar
    python licitaciones_excel.py --nombres TL-LYC-CM-001
    python licitaciones_excel.py --si --diagnostico
"""

import argparse
from datetime import date
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "registro_licitaciones.xlsx"
HOJA = "Licitaciones"
SERVICIOS = ["Arrendamiento de circuitos", "Otros servicios", "Televisión de paga",
             "Internet Fijo", "Servicio Móvil", "Telefonía Fija"]
OBLIGATORIAS = ["Duración del contrato (días)", "Identificador de contrato", "Ámbito", "Valor de contrato",
                "Moneda", "Documento de Empresa", "Nombre de empresa", "Descripción del Servicio",
                "Atributos de la tarifa", "Ámbito geográfico"]


def leer_argumentos():
    p = argparse.ArgumentParser(description="Licitación y Contratos desde el Excel.")
    p.add_argument("--excel", default=str(EXCEL))
    p.add_argument("--nombres", nargs="+")
    p.add_argument("--limite", type=int, default=0)
    p.add_argument("--repetir", action="store_true")
    p.add_argument("--sin-definitivo", action="store_true")
    p.add_argument("--definitivo-en-bloque", action="store_true",
                   help="registro definitivo de varias tarifas a la vez (lo que falle sigue una por una)")
    p.add_argument("--tamano-bloque", type=int, default=10)
    p.add_argument("--solo-listar", action="store_true")
    p.add_argument("--si", action="store_true")
    p.add_argument("--diagnostico", action="store_true")
    return p.parse_args()


def validar(f: dict) -> dict:
    nombre = texto(f["Nombre de la tarifa"])
    faltan = [c for c in OBLIGATORIAS if not texto(f.get(c))]
    if faltan:
        raise DatoInvalido("faltan: " + ", ".join(faltan))
    duracion = int(xc.numero(f["Duración del contrato (días)"], "Duración del contrato (días)", minimo=1, entero=True))
    xc.numero(f["Valor de contrato"], "Valor de contrato", minimo=0)
    inicio = xc.leer_fecha(f.get("Inicio de vigencia"), "Inicio de vigencia")
    if inicio and inicio < date.today():
        raise DatoInvalido("«Inicio de vigencia» es anterior a hoy")
    servicios = [s for s in SERVICIOS if xc.es_si(f.get(f"{s}: incluir"))]
    if not servicios:
        raise DatoInvalido("marca al menos un servicio con «incluir» = Sí")
    sin_detalle = [s for s in servicios if not texto(f.get(f"{s}: detalle del servicio"))]
    if sin_detalle:
        raise DatoInvalido("falta el «detalle del servicio» de: " + ", ".join(sin_detalle))
    avisos = []
    if texto(f.get("Tarifa instalación (solo con Otros servicios)")) and "Otros servicios" not in servicios:
        avisos.append("se ignoró la tarifa de instalación (solo aplica con Otros servicios)")
    return {"nombre": nombre, "servicios": servicios, "duracion": duracion, "inicio": inicio or date.today(),
            "avisos": avisos, "definitivo": xc.es_si(f.get("Registro definitivo"))}


def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    libro = xc.abrir_libro(ruta, args.solo_listar)
    hoja = xc.Hoja(libro[HOJA], obligatorias=["Nombre de la tarifa", *OBLIGATORIAS])
    csv_lyc = xc.RESULTADOS / "LYC" / f"resultados_{xc.LOTE_EXCEL}.csv"
    csv_def = xc.RESULTADOS / f"registro_definitivo_{xc.LOTE_EXCEL}.csv"

    trabajos, errores, omitidas, ya = [], [], [], []
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
            d = validar(f)
        except DatoInvalido as problema:
            errores.append((r, nombre, str(problema)))
            continue
        d["fila"] = r
        previo = xc.exitosa_en(csv_lyc, nombre)
        if previo and not args.repetir:
            d["previo"] = previo
            ya.append(d)
        else:
            trabajos.append(d)
    if args.limite:
        trabajos = trabajos[: args.limite]
    pend_def = [d for d in ya if d["definitivo"] and not args.sin_definitivo and not xc.exitosa_en(csv_def, d["nombre"])]

    print(f"Excel: {ruta}")
    print(f"A registrar: {len(trabajos)} · ya registradas: {len(ya)} · omitidas: {len(omitidas)} · con errores de datos: {len(errores)}")
    for r, nombre, problema in errores:
        print(f"  ✖ fila {r} ({nombre}): {problema}")
    for k, d in enumerate(trabajos, start=1):
        print(f"  {k:>2}. fila {d['fila']}: {d['nombre']}  {len(d['servicios'])} serv.: {' + '.join(d['servicios'])}"
              f"  · vigencia {d['inicio']:%d/%m/%Y} + {d['duracion']} días"
              + ("  + registro definitivo" if d["definitivo"] and not args.sin_definitivo else ""))
        for a in d["avisos"]:
            print(f"        aviso: {a}")
    for d in pend_def:
        print(f"   · {d['nombre']}: ya registrada; falta su registro definitivo")
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT ni el Excel.")
        return

    for r in omitidas:
        hoja.escribir(r, "Omitida")
    for r, _, problema in errores:
        hoja.escribir(r, "Error de datos", detalle=problema)
    for d in ya:
        if hoja.estado(d["fila"]) not in ("Registrada", "Registro definitivo"):
            hoja.escribir(d["fila"], "Ya registrada", d["previo"].get("id_registro", ""), d["previo"].get("fecha_hora", ""),
                          "Ya figura como registrada en el lote EXCEL (no se repitió).")
    xc.guardar(libro, ruta)       # licitacion.py lee este mismo archivo
    if not trabajos and not pend_def:
        print("\nNo hay tarifas para registrar.")
        return
    if not xc.confirmar(f"se registrarán {len(trabajos)} tarifa(s) de Licitación y Contratos CONFIRMADAS en SIRT", args.si):
        print("Cancelado.")
        return

    resumen = {}
    for k, d in enumerate(trabajos, start=1):
        print("\n" + "=" * 70)
        print(f"[{k}/{len(trabajos)}] fila {d['fila']}: {d['nombre']}")
        desde = xc.marca_tiempo()
        salida = xc.correr(["licitacion.py", "--excel", str(ruta), "--nombres", d["nombre"]]
                           + (["--repetir"] if args.repetir else []) + (["--diagnostico"] if args.diagnostico else []))
        fila = xc.ultimo_resultado(csv_lyc, d["nombre"], desde)
        previo = " | ".join(d["avisos"])
        if fila and fila.get("resultado") == "EXITOSO":
            hoja.escribir(d["fila"], "Registrada", fila.get("id_registro", ""), fila.get("fecha_hora", ""),
                          (previo + " | " if previo else "") + fila.get("detalle", "confirmada"))
            d["registrado"] = True
            estado = "Registrada"
        else:
            motivo = fila.get("detalle", "") if fila else xc.resumen_salida(salida)
            hoja.escribir(d["fila"], "Error", fila.get("id_registro", "") if fila else "",
                          fila.get("fecha_hora", "") if fila else "", (previo + " | " if previo else "") + motivo)
            estado = "Error"
        resumen[estado] = resumen.get(estado, 0) + 1
        xc.guardar(libro, ruta)

    if not args.sin_definitivo:
        para_definitivo = [t for t in trabajos if t.get("registrado") and t["definitivo"]] + pend_def
        if args.definitivo_en_bloque:
            xc.precargar_definitivos([(d["nombre"], "") for d in para_definitivo],
                                     xc.LOTE_EXCEL, args.tamano_bloque, args.diagnostico)
        for d in para_definitivo:
            ok, det = xc.hacer_registro_definitivo(d["nombre"], xc.LOTE_EXCEL, "", args.diagnostico)
            codigo = xc.texto(hoja.hoja.cell(d["fila"], hoja.c_codigo).value) or d.get("previo", {}).get("id_registro", "")
            actual = xc.texto(hoja.hoja.cell(d["fila"], hoja.c_detalle).value)
            hoja.escribir(d["fila"], "Registro definitivo" if ok else "Error", codigo, xc.marca_tiempo(),
                          (actual + " | " if actual else "") + det)
            clave = "Registro definitivo" if ok else "Error en registro definitivo"
            resumen[clave] = resumen.get(clave, 0) + 1
            xc.guardar(libro, ruta)

    copia = xc.guardar_con_copia(libro, ruta, "LICITACIONES_EXCEL", "licitaciones_excel")
    print("\n" + "=" * 70)
    print("Resumen: " + (" · ".join(f"{k}: {v}" for k, v in resumen.items()) or "sin cambios"))
    print(f"Tiempo total: {xc.tiempo_total(len(trabajos))}")
    print(f"Excel actualizado: {ruta}\nCopia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
