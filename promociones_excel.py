"""
Registra TARIFAS PROMOCIONALES según automatizacion/promociones.xlsx.

Hoja «Promociones»: una fila por promoción (1 a 4 tarifas base con registro definitivo).
Hoja «Atributos»: solo para las filas con «Atributos a promocionar» = «Detalle en hoja Atributos»
(un renglón por atributo, con su servicio, tipo de variación, variación y duración).

Para cada fila con «Ejecutar» = Sí (o vacío): valida, la registra con promocionar.py (el mismo
código ya probado) pasándole los datos de la fila, y si «Registro definitivo» = Sí hace el registro
definitivo (sin disponibilidad comercial). Escribe Estado, Código SIRT, Fecha y Detalle y guarda una
copia en resultados/PROMOCIONES_EXCEL/.

Uso:
    python promociones_excel.py --solo-listar
    python promociones_excel.py --nombres TP-IF-CM-001
    python promociones_excel.py --si --diagnostico
"""

import argparse
import json
from datetime import date
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "promociones.xlsx"
HOJA, HOJA_ATRIBUTOS = "Promociones", "Atributos"
TIPOS = {"porcentual": "Porcentual", "cantidad": "Cantidad", "ilimitado": "Ilimitado"}
MODOS = {"los de siempre": "probados", "todos": "todos", "lista": "lista", "detalle en hoja atributos": "detalle"}
DURACION_MAXIMA = 180


def leer_argumentos():
    p = argparse.ArgumentParser(description="Promociones desde el Excel.")
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


def variacion(tipo, valor, duracion, campo: str, con_duracion=True) -> dict:
    """Valida tipo / variación / duración de una fila de promoción."""
    t = TIPOS.get(texto(tipo).lower()) if texto(tipo) else None
    if texto(tipo) and not t:
        raise DatoInvalido(f"{campo}: tipo de variación «{texto(tipo)}» no válido (Porcentual, Cantidad o Ilimitado)")
    datos = {"tipo": t}
    if t != "Ilimitado" and texto(valor):
        datos["variacion"] = xc.numero(valor, f"{campo}: variación", minimo=0,
                                       maximo=100 if t == "Porcentual" else None)
    elif t in ("Porcentual", "Cantidad"):
        raise DatoInvalido(f"{campo}: falta la variación")
    if con_duracion and texto(duracion):
        datos["duracion"] = xc.numero(duracion, f"{campo}: duración", minimo=1, maximo=DURACION_MAXIMA, entero=True)
    return datos


def leer_atributos(libro) -> dict:
    """{promoción: {servicio: [{nombre, tipo, variacion, duracion}]}} de la hoja «Atributos»."""
    if HOJA_ATRIBUTOS not in libro.sheetnames:
        return {}
    hoja = xc.Hoja(libro[HOJA_ATRIBUTOS], obligatorias=("Nombre de la promoción", "Servicio", "Atributo"))
    salida = {}
    for r, f in hoja.filas("Nombre de la promoción"):
        promo, servicio, atributo = texto(f["Nombre de la promoción"]), texto(f["Servicio"]).upper(), texto(f["Atributo"])
        if not servicio or not atributo:
            raise DatoInvalido(f"hoja «Atributos», fila {r}: faltan el servicio o el atributo")
        datos = variacion(f.get("Tipo de variación"), f.get("Variación"), f.get("Duración (días)"),
                          f"«Atributos» fila {r} ({atributo})")
        salida.setdefault(promo, {}).setdefault(servicio, []).append({"nombre": atributo, **datos})
    return salida


def construir(f: dict, atributos: dict) -> dict:
    nombre = texto(f["Nombre de la promoción"])
    columnas = sorted((c for c in f if c.startswith("Tarifa base ") and c[12:].isdigit()), key=lambda c: int(c[12:]))
    tarifas = [texto(f.get(c)) for c in columnas if texto(f.get(c))]
    if not tarifas:
        raise DatoInvalido("falta al menos «Tarifa base 1»")
    if len(set(tarifas)) != len(tarifas):
        raise DatoInvalido("hay una tarifa base repetida")
    config, extra = {}, ["--tarifas", *tarifas, "--lote", xc.LOTE_EXCEL, "--nombre", nombre]

    inicio = xc.leer_fecha(f.get("Inicio comercialización"), "Inicio comercialización")
    if inicio:
        if inicio < date.today():
            raise DatoInvalido("«Inicio comercialización» es anterior a hoy")
        config["inicio_comercializacion"] = inicio.isoformat()
    if texto(f.get("Días de comercialización")):
        extra += ["--dias-comercializacion",
                  xc.numero(f["Días de comercialización"], "Días de comercialización", minimo=1, entero=True)]
    for columna, clave in (("Condiciones generales", "condiciones"), ("Otras condiciones tarifa", "otras_condiciones")):
        if texto(f.get(columna)):
            config[clave] = texto(f[columna])

    def se_promociona(prefijo):
        marcado = texto(f.get(f"{prefijo}: promocionar"))
        if marcado:
            return xc.es_si(marcado)
        return bool(texto(f.get(f"{prefijo}: tipo de variación")) or texto(f.get(f"{prefijo}: variación")))

    if se_promociona("Valor del plan"):
        config["valor_plan"] = variacion(f.get("Valor del plan: tipo de variación"), f.get("Valor del plan: variación"),
                                         f.get("Valor del plan: duración (días)"), "Valor del plan")
    else:
        extra.append("--sin-valor-plan")
    if se_promociona("Tarifa instalación"):
        config["instalacion"] = variacion(f.get("Tarifa instalación: tipo de variación"),
                                          f.get("Tarifa instalación: variación"), None, "Tarifa instalación", False)
    else:
        extra.append("--sin-instalacion")

    for columna, opcion, maximo in (("Variación % por defecto", "--variacion-porcentual", 100),
                                    ("Variación en cantidad por defecto", "--variacion-cantidad", None),
                                    ("Duración por defecto (días)", "--duracion", DURACION_MAXIMA)):
        if texto(f.get(columna)):
            extra += [opcion, xc.numero(f[columna], columna, minimo=0 if opcion != "--duracion" else 1,
                                        maximo=maximo, entero=opcion == "--duracion")]

    preferido = texto(f.get("Tipo de variación preferido"))
    if preferido:
        tipo = TIPOS.get(preferido.lower())
        if not tipo:
            raise DatoInvalido(f"«Tipo de variación preferido» «{preferido}» no válido (Porcentual, Cantidad o Ilimitado)")
        extra += ["--tipo-preferido", tipo]

    modo_texto = texto(f.get("Atributos a promocionar")) or "Los de siempre"
    modo = MODOS.get(modo_texto.lower())
    if not modo:
        raise DatoInvalido(f"«Atributos a promocionar» «{modo_texto}» no válido")
    descripcion = modo_texto
    if modo in ("probados", "todos"):
        extra += ["--atributos", modo]
    elif modo == "lista":
        lista = [a.strip() for a in texto(f.get("Lista de atributos (separados por ;)")).split(";") if a.strip()]
        if not lista:
            raise DatoInvalido("modo «Lista» sin «Lista de atributos»")
        extra += ["--atributos", ";".join(lista)]
        descripcion = "; ".join(lista)
    else:
        propios = atributos.get(nombre)
        if not propios:
            raise DatoInvalido("modo «Detalle en hoja Atributos» pero la hoja «Atributos» no tiene filas de esta promoción")
        config["atributos_por_servicio"] = propios
        descripcion = "; ".join(f"{s}: {', '.join(a['nombre'] for a in lista)}" for s, lista in propios.items())
    if preferido and modo != "detalle":
        descripcion += f" (preferido: {TIPOS[preferido.lower()]})"
    return {"nombre": nombre, "tarifas": tarifas, "config": config, "extra": extra, "descripcion": descripcion,
            "definitivo": xc.es_si(f.get("Registro definitivo"))}


def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    libro = xc.abrir_libro(ruta, args.solo_listar)
    hoja = xc.Hoja(libro[HOJA], obligatorias=("Nombre de la promoción", "Tarifa base 1"))
    csv_prm = xc.RESULTADOS / "PRM" / f"resultados_{xc.LOTE_EXCEL}.csv"
    csv_def = xc.RESULTADOS / f"registro_definitivo_{xc.LOTE_EXCEL}.csv"
    try:
        atributos = leer_atributos(libro)
        error_atributos = ""
    except DatoInvalido as problema:
        atributos, error_atributos = {}, str(problema)

    trabajos, errores, omitidas, ya = [], [], [], []
    for r, f in hoja.filas("Nombre de la promoción"):
        nombre = texto(f["Nombre de la promoción"])
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
            if error_atributos and "detalle" in texto(f.get("Atributos a promocionar")).lower():
                raise DatoInvalido(error_atributos)
            d = construir(f, atributos)
        except DatoInvalido as problema:
            errores.append((r, nombre, str(problema)))
            continue
        d["fila"] = r
        previo = xc.exitosa_en(csv_prm, nombre)
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
        vp = d["config"].get("valor_plan")
        print(f"  {k:>2}. fila {d['fila']}: {d['nombre']} ← {' + '.join(d['tarifas'])}"
              + ("  + registro definitivo" if d["definitivo"] and not args.sin_definitivo else ""))
        print(f"        valor del plan: {'no' if vp is None else (vp.get('tipo') or 'tipo por defecto') + ' ' + vp.get('variacion', '')}"
              f" · atributos: {d['descripcion']}")
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
    if not trabajos and not pend_def:
        xc.guardar(libro, ruta)
        print("\nNo hay promociones para registrar.")
        return
    if not xc.confirmar(f"se registrarán {len(trabajos)} promoción(es) CONFIRMADAS en SIRT", args.si):
        print("Cancelado.")
        return

    carpeta_config = xc.RESULTADOS / "PRM" / "config_excel"
    carpeta_config.mkdir(parents=True, exist_ok=True)
    resumen = {}
    for k, d in enumerate(trabajos, start=1):
        print("\n" + "=" * 70)
        print(f"[{k}/{len(trabajos)}] fila {d['fila']}: {d['nombre']}")
        archivo = carpeta_config / f"{d['nombre']}.json"
        archivo.write_text(json.dumps(d["config"], ensure_ascii=False, indent=2), encoding="utf-8")
        desde = xc.marca_tiempo()
        salida = xc.correr(["promocionar.py", *d["extra"], "--config", str(archivo)]
                           + (["--diagnostico"] if args.diagnostico else []))
        fila = xc.ultimo_resultado(csv_prm, d["nombre"], desde)
        if fila and fila.get("resultado") == "EXITOSO":
            hoja.escribir(d["fila"], "Registrada", fila.get("id_registro", ""), fila.get("fecha_hora", ""), fila.get("detalle", ""))
            d["registrado"] = True
            estado = "Registrada"
        else:
            motivo = fila.get("detalle", "") if fila else xc.resumen_salida(salida)
            hoja.escribir(d["fila"], "Error", fila.get("id_registro", "") if fila else "",
                          fila.get("fecha_hora", "") if fila else "", motivo)
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

    copia = xc.guardar_con_copia(libro, ruta, "PROMOCIONES_EXCEL", "promociones_excel")
    print("\n" + "=" * 70)
    print("Resumen: " + (" · ".join(f"{k}: {v}" for k, v in resumen.items()) or "sin cambios"))
    print(f"Tiempo total: {xc.tiempo_total(len(trabajos))}")
    print(f"Excel actualizado: {ruta}\nCopia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
