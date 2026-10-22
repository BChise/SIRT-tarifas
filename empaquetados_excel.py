"""
Registra EMPAQUETADOS según automatizacion/empaquetados.xlsx (hoja «Empaquetados»).

Cada fila = un paquete. Para cada fila con «Ejecutar» = Sí (o vacío):
1. Valida: nombre, 2 a 6 tarifas base distintas, valores numéricos.
2. Lo registra con empaquetar.py (el mismo código ya probado), con el alcance, las fechas, la tarifa
   de instalación, los datos de contacto y los valores de la fila. Valor vacío = valor de la tarifa
   base − 1. empaquetar.py revisa antes las reglas de SIRT (registro definitivo de las bases, mismo
   alcance, suma de valores) y, si no se cumplen, no toca SIRT y la fila queda en «Error».
3. Si «Registro definitivo» = Sí, hace el registro definitivo (disponibilidad: la de la fila, que
   debe estar en la intersección de las de sus tarifas; vacío = la intersección completa).
4. Escribe Estado, Código SIRT (n.º de registro), Fecha y Detalle; copia en resultados/EMPAQUETADOS_EXCEL/.

Uso:
    python empaquetados_excel.py --solo-listar
    python empaquetados_excel.py --nombres TE-EMP-CM-001
    python empaquetados_excel.py --si --diagnostico
    python empaquetados_excel.py --sin-definitivo
"""

import argparse
from datetime import date
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "empaquetados.xlsx"
HOJA = "Empaquetados"
MAX_TARIFAS = 6
ERRORES_SEGUIDOS_MAX = 3       # tras 3 errores seguidos en la misma etapa, se detiene


def leer_argumentos():
    p = argparse.ArgumentParser(description="Empaquetados desde el Excel.")
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


def valor_base(nombre: str) -> str | None:
    """Valor con que se registró la tarifa base (Excel) o el fijo de su servicio."""
    import comun
    from empaquetar import buscar_tarifa_base, servicios_por_codigo
    registrado = comun.valor_tarifa_registrado(nombre)
    if registrado:
        return registrado
    base = buscar_tarifa_base(nombre)
    if base is None:
        return None
    modulo = servicios_por_codigo().get(base["codigo"])
    return modulo.VALORES.get("valor_servicio") if modulo else None


def construir(f: dict) -> dict:
    nombre = texto(f["Nombre del paquete"])
    tarifas, valores = [], []
    for n in range(1, MAX_TARIFAS + 1):
        t = texto(f.get(f"Tarifa {n}"))
        v = texto(f.get(f"Valor {n} en el paquete"))
        if not t:
            if v:
                raise DatoInvalido(f"«Valor {n}» tiene dato pero «Tarifa {n}» está vacía")
            continue
        tarifas.append(t)
        valores.append(xc.numero(v, f"Valor {n} en el paquete", minimo=0) if v else "")
    if len(tarifas) < 2:
        raise DatoInvalido("un paquete necesita al menos 2 tarifas")
    if len(set(tarifas)) != len(tarifas):
        raise DatoInvalido("hay una tarifa repetida en el paquete")
    avisos = []
    if any(valores) and not all(valores):
        from empaquetar import valor_por_defecto
        for k, (t, v) in enumerate(zip(tarifas, valores)):
            if not v:
                base = valor_base(t)
                if base is None:
                    raise DatoInvalido(f"no se encontró la tarifa «{t}» en los resultados (¿está registrada?)")
                valores[k] = valor_por_defecto(base)
                avisos.append(f"{t}: valor vacío → {valores[k]} (base {base} − 1)")
    extra = ["--tarifas", *tarifas, "--lote", xc.LOTE_EXCEL, "--nombre", nombre]
    if all(valores):
        extra += ["--valores", *valores]
    alcance = texto(f.get("Alcance de la tarifa"))
    if alcance:
        extra += ["--alcance", alcance]
    hoy = date.today()
    for columna, opcion in (("Inicio de vigencia", "--inicio-vigencia"), ("Inicio comercialización", "--inicio-comercializacion")):
        fecha = xc.leer_fecha(f.get(columna), columna)
        if fecha:
            if fecha < hoy:
                raise DatoInvalido(f"«{columna}» es anterior a hoy")
            extra += [opcion, fecha.isoformat()]
    if texto(f.get("Tarifa instalación")):
        extra += ["--instalacion", xc.numero(f["Tarifa instalación"], "Tarifa instalación", minimo=0)]
    else:
        extra.append("--sin-instalacion")         # vacío = el paquete no lleva tarifa de instalación
    for columna, opcion in (("Otros datos generales", "--otros-datos"), ("Enlace Web", "--enlace"),
                            ("Fono Atención al cliente", "--fono"), ("Email de atención al cliente", "--email")):
        if texto(f.get(columna)):
            extra += [opcion, texto(f[columna])]
    return {"nombre": nombre, "tarifas": tarifas, "valores": valores, "extra": extra, "avisos": avisos,
            "definitivo": xc.es_si(f.get("Registro definitivo")),
            "disponibilidad": texto(f.get("Disponibilidad comercial"))}


def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    libro = xc.abrir_libro(ruta, args.solo_listar)
    hoja = xc.Hoja(libro[HOJA], obligatorias=("Nombre del paquete", "Tarifa 1", "Tarifa 2"))
    csv_emp = xc.RESULTADOS / "EMP" / f"resultados_{xc.LOTE_EXCEL}.csv"
    csv_def = xc.RESULTADOS / f"registro_definitivo_{xc.LOTE_EXCEL}.csv"

    trabajos, errores, omitidas, ya = [], [], [], []
    for r, f in hoja.filas("Nombre del paquete"):
        nombre = texto(f["Nombre del paquete"])
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
        previo = xc.exitosa_en(csv_emp, nombre)
        if previo and not args.repetir:
            d["previo"] = previo
            ya.append(d)
        else:
            trabajos.append(d)
    if args.limite:
        trabajos = trabajos[: args.limite]
    pend_def = [d for d in ya if d["definitivo"] and not args.sin_definitivo and not xc.exitosa_en(csv_def, d["nombre"])]

    print(f"Excel: {ruta}")
    print(f"A registrar: {len(trabajos)} · ya registrados: {len(ya)} · omitidos: {len(omitidas)} · con errores de datos: {len(errores)}")
    for r, nombre, problema in errores:
        print(f"  ✖ fila {r} ({nombre}): {problema}")
    for k, d in enumerate(trabajos, start=1):
        vals = " + ".join(f"{t} ({v or 'base − 1'})" for t, v in zip(d["tarifas"], d["valores"]))
        print(f"  {k:>2}. fila {d['fila']}: {d['nombre']} = {vals}"
              + ("  + registro definitivo" if d["definitivo"] and not args.sin_definitivo else ""))
        for a in d["avisos"]:
            print(f"        aviso: {a}")
    for d in pend_def:
        print(f"   · {d['nombre']}: ya registrado; falta su registro definitivo")
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
                          "Ya figura como registrado en el lote EXCEL (no se repitió).")
    if not trabajos and not pend_def:
        xc.guardar(libro, ruta)
        print("\nNo hay paquetes para registrar.")
        return
    if not xc.confirmar(f"se registrarán {len(trabajos)} paquete(s) CONFIRMADOS en SIRT", args.si):
        print("Cancelado.")
        return

    resumen = {}
    etapas_con_error = []          # para detenerse si SIRT deja de responder (p. ej. sesión vencida)
    for k, d in enumerate(trabajos, start=1):
        if len(etapas_con_error) >= ERRORES_SEGUIDOS_MAX and len(set(etapas_con_error[-ERRORES_SEGUIDOS_MAX:])) == 1:
            print("\n" + "!" * 70)
            print(f"Se detiene: {ERRORES_SEGUIDOS_MAX} paquetes seguidos fallaron en la misma etapa "
                  f"({etapas_con_error[-1]}). Revisa en Edge que la sesión de SIRT siga activa (recarga con F5"
                  " o vuelve a iniciar sesión) y corre de nuevo el comando: continúa con lo pendiente.")
            print("!" * 70)
            break
        print("\n" + "=" * 70)
        print(f"[{k}/{len(trabajos)}] fila {d['fila']}: {d['nombre']}")
        desde = xc.marca_tiempo()
        salida = xc.correr(["empaquetar.py", *d["extra"]] + (["--diagnostico"] if args.diagnostico else []))
        fila = xc.ultimo_resultado(csv_emp, d["nombre"], desde)
        detalle = " | ".join(d["avisos"])
        if fila and fila.get("resultado") == "EXITOSO":
            hoja.escribir(d["fila"], "Registrada", fila.get("id_registro", ""), fila.get("fecha_hora", ""),
                          (detalle + " | " if detalle else "") + "paquete confirmado")
            d["registrado"] = True
            estado = "Registrada"
            etapas_con_error = []
        else:
            motivo = fila.get("detalle", "") if fila else xc.resumen_salida(salida)
            hoja.escribir(d["fila"], "Error", fila.get("id_registro", "") if fila else "", fila.get("fecha_hora", "") if fila else "",
                          (detalle + " | " if detalle else "") + motivo)
            estado = "Error"
            etapas_con_error.append(motivo.split("]")[0] + "]" if motivo.startswith("[") else "sin etapa")
        resumen[estado] = resumen.get(estado, 0) + 1
        xc.guardar(libro, ruta)

    if not args.sin_definitivo:
        para_definitivo = [t for t in trabajos if t.get("registrado") and t["definitivo"]] + pend_def
        if args.definitivo_en_bloque:
            xc.precargar_definitivos([(d["nombre"], d["disponibilidad"]) for d in para_definitivo],
                                     xc.LOTE_EXCEL, args.tamano_bloque, args.diagnostico)
        for d in para_definitivo:
            ok, det = xc.hacer_registro_definitivo(d["nombre"], xc.LOTE_EXCEL, d["disponibilidad"], args.diagnostico)
            codigo = xc.texto(hoja.hoja.cell(d["fila"], hoja.c_codigo).value) or d.get("previo", {}).get("id_registro", "")
            actual = xc.texto(hoja.hoja.cell(d["fila"], hoja.c_detalle).value)
            hoja.escribir(d["fila"], "Registro definitivo" if ok else "Error", codigo, xc.marca_tiempo(),
                          (actual + " | " if actual else "") + det)
            clave = "Registro definitivo" if ok else "Error en registro definitivo"
            resumen[clave] = resumen.get(clave, 0) + 1
            xc.guardar(libro, ruta)

    copia = xc.guardar_con_copia(libro, ruta, "EMPAQUETADOS_EXCEL", "empaquetados_excel")
    print("\n" + "=" * 70)
    print("Resumen: " + (" · ".join(f"{k}: {v}" for k, v in resumen.items()) or "sin cambios"))
    print(f"Tiempo total: {xc.tiempo_total(len(trabajos))}")
    print(f"Excel actualizado: {ruta}\nCopia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
