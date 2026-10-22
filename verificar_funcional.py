"""
Verifica la PRUEBA FUNCIONAL: en cada Excel de automatizacion/pruebas_funcionales/ compara el
«Estado» que dejó el script con el «Resultado esperado» de la fila y arma la matriz de cobertura
(una línea por fila: qué se prueba → OK / FALLA / PENDIENTE). La guarda también en
resultados/verificacion_funcional_<fecha>.csv para la presentación.

Si un Excel no tiene «Resultado esperado» (p. ej. modificaciones_funcional), se deduce:
«ERROR esperado…» → Error de datos; «Ejecutar» = No → Omitida; el resto → Aplicada.

Uso:
    python verificar_funcional.py                      (Excel de pruebas_funcionales)
    python verificar_funcional.py --carpeta Prueba4    (Excel de pruebas_funcionales/Prueba4)
"""

import csv
from datetime import datetime

import excel_comun as xc

CARPETA = xc.AUTOMATIZACION / "pruebas_funcionales"
EQUIVALENTES = {"Aplicada": {"Aplicada", "Ya aplicada"}, "Registro definitivo": {"Registro definitivo"},
                "Registrada": {"Registrada", "Ya registrada"}}


def evaluar(ruta) -> list[dict]:
    """Filas de un Excel de prueba con su veredicto (OK / FALLA / PENDIENTE)."""
    filas = []
    libro = xc.abrir_libro(ruta, True)
    for ws in libro.worksheets:
        hoja = xc.Hoja(ws)
        if not (hoja.c_estado and hoja.c_nombre and "Qué se prueba" in hoja.col):
            continue
        for r in range(xc.FILA_CABECERA + 1, ws.max_row + 1):
            nombre = xc.texto(ws.cell(r, hoja.c_nombre).value)
            if not nombre:
                continue
            prueba = xc.texto(ws.cell(r, hoja.col["Qué se prueba"]).value)
            esperado = xc.texto(ws.cell(r, hoja.col["Resultado esperado"]).value) if "Resultado esperado" in hoja.col \
                else "Error de datos" if prueba.startswith("ERROR") \
                else "Omitida" if xc.texto(ws.cell(r, hoja.col.get("Ejecutar", 2)).value).lower() == "no" \
                else "Aplicada"
            estado = hoja.estado(r)
            if not estado:
                veredicto = "PENDIENTE"
            elif estado in EQUIVALENTES.get(esperado, {esperado}):
                veredicto = "OK"
            else:
                veredicto = "FALLA"
            detalle = xc.texto(ws.cell(r, hoja.c_detalle).value)[:150] if hoja.c_detalle and veredicto == "FALLA" else ""
            filas.append({"excel": ruta.stem, "hoja": ws.title, "nombre": nombre, "que_se_prueba": prueba,
                          "esperado": esperado, "estado": estado, "veredicto": veredicto, "detalle": detalle})
    return filas


def main() -> None:
    import argparse
    p = argparse.ArgumentParser(description="Verifica los Excel de la prueba funcional.")
    p.add_argument("--carpeta", help="subcarpeta de pruebas_funcionales (p. ej. Prueba4)")
    args = p.parse_args()
    carpeta = CARPETA / args.carpeta if args.carpeta else CARPETA
    if not carpeta.is_dir():
        raise SystemExit(f"No existe {carpeta}")
    print(f"Carpeta: {carpeta}")
    filas, totales = [], {"OK": 0, "FALLA": 0, "PENDIENTE": 0}
    for ruta in sorted(carpeta.glob("*.xlsx")):
        if ruta.name.startswith("~$"):
            continue
        filas += evaluar(ruta)
    for f in filas:
        totales[f["veredicto"]] += 1

    actual = None
    for f in filas:
        if f["excel"] != actual:
            actual = f["excel"]
            print(f"\n=== {actual}")
        marca = {"OK": "✔", "FALLA": "✖", "PENDIENTE": "·"}[f["veredicto"]]
        print(f" {marca} {f['nombre']:<18} {f['que_se_prueba'][:70]:<70} esperado: {f['esperado']:<20} "
              f"estado: {f['estado'] or '—'}")
        if f["detalle"]:
            print(f"      detalle: {f['detalle']}")
    print(f"\nTotal: {len(filas)} · OK {totales['OK']} · FALLA {totales['FALLA']} · PENDIENTE {totales['PENDIENTE']}")
    sufijo = f"_{args.carpeta}" if args.carpeta else ""
    salida = xc.RESULTADOS / f"verificacion_funcional{sufijo}_{datetime.now():%Y%m%d_%H%M}.csv"
    salida.parent.mkdir(parents=True, exist_ok=True)
    with salida.open("w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=list(filas[0]) if filas else ["excel"], delimiter=";")
        escritor.writeheader()
        escritor.writerows(filas)
    print(f"Matriz de cobertura: {salida}")


if __name__ == "__main__":
    main()
