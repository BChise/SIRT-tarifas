"""
Llena automatizacion/promociones.xlsx con las promociones de prueba (55):

- 45: tres por servicio (un monoproducto de cada uno de los 15 servicios), con tipo de variación
  preferido Porcentual, Cantidad e Ilimitado (este último solo donde SIRT lo admite).
- 2: cada una con 15 monoproductos, uno de cada servicio.
- 2: cada una con 3 monoproductos y 3 empaquetados.
- 6: una por empaquetado (de 4 y 6 servicios, y de una sola familia).

Todas con: valor del plan e instalación promocionados (Porcentual 10, 30 días), TODOS los
atributos promocionables de cada servicio («Todos»), 60 días de comercialización y registro
definitivo. Solo usa tarifas con registro definitivo (resultados/registro_definitivo_*.csv) y
de alcance «Residencial» (para no mezclar alcances en una misma promoción).

Uso (con promociones.xlsx CERRADO, desde la carpeta del proyecto):
    python automatizacion\\generar_promociones_prueba.py
"""

import csv
import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))
import excel_comun as xc  # noqa: E402

EXCEL = PROYECTO / "automatizacion" / "promociones.xlsx"
REGISTRO = PROYECTO / "automatizacion" / "registro_tarifas_establecidas.xlsx"
SERVICIOS = ["TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "SMP", "SMC", "SMO", "SMQ", "OS", "AC"]
# Empaquetados (de empaquetados.xlsx) elegidos por su variedad; se descartan los que no tengan
# registro definitivo.
EMP_SOLOS = ["TE-EMP-CM-103", "TE-EMP-CM-125", "TE-EMP-CM-160", "TE-EMP-CM-184",   # 4 servicios
             "TE-EMP-CM-197",                                                     # 6 servicios
             "TE-EMP-CM-200"]                                                     # solo TV
EMP_MIXTAS = [["TE-EMP-CM-101", "TE-EMP-CM-198", "TE-EMP-CM-201"],               # 4 serv., solo TF, solo móvil
              ["TE-EMP-CM-110", "TE-EMP-CM-140", "TE-EMP-CM-190"]]
PRIMER_NUMERO = 101


def con_definitivo() -> set[str]:
    nombres = set()
    for ruta in (PROYECTO / "resultados").glob("registro_definitivo_*.csv"):
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            nombres |= {f["nombre"] for f in csv.DictReader(entrada, delimiter=";") if f.get("resultado") == "EXITOSO"}
    return nombres


def residenciales(libro, definitivas: set[str]) -> dict[str, list[str]]:
    salida = {}
    for codigo in SERVICIOS:
        if codigo not in libro.sheetnames:
            continue
        hoja = xc.Hoja(libro[codigo])
        salida[codigo] = [xc.texto(f["Nombre de la tarifa"]) for _, f in hoja.filas("Nombre de la tarifa")
                          if xc.texto(f.get("Alcance de la tarifa")) == "Residencial"
                          and xc.texto(f["Nombre de la tarifa"]) in definitivas]
    return salida


def main() -> None:
    if xc.excel_esta_abierto(EXCEL):
        sys.exit(f"Cierra el Excel antes de generar: {EXCEL}")
    definitivas = con_definitivo()
    bases = residenciales(xc.abrir_libro(REGISTRO, True), definitivas)
    faltan = [c for c in SERVICIOS if not bases.get(c)]
    if faltan:
        print("Aviso: sin tarifa Residencial con registro definitivo para: " + ", ".join(faltan))

    def tarifa(codigo: str, k: int) -> str | None:
        lista = bases.get(codigo) or []
        return lista[k % len(lista)] if lista else None

    promos = []
    for preferido in ("Porcentual", "Cantidad", "Ilimitado"):          # 45: tres por servicio
        for codigo in SERVICIOS:
            if tarifa(codigo, 0):
                promos.append((codigo, [tarifa(codigo, 0)],
                               f"Promoción de {codigo}: todos sus atributos, preferido {preferido}", preferido))
    for k in (1, 2):                                                   # 2: 15 monoproductos
        grupo = [t for t in (tarifa(c, k) for c in SERVICIOS) if t]
        promos.append(("MIX", grupo, f"Promoción de {len(grupo)} monoproductos, uno por servicio", "Porcentual"))
    for k, emps in enumerate(EMP_MIXTAS):                              # 2: 3 monoproductos + 3 empaquetados
        monos = [tarifa(c, 3 + k) for c in (["TF", "SMP", "IF"], ["TVP", "SMC", "OS"])[k]]
        emps = [e for e in emps if e in definitivas]
        promos.append(("MIX", [m for m in monos if m] + emps, "Promoción de 3 monoproductos y 3 empaquetados",
                       "Porcentual"))
    for e in EMP_SOLOS:                                                # 6: un empaquetado
        if e in definitivas:
            promos.append(("EMP", [e], f"Promoción del empaquetado {e}", "Porcentual"))
        else:
            print(f"Aviso: {e} no tiene registro definitivo; se omite.")

    libro = xc.abrir_libro(EXCEL, True)
    hoja = libro["Promociones"]
    col = {xc.texto(c.value): c.column for c in hoja[xc.FILA_CABECERA] if c.value}
    maximo = max(int(t[12:]) for t in col if t.startswith("Tarifa base ") and t[12:].isdigit())
    if max(len(p[1]) for p in promos) > maximo:
        sys.exit(f"El Excel admite {maximo} tarifas por promoción; regenera la plantilla "
                 "(python automatizacion\\generar_plantilla_promociones.py).")
    ultima = max(col.values())
    for r in range(xc.FILA_CABECERA + 1, xc.FILA_CABECERA + 1 + len(promos) + 20):
        for c in range(1, ultima + 1):
            hoja.cell(r, c).value = None
    if "Tipo de variación preferido" not in col:
        sys.exit("El Excel no tiene la columna «Tipo de variación preferido»; regenera la plantilla "
                 "(python automatizacion\\generar_plantilla_promociones.py).")
    for k, (prefijo, tarifas, condiciones, preferido) in enumerate(promos):
        r = xc.FILA_CABECERA + 1 + k
        valores = {
            "N°": k + 1, "Ejecutar": "Sí", "Nombre de la promoción": f"TP-{prefijo}-CM-{PRIMER_NUMERO + k:03d}",
            "Registro definitivo": "Sí", "Días de comercialización": 60, "Condiciones generales": condiciones,
            "Valor del plan: promocionar": "Sí", "Valor del plan: tipo de variación": "Porcentual",
            "Valor del plan: variación": "10", "Valor del plan: duración (días)": "30",
            "Tarifa instalación: promocionar": "Sí", "Tarifa instalación: tipo de variación": "Porcentual",
            "Tarifa instalación: variación": "10", "Atributos a promocionar": "Todos",
            "Variación % por defecto": "10", "Variación en cantidad por defecto": "1",
            "Duración por defecto (días)": "30", "Tipo de variación preferido": preferido,
        }
        for j, t in enumerate(tarifas, start=1):
            valores[f"Tarifa base {j}"] = t
        for titulo, valor in valores.items():
            if titulo in col:
                hoja.cell(r, col[titulo]).value = valor
    libro.save(EXCEL)
    print(f"Listo: {len(promos)} promociones en {EXCEL.name} "
          f"(TP-…-CM-{PRIMER_NUMERO:03d} a TP-…-CM-{PRIMER_NUMERO + len(promos) - 1:03d})")


if __name__ == "__main__":
    main()
