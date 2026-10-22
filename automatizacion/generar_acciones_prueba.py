"""
Llena automatizacion/acciones_basicas.xlsx con las acciones de prueba (81), acordadas el 20/10/2026:

Monoproductos (cada uno de los 15 servicios, en el orden de su hoja del Excel de registro):
  1.ª tarifa: sin acción (queda tal como se registró) · 2.ª: renombre · 3.ª: cese · 4.ª: fe de erratas
  Renombre y cese no se aplican a tarifas que son base de un empaquetado o de una promoción: en ese
  caso se usa la siguiente tarifa del servicio que no lo sea (se avisa al generar).
Empaquetados: 1 renombre, 1 cese y 4 fe de erratas (de 4 servicios, de 6, solo TF y solo móvil).
Promociones: fe de erratas, renovación y renombre: 2 de monoproducto, 2 de empaquetado y 2 mixtas
  cada una (las 2 mixtas reciben las tres; el script las aplica en orden: fe de erratas, renovación,
  renombre).
Licitaciones: fe de erratas a las 6 de un servicio y a 6 de varios servicios (3, 3, 4, 4, 5 y 6).

Uso (con acciones_basicas.xlsx CERRADO):
    python automatizacion\\generar_acciones_prueba.py
Correr las acciones cuando las tarifas ya estén «En comercialización» (el día siguiente al registro
definitivo): python acciones_excel.py --solo-listar, luego --limite 1, luego --si.
"""

import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))
import excel_comun as xc  # noqa: E402

AUTO = PROYECTO / "automatizacion"
EXCEL = AUTO / "acciones_basicas.xlsx"
SERVICIOS = ["TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "SMP", "SMC", "SMO", "SMQ", "OS", "AC"]

EMP_RENOMBRE = ["TE-EMP-CM-102"]
EMP_CESE = ["TE-EMP-CM-104"]
EMP_FE = ["TE-EMP-CM-105", "TE-EMP-CM-197", "TE-EMP-CM-198", "TE-EMP-CM-201"]   # 4 serv., 6 serv., solo TF, solo móvil
PROMO = {   # (2 de monoproducto, 2 de empaquetado, 2 mixtas)
    "Fe de erratas": ["TP-TF-CM-101", "TP-SMP-CM-110", "TP-EMP-CM-150", "TP-EMP-CM-151", "TP-MIX-CM-148", "TP-MIX-CM-149"],
    "Renovación": ["TP-IF-CM-105", "TP-TVP-CM-107", "TP-EMP-CM-152", "TP-EMP-CM-153", "TP-MIX-CM-148", "TP-MIX-CM-149"],
    "Renombre": ["TP-OS-CM-114", "TP-AC-CM-115", "TP-EMP-CM-154", "TP-EMP-CM-155", "TP-MIX-CM-148", "TP-MIX-CM-149"],
}
LYC_FE = [f"TL-LYC-CM-{n}" for n in (101, 102, 103, 104, 105, 106,       # un servicio
                                     107, 120, 127, 140, 142, 148)]      # 3, 3, 4, 4, 5 y 6 servicios


def bases_usadas() -> set[str]:
    """Tarifas que son base de algún empaquetado o promoción (según sus Excel)."""
    usadas = set()
    for archivo, hoja, prefijo in (("empaquetados.xlsx", "Empaquetados", "Tarifa "),
                                   ("promociones.xlsx", "Promociones", "Tarifa base ")):
        ruta = AUTO / archivo
        if not ruta.exists():
            continue
        h = xc.Hoja(xc.abrir_libro(ruta, True)[hoja])
        nombre = "Nombre del paquete" if archivo.startswith("emp") else "Nombre de la promoción"
        for _, f in h.filas(nombre):
            usadas |= {xc.texto(v) for t, v in f.items()
                       if t.startswith(prefijo) and t[len(prefijo):].isdigit() and xc.texto(v)}
    return usadas


def main() -> None:
    if xc.excel_esta_abierto(EXCEL):
        sys.exit(f"Cierra el Excel antes de generar: {EXCEL}")
    usadas = bases_usadas()
    registro = xc.abrir_libro(AUTO / "registro_tarifas_establecidas.xlsx", True)
    acciones, avisos = [], []
    for codigo in SERVICIOS:
        nombres = [xc.texto(f["Nombre de la tarifa"]) for _, f in xc.Hoja(registro[codigo]).filas("Nombre de la tarifa")]
        ocupadas = {nombres[0]}                                   # la 1.ª queda sin acción
        for posicion, accion in ((1, "Renombre"), (2, "Cese"), (3, "Fe de erratas")):
            elegida = nombres[posicion] if posicion < len(nombres) else None
            if accion != "Fe de erratas" and (elegida in usadas or elegida in ocupadas):
                libre = next((n for n in nombres[4:] if n not in usadas and n not in ocupadas), None)
                avisos.append(f"{codigo}: {accion.lower()} → {libre} en vez de {elegida} (es base de un paquete o promoción)")
                elegida = libre
            if elegida:
                ocupadas.add(elegida)
                acciones.append((elegida, accion))
            else:
                avisos.append(f"{codigo}: no hay tarifa libre para {accion.lower()}")
    acciones += [(n, "Renombre") for n in EMP_RENOMBRE] + [(n, "Cese") for n in EMP_CESE] + [(n, "Fe de erratas") for n in EMP_FE]
    for accion, nombres in PROMO.items():
        acciones += [(n, accion) for n in nombres]
    acciones += [(n, "Fe de erratas") for n in LYC_FE]

    libro = xc.abrir_libro(EXCEL, True)
    hoja = libro["Acciones"]
    col = {xc.texto(c.value): c.column for c in hoja[xc.FILA_CABECERA] if c.value}
    ultima = max(col.values())
    for r in range(xc.FILA_CABECERA + 1, xc.FILA_CABECERA + 1 + len(acciones) + 20):
        for c in range(1, ultima + 1):
            hoja.cell(r, c).value = None
    for k, (nombre, accion) in enumerate(acciones):
        r = xc.FILA_CABECERA + 1 + k
        for titulo, valor in (("N°", k + 1), ("Ejecutar", "Sí"), ("Nombre de la tarifa", nombre), ("Acción", accion)):
            hoja.cell(r, col[titulo]).value = valor
    libro.save(EXCEL)
    for a in avisos:
        print("Aviso: " + a)
    from collections import Counter
    print(f"Listo: {len(acciones)} acciones en {EXCEL.name}: "
          + ", ".join(f"{v} {k.lower()}" for k, v in Counter(a for _, a in acciones).items()))


if __name__ == "__main__":
    main()
