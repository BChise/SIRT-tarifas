"""
Llena automatizacion/registro_licitaciones.xlsx con las tarifas de prueba de Licitación y Contratos:
todas las combinaciones de los 6 servicios en grupos de 1, 3, 4, 5 y 6 (6 + 20 + 15 + 6 + 1 = 48).

Nombres TL-LYC-CM-101 en adelante. Reemplaza las filas de datos de la hoja «Licitaciones»
(conserva formato, listas desplegables y la columna calculada «Cantidad de servicios»).

Uso (desde la carpeta del proyecto, con el Excel CERRADO):
    python automatizacion\\generar_licitaciones_prueba.py
    python automatizacion\\generar_licitaciones_prueba.py --tamanos 1 2 3 4 5 6      (incluye los de 2: 63)
"""

import argparse
import sys
from itertools import combinations
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))
import excel_comun as xc  # noqa: E402

EXCEL = PROYECTO / "automatizacion" / "registro_licitaciones.xlsx"
SERVICIOS = [("Arrendamiento de circuitos", "AC"), ("Otros servicios", "OS"), ("Televisión de paga", "TVP"),
             ("Internet Fijo", "IF"), ("Servicio Móvil", "SM"), ("Telefonía Fija", "TF")]
AMBITOS = ["Público", "Privado"]
MONEDAS = ["Soles", "Dolares", "Euro"]
PRIMER_NUMERO = 101


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--tamanos", nargs="+", type=int, default=[1, 3, 4, 5, 6])
    args = p.parse_args()
    if xc.excel_esta_abierto(EXCEL):
        sys.exit(f"Cierra el Excel antes de generar: {EXCEL}")
    libro = xc.abrir_libro(EXCEL, True)
    hoja = libro["Licitaciones"]
    col = {xc.texto(c.value): c.column for c in hoja[xc.FILA_CABECERA] if c.value}
    calculada = col["Cantidad de servicios"]

    grupos = [g for n in sorted(args.tamanos) for g in combinations(SERVICIOS, n)]
    ultima = max(col.values())
    for r in range(xc.FILA_CABECERA + 1, xc.FILA_CABECERA + 1 + max(len(grupos), 50) + 10):
        for c in range(1, ultima + 1):
            if c != calculada:
                hoja.cell(r, c).value = None

    for k, grupo in enumerate(grupos):
        r = xc.FILA_CABECERA + 1 + k
        numero = PRIMER_NUMERO + k
        codigos = "+".join(c for _, c in grupo)
        incluidos = {s for s, _ in grupo}
        valores = {
            "N°": k + 1, "Ejecutar": "Sí", "Nombre de la tarifa": f"TL-LYC-CM-{numero:03d}",
            "Registro definitivo": "Sí", "Duración del contrato (días)": 365,
            "Identificador de contrato": f"CTR-CM-{numero:03d}", "Ámbito": AMBITOS[k % 2],
            "Valor de contrato": f"{1000 + 250 * len(grupo) + k:.2f}", "Moneda": MONEDAS[k % 3],
            "Documento de Empresa": f"20100000{numero:03d}",
            "Nombre de empresa": f"Empresa contratante de prueba {numero:03d} SAC",
            "Descripción del Servicio": f"Servicio licitado de prueba ({codigos})",
            "Atributos de la tarifa": f"Atributos de la tarifa de prueba {numero:03d}",
            "Ámbito geográfico": "Lima Metropolitana",
            "Tarifa instalación (solo con Otros servicios)": "10" if "Otros servicios" in incluidos else None,
        }
        for servicio, _ in SERVICIOS:
            dentro = servicio in incluidos
            valores[f"{servicio}: incluir"] = "Sí" if dentro else "No"
            valores[f"{servicio}: detalle del servicio"] = \
                f"Detalle de {servicio.lower()} para la tarifa {numero:03d}" if dentro else None
        for titulo, valor in valores.items():
            hoja.cell(r, col[titulo]).value = valor
    libro.save(EXCEL)
    por_tamano = {n: sum(1 for g in grupos if len(g) == n) for n in sorted(args.tamanos)}
    print(f"Listo: {len(grupos)} tarifas en {EXCEL.name} "
          + "(" + ", ".join(f"{v} de {n}" for n, v in por_tamano.items()) + ")"
          + f": TL-LYC-CM-{PRIMER_NUMERO:03d} a TL-LYC-CM-{PRIMER_NUMERO + len(grupos) - 1:03d}")


if __name__ == "__main__":
    main()
