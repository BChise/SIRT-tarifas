import csv
import re
from pathlib import Path


ARCHIVO_ENTRADA = Path("listado_tarifas_lyc.csv")
ARCHIVO_OBJETIVO = Path("tarifas_ejecucion_511.csv")
ARCHIVO_OTRAS = Path("tarifas_otros_lotes.csv")

IDENTIFICADOR_LOTE = "20260909-1503"

PATRON_LOTE = re.compile(
    rf"^LyC-(0[1-9])S-{re.escape(IDENTIFICADOR_LOTE)}-(\d{{3}})(?:_REN|_renombre)?$",
    re.IGNORECASE,
)


def leer_tarifas() -> list[dict[str, str]]:
    with ARCHIVO_ENTRADA.open("r", newline="", encoding="utf-8-sig") as archivo:
        return list(csv.DictReader(archivo, delimiter=";"))


def guardar_csv(ruta: Path, filas: list[dict[str, str]], campos: list[str]) -> None:
    with ruta.open("w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=campos, delimiter=";")
        escritor.writeheader()
        escritor.writerows(filas)


def ejecutar() -> None:
    tarifas = leer_tarifas()
    tarifas_objetivo = []
    otras_tarifas = []

    for tarifa in tarifas:
        nombre = tarifa["nombre"].strip()
        coincidencia = PATRON_LOTE.fullmatch(nombre)

        if coincidencia:
            tarifa["cantidad_servicios"] = coincidencia.group(1)
            tarifa["correlativo_lote"] = coincidencia.group(2)
            tarifas_objetivo.append(tarifa)
        else:
            otras_tarifas.append(tarifa)

    correlativos = {
        int(tarifa["correlativo_lote"])
        for tarifa in tarifas_objetivo
    }

    correlativos_esperados = set(range(1, 512))
    faltantes = sorted(correlativos_esperados - correlativos)
    adicionales = sorted(correlativos - correlativos_esperados)

    campos_objetivo = list(tarifas[0].keys()) + [
        "cantidad_servicios",
        "correlativo_lote",
    ]

    campos_objetivo = list(dict.fromkeys(campos_objetivo))
    campos_originales = list(tarifas[0].keys())

    guardar_csv(ARCHIVO_OBJETIVO, tarifas_objetivo, campos_objetivo)
    guardar_csv(ARCHIVO_OTRAS, otras_tarifas, campos_originales)

    print("\nVALIDACIÓN DEL LOTE")
    print("-------------------")
    print(f"Identificador: {IDENTIFICADOR_LOTE}")
    print(f"Tarifas encontradas del lote: {len(tarifas_objetivo)}")
    print(f"Correlativos únicos: {len(correlativos)}")
    print(f"Correlativos faltantes: {len(faltantes)}")
    print(f"Correlativos fuera de 001-511: {len(adicionales)}")
    print(f"Otras tarifas excluidas: {len(otras_tarifas)}")

    if faltantes:
        print(f"Faltantes: {faltantes}")

    if adicionales:
        print(f"Adicionales: {adicionales}")

    print(f"\nLote objetivo: {ARCHIVO_OBJETIVO.resolve()}")
    print(f"Otros lotes: {ARCHIVO_OTRAS.resolve()}")


if __name__ == "__main__":
    ejecutar()