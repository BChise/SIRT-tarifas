"""
Resume el JSON del GET del paso 2 de un servicio.

Uso:
    python inspeccionar_json.py telefonia_fija_prepago

Lee grabaciones/<servicio>/paso2.json y muestra:
- datos del servicio (nombre exacto, si registra Tarifa Instalación),
- cada campo del paso 2: tipo, obligatoriedad, unidad y opciones exactas,
- un borrador de DIMENSIONES con las listas, para copiar al servicio.
"""

import json
import sys
from pathlib import Path


TIPOS_CON_BLOQUE = {
    "Condicional triple": "completar_condicional_triple",
    "Bloque de tarifas por paises": "agregar_tarifa_paises",
}


def mostrar_apartado(apartado: dict, sangria: str = "  ") -> None:
    tipo_dato = (apartado.get("tipoDato") or {}).get("nombre", "")
    obligatorio = "obligatorio" if apartado.get("esObligatorio") else "opcional"
    unidad = f" [{apartado['unidadMedida']}]" if apartado.get("unidadMedida") else ""
    bloque = TIPOS_CON_BLOQUE.get(tipo_dato)
    nota = f"  → comun.{bloque}" if bloque else ""

    print(f"{sangria}- {apartado['nombre'].strip()!r}{unidad}: {tipo_dato}, {obligatorio}{nota}")

    opciones = [opcion.get("nombre") for opcion in apartado.get("lista") or []]
    if opciones:
        print(f"{sangria}    opciones: {opciones}")

    patron = (apartado.get("validarPatron") or {}).get("mensaje")
    if patron:
        print(f"{sangria}    validación: {patron}")

    if apartado.get("caracteres"):
        print(f"{sangria}    máx. caracteres: {apartado['caracteres']}")

    for sub in apartado.get("subApartados") or []:
        mostrar_apartado(sub, sangria + "    ")


def main() -> None:
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)

    ruta = Path(__file__).parent / "grabaciones" / sys.argv[1] / "paso2.json"
    if not ruta.exists() or ruta.stat().st_size == 0:
        print(f"No existe o está vacío: {ruta}")
        sys.exit(1)

    datos = json.loads(ruta.read_text(encoding="utf-8"))["data"]
    modalidad = datos["tarifa"]["modalidad"]

    print(f"Servicio (texto exacto): {modalidad['nombre']!r}")
    print(f"Registra Tarifa Instalación: {modalidad.get('registraTarifaInstalacion')}")

    listas = []

    for seccion in datos["secciones"]:
        print(f"\n## {seccion['seccion']}")
        for apartado in seccion["apartados"]:
            mostrar_apartado(apartado)
            opciones = [opcion.get("nombre") for opcion in apartado.get("lista") or []]
            if opciones:
                tipo_dato = (apartado.get("tipoDato") or {}).get("nombre", "").lower()
                listas.append((apartado["nombre"].strip(), opciones, "multiple" in tipo_dato))

    print("\n## Borrador de DIMENSIONES (revisar 'multiple' contra la pantalla)")
    print('    "alcance": {"opciones": ["Residencial", "Comercial"], "multiple": True},')
    for nombre, opciones, multiple in listas:
        clave = nombre.lower().replace(" de ", "_").replace(" ", "_")
        print(f'    "{clave}": {{"opciones": {opciones}, "multiple": {multiple}}},   # {nombre!r}')


if __name__ == "__main__":
    main()
