"""
Corre la PRUEBA FUNCIONAL de un día completo, módulo tras módulo, sin tener que lanzar cada
comando. Para cada módulo: primero 1 fila (--limite 1) y, si esa fila dio lo esperado, el resto.
Después de cada módulo revisa los resultados (como verificar_funcional.py) y SE DETIENE si alguna
fila no dio lo esperado: así no sigue con módulos que dependen de él (p. ej. promociones de un
empaquetado que falló). Al volver a correr, continúa donde quedó (lo hecho no se repite).

  Día 1: registro → empaquetados → licitaciones → promociones → verificación
  Día 2: acciones → modificaciones → verificación

Uso (desde la carpeta del proyecto, Excel cerrados, Edge abierto y sin minimizar):
    python prueba_funcional.py --dia 1 --serie FB      (genera los Excel de la serie FB y corre el día 1)
    python prueba_funcional.py --dia 1                 (corre el día 1 con los Excel que ya están)
    python prueba_funcional.py --dia 2
    python prueba_funcional.py --dia 1 --desde promociones   (empieza en ese módulo)
    python prueba_funcional.py --dia 1 --serie P4 --carpeta Prueba4   (serie nueva en su carpeta)
    python prueba_funcional.py --dia 2 --carpeta Prueba4
"""

import argparse
import sys
import time

import excel_comun as xc
from verificar_funcional import CARPETA, evaluar

MODULOS = {
    1: [("registro", "registro_excel.py", "registro_funcional.xlsx", ["--definitivo-en-bloque"]),
        ("empaquetados", "empaquetados_excel.py", "empaquetados_funcional.xlsx", ["--definitivo-en-bloque"]),
        ("licitaciones", "licitaciones_excel.py", "licitaciones_funcional.xlsx", ["--definitivo-en-bloque"]),
        ("promociones", "promociones_excel.py", "promociones_funcional.xlsx", ["--definitivo-en-bloque"])],
    2: [("acciones", "acciones_excel.py", "acciones_funcional.xlsx", []),
        ("modificaciones", "modificaciones_excel.py", "modificaciones_funcional.xlsx", [])],
}


def leer_argumentos():
    p = argparse.ArgumentParser(description="Prueba funcional de un día, módulo tras módulo.")
    p.add_argument("--dia", type=int, choices=(1, 2), required=True)
    p.add_argument("--serie", help="(día 1) genera antes los Excel con esta serie nueva (p. ej. FB)")
    p.add_argument("--carpeta", help="subcarpeta de pruebas_funcionales con los Excel (p. ej. Prueba4)")
    p.add_argument("--desde", help="empieza en este módulo (registro, empaquetados, …)")
    p.add_argument("--solo", nargs="+", help="corre solo estos módulos (p. ej. --solo registro)")
    p.add_argument("--sin-prueba-de-1", action="store_true", help="no corre antes la fila de prueba")
    return p.parse_args()


DIR = CARPETA


def fallas(archivo: str, solo_errores: bool = False) -> list[dict]:
    """Filas que no dieron lo esperado. solo_errores=True (tras la prueba de 1 fila): solo las que
    terminaron en error; una «Registrada» a la espera de su registro definitivo no cuenta."""
    malas = [f for f in evaluar(DIR / archivo) if f["veredicto"] == "FALLA"]
    if solo_errores:
        malas = [f for f in malas if f["estado"].startswith(("Error", "No aplica"))]
    return malas


def detener(modulo: str, malas: list[dict]) -> None:
    print("\n" + "!" * 70)
    print(f"SE DETIENE en «{modulo}»: {len(malas)} fila(s) no dieron lo esperado:")
    for f in malas:
        print(f"  ✖ {f['nombre']}: esperado «{f['esperado']}», quedó «{f['estado']}». {f['detalle']}")
    print("Revisa el motivo (columna «Detalle» del Excel). Cuando esté resuelto, vuelve a correr el mismo")
    print("comando: lo que ya salió bien no se repite.")
    print(f"Tiempo hasta aquí: {xc.tiempo_total(0)}")
    print("!" * 70)
    sys.exit(1)


def main() -> None:
    global DIR
    args = leer_argumentos()
    if args.carpeta:
        DIR = CARPETA / args.carpeta
    if args.serie and (DIR / "registro_funcional.xlsx").exists() and args.carpeta:
        # Volver a lanzar el mismo comando tras un corte: NO se regeneran los Excel (guardan el
        # avance); se continúa con lo pendiente.
        print(f"La carpeta {DIR.name} ya tiene sus Excel: se continúa con lo pendiente (no se regeneran).")
    elif args.serie:
        if args.dia != 1:
            sys.exit("--serie solo se usa con --dia 1 (genera los Excel de ambos días).")
        print(f"Generando los Excel de la serie {args.serie.upper()}…")
        xc.correr([str(CARPETA / "generar_pruebas_funcionales.py"), "--serie", args.serie]
                  + (["--carpeta", args.carpeta] if args.carpeta else []))
    modulos = MODULOS[args.dia]
    if args.desde:
        nombres = [m[0] for m in modulos]
        if args.desde not in nombres:
            sys.exit(f"--desde: uno de {', '.join(nombres)}")
        modulos = modulos[nombres.index(args.desde):]
    if args.solo:
        nombres = [m[0] for m in modulos]
        otros = [n for n in args.solo if n not in nombres]
        if otros:
            sys.exit(f"--solo: uno o más de {', '.join(nombres)}")
        modulos = [m for m in modulos if m[0] in args.solo]
    for archivo in {m[2] for m in modulos}:
        ruta = DIR / archivo
        if not ruta.exists():
            sys.exit(f"Falta {ruta}. Genera los Excel: python prueba_funcional.py --dia 1 --serie XX")
        if xc.excel_esta_abierto(ruta):
            sys.exit(f"El Excel está abierto: ciérralo y vuelve a correr.\n  {ruta}")

    for numero, (modulo, script, archivo, extra) in enumerate(modulos, start=1):
        excel = ["--excel", str(DIR / archivo)]
        print("\n" + "#" * 70)
        print(f"# DÍA {args.dia} · {numero}/{len(modulos)}: {modulo.upper()}")
        print("#" * 70)
        desde = time.monotonic()
        if not args.sin_prueba_de_1:
            print(f"\n--- {modulo}: prueba con 1 fila")
            # Errores que ya estaban (p. ej. de una corrida cortada por sesión vencida): se
            # reintentan en la corrida completa; tras la prueba de 1 solo cuentan los NUEVOS.
            previas = {(f["nombre"], f["estado"]) for f in fallas(archivo)}
            xc.correr([script, *excel, "--limite", "1", "--si"])
            malas = [f for f in fallas(archivo, solo_errores=True) if (f["nombre"], f["estado"]) not in previas]
            if malas:
                detener(modulo, malas)
        print(f"\n--- {modulo}: todas las filas")
        xc.correr([script, *excel, "--si", *extra])
        malas = fallas(archivo)
        if malas:
            detener(modulo, malas)
        minutos = int(time.monotonic() - desde)
        print(f"\n✔ {modulo}: todo lo esperado ({minutos // 60} min {minutos % 60} s)")

    print("\n" + "#" * 70)
    print(f"# DÍA {args.dia} TERMINADO · verificación de todos los Excel")
    print("#" * 70)
    xc.correr(["verificar_funcional.py"] + (["--carpeta", args.carpeta] if args.carpeta else []))
    print(f"\nTiempo total del día {args.dia}: {xc.tiempo_total()}")


if __name__ == "__main__":
    main()
