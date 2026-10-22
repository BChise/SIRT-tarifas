"""
Llena automatizacion/modificaciones.xlsx con la PROPUESTA de rectificaciones y cambios de prueba:

Monoproductos (cada uno de los 15 servicios), por posición en su hoja del Excel de registro:
  5.ª rectificación («Todos los campos») · 6.ª cambio BENEFICIOSO («Todos los campos» + agregar
  fila) · 7.ª cambio PERJUDICIAL («Todos los campos» + ilimitado → limitado)                  = 45
  Si esa posición ya tiene otra acción (acciones_basicas.xlsx) o es base de un empaquetado /
  promoción, se usa la siguiente libre. NUNCA se usan bases: modificar una tarifa que forma parte de
  un paquete o promoción puede tener consecuencias en ellos.
Empaquetados: 2 rectificaciones, 2 cambios beneficiosos, 2 cambios perjudiciales           =  6
Licitaciones (solo rectificación): 1 de un servicio y 1 de varios                           =  2
Ejemplos de los otros modos: 1 rectificación «Solo la lista» y 1 cambio «Detalle en hoja Campos» = 2

Usa tarifas que NO sean la 1.ª de su servicio (queda tal como se registró), ni base de un
empaquetado o promoción, ni tengan otra acción en acciones_basicas.xlsx. Cada tarifa, una sola fila
(SIRT admite un solo cambio pendiente por tarifa).

Uso (con modificaciones.xlsx CERRADO):
    python automatizacion\\generar_plantilla_modificaciones.py
    python automatizacion\\generar_modificaciones_prueba.py
"""

import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO))
import excel_comun as xc  # noqa: E402

AUTO = PROYECTO / "automatizacion"
EXCEL = AUTO / "modificaciones.xlsx"
SERVICIOS = ["TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "SMP", "SMC", "SMO", "SMQ", "OS", "AC"]
EMP_CANDIDATOS = [f"TE-EMP-CM-{n}" for n in range(106, 197)]
LYC = [("TL-LYC-CM-102", "Rectificación"), ("TL-LYC-CM-130", "Rectificación")]


def ocupadas() -> tuple[set[str], set[str]]:
    """(bases de empaquetados / promociones, tarifas con acción en acciones_basicas.xlsx)."""
    bases, con_accion = set(), set()
    for archivo, hoja, nombre, prefijo in (("empaquetados.xlsx", "Empaquetados", "Nombre del paquete", "Tarifa "),
                                           ("promociones.xlsx", "Promociones", "Nombre de la promoción", "Tarifa base "),
                                           ("acciones_basicas.xlsx", "Acciones", "Nombre de la tarifa", None)):
        ruta = AUTO / archivo
        if not ruta.exists():
            continue
        for _, f in xc.Hoja(xc.abrir_libro(ruta, True)[hoja]).filas(nombre):
            if prefijo is None:
                con_accion.add(xc.texto(f[nombre]))
            else:
                bases |= {xc.texto(v) for t, v in f.items()
                          if t.startswith(prefijo) and t[len(prefijo):].isdigit() and xc.texto(v)}
    return bases, con_accion


def main() -> None:
    if xc.excel_esta_abierto(EXCEL):
        sys.exit(f"Cierra el Excel antes de generar: {EXCEL}")
    bases, con_accion = ocupadas()
    usadas = bases | con_accion
    registro = xc.abrir_libro(AUTO / "registro_tarifas_establecidas.xlsx", True)
    filas, avisos = [], []
    TODO = "Todos los campos"
    for codigo in SERVICIOS:
        nombres = [xc.texto(f["Nombre de la tarifa"]) for _, f in xc.Hoja(registro[codigo]).filas("Nombre de la tarifa")]
        planes = [("Rectificación", "", TODO, "", "No", "No"),
                  ("Cambio", "Beneficioso", TODO, "", "Sí", "No"),
                  ("Cambio", "Perjudicial", TODO, "", "No", "Sí")]
        if codigo == "TF":
            planes.append(("Rectificación", "", "Solo la lista", "Otros datos generales; Valor del servicio", "No", "No"))
        if codigo == "SMQ":
            planes.append(("Cambio", "", "Detalle en hoja Campos", "", "No", "No"))
        elegidas = set()
        for k, plan in enumerate(planes):
            preferida = nombres[4 + k] if 4 + k < len(nombres) else None        # 5.ª, 6.ª, 7.ª…
            libres = [n for n in nombres[4:] if n not in usadas and n not in elegidas]
            if preferida and preferida not in usadas and preferida not in elegidas:
                nombre = preferida
            elif libres:
                nombre = libres[0]
            else:
                avisos.append(f"{codigo}: no hay tarifa libre para {plan[0].lower()} {plan[1].lower()} "
                              "(agrega tarifas al Excel de registro)")
                continue
            if nombre != preferida:
                avisos.append(f"{codigo}: {plan[0].lower()} {plan[1].lower()} → {nombre} (la {5 + k}.ª está ocupada)")
            elegidas.add(nombre)
            filas.append((nombre, *plan))
    emp = [e for e in EMP_CANDIDATOS if e not in usadas][:6]
    planes_emp = [("Rectificación", "", TODO, "", "No", "No")] * 2 + [("Cambio", "Beneficioso", TODO, "", "No", "No")] * 2 \
        + [("Cambio", "Perjudicial", TODO, "", "No", "No")] * 2
    for nombre, plan in zip(emp, planes_emp):
        filas.append((nombre, *plan))
    for nombre, accion in LYC:
        filas.append((nombre, accion, "", "Por defecto", "", "No", "No"))

    libro = xc.abrir_libro(EXCEL, True)
    hoja = libro["Modificaciones"]
    col = {xc.texto(c.value): c.column for c in hoja[xc.FILA_CABECERA] if c.value}
    if "Cambio: sentido (Beneficioso / Perjudicial)" not in col:
        sys.exit("Regenera la plantilla primero: python automatizacion\\generar_plantilla_modificaciones.py")
    ultima = max(col.values())
    for r in range(xc.FILA_CABECERA + 1, xc.FILA_CABECERA + 1 + len(filas) + 20):
        for c in range(1, ultima + 1):
            hoja.cell(r, c).value = None
    for k, (nombre, accion, sentido, campos, lista, agregar, ilimitado) in enumerate(filas):
        r = xc.FILA_CABECERA + 1 + k
        valores = {"N°": k + 1, "Ejecutar": "Sí", "Nombre de la tarifa": nombre, "Acción": accion,
                   "Cambio: sentido (Beneficioso / Perjudicial)": sentido, "Campos a modificar": campos,
                   "Lista de campos (separados por ;)": lista, "Agregar una fila a cada tabla": agregar,
                   "Pasar «ilimitado» a limitado": ilimitado}
        for titulo, valor in valores.items():
            hoja.cell(r, col[titulo]).value = valor or None
    # Hoja Campos: el ejemplo de detalle apunta a la tarifa elegida
    hc = libro["Campos"]
    for r in range(6, hc.max_row + 1):
        for c in range(1, 5):
            hc.cell(r, c).value = None
    ejemplo = next((f[0] for f in filas if f[3] == "Detalle en hoja Campos"), None)
    if ejemplo:
        for i, (campo, valor, delta) in enumerate([
                ("Datos tope descarga", "", "5"), ("Datos para tethering", "3", ""), ("Valor del servicio", "", "-2"),
                ("Minutos móvil on net › Cantidad de minutos", "300", ""),
                ("Tecnologías y velocidades › 4G › Descarga máx. contratada", "", "5"),
                ("Tecnologías y velocidades › 5G", "", "-1")], start=6):
            for j, v in enumerate((ejemplo, campo, valor, delta), start=1):
                hc.cell(i, j).value = v or None
    libro.save(EXCEL)
    for a in avisos:
        print("Aviso: " + a)
    from collections import Counter
    print(f"Listo: {len(filas)} filas en {EXCEL.name}: "
          + ", ".join(f"{v} {k}" for k, v in Counter(f"{f[1].lower()}{' ' + f[2].lower() if f[2] else ''}" for f in filas).items()))


if __name__ == "__main__":
    main()
