"""
Catálogo para la app web: servicios, campos de cada servicio por acción y tarifas
conocidas (sacadas de los CSV de resultados). No abre el navegador ni toca SIRT.

Fuentes:
- servicios/*.py ............ código (TF, IF...) y nombre de cada servicio.
- grabaciones/<servicio>/paso2.json  campos del paso 2 del REGISTRO (GET de SIRT).
- CAMBIO_SIN ................ campos que el paso 2 de "Cambio de Tarifa" no muestra
                              (verificacion_cambio.xlsx, 16/10).
- resultados/**.csv ......... nombres de tarifas registradas, paquetes y promociones.
"""

import csv
import importlib
import json
import re
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
RESULTADOS = PROYECTO / "resultados"
GRABACIONES = PROYECTO / "grabaciones"

# tipoDato del JSON de SIRT → qué es para la app.
TIPOS = {
    "Texto área": "texto",
    "Numérica (Decimales)": "numero",
    "Numérica (enteros)": "numero",
    "Condicional triple": "numero",            # minutos / SMS: cantidad + tarifa
    "Categórica Multiple": "lista",
    "Listado multiple tecnologías": "lista",
    "Listado multiple con incremento en promoción": "lista",
    "Categorica Excluyente": "lista_simple",
    "Bloque de tarifas por paises": "tabla",
    "Tabla de velocidades de descarga y carga": "tabla",
    "Bolsa de gigas para plataformas": "tabla",
    "Carga y descarga": "archivo",
    "Texto": "texto_corto",                    # input de una línea (p. ej. Tecnología de transmisión)
}
# Lo que todavía no se puede cambiar en una acción (y por qué).
NO_SOPORTADO = {
    "lista_simple": "lista de una sola opción: aún no se cambia",
    "tabla": "tabla: falta grabar cómo se edita una fila",
    "archivo": "archivo Excel: aún no se cambia",
    "texto_corto": "texto de una línea: aún no se cambia",
}
# Qué tipos admite cada acción.
ACCION_TIPOS = {
    "fe_erratas": {"texto"},
    "rectificar": {"texto", "numero", "lista"},
    "cambio": {"texto", "numero", "lista"},
}
# Campos que NO aparecen en el paso 2 de "Cambio de Tarifa" (verificacion_cambio.xlsx).
CAMBIO_SIN = {
    "TF": ["Periodicidad", "Tipo de acceso", "Tecnologia"],
    "TFP": ["Duración", "Tipo de acceso", "Tecnologia"],
    "OTF": ["Periodicidad", "Duración", "Tipo de línea"],
    "PTF": ["Duración", "Tipo de línea"],
    "IF": ["Periodicidad", "Tipo de Acceso (IF)", "Tecnología (IF)", "Internet Ilimitado"],
    "PIF": ["Duración", "Tecnología (IF)"],
    "TVP": ["Periodicidad", "Tipo de Acceso (TV)", "Tecnologia"],
    "TVPR": ["Duración", "Tipo de Acceso (TV)", "Tecnologia"],
    "PTV": ["Duración", "Tipo de Contratación"],
    "SMP": ["Periodicidad", "Tipo de acceso", "Componentes del plan", "Internet Ilimitado"],
    "SMC": ["Periodicidad", "Tipo de acceso", "Componentes del plan", "Internet Ilimitado"],
    "SMO": ["Periodicidad", "Duración", "Tipo de línea móvil", "Tipo de acceso",
            "Componentes del plan", "Internet Ilimitado"],
    "SMQ": ["Duración", "Tipo de línea móvil", "Componentes del plan", "Internet Ilimitado"],
    "OS": ["Tipo de moneda", "Periodicidad"],
    "AC": ["Tipo de moneda", "Periodicidad", "Tecnología de transmisión"],
}
CON_INSTALACION = {"TFP", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV", "OS"}

# Servicios sin paso2.json del registro en grabaciones/: se arman desde uno parecido.
RESPALDO = {
    "TF": ("telefonia_fija_prepago", {"Duración": ("Periodicidad", "Categorica Excluyente")}),
    "IF": ("paquetes_internet_fijo", {"Duración": ("Periodicidad", "Categorica Excluyente")},
           [("Tipo de Acceso (IF)", "Categórica Multiple"), ("Internet Ilimitado", "Categorica Excluyente")]),
}


def tipo_de(tipo_dato: str) -> str:
    buscado = _sin_tildes(tipo_dato)
    for clave, valor in TIPOS.items():
        if _sin_tildes(clave) == buscado:
            return valor
    if "tabla" in buscado or "bolsa" in buscado or "paises" in buscado:
        return "tabla"
    return "otro"


def _sin_tildes(texto: str) -> str:
    import unicodedata
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"\s+", " ", texto).strip()


def servicios() -> list[dict]:
    """[{codigo, nombre, archivo}] en el orden de los archivos de servicios/."""
    lista = []
    for ruta in sorted((PROYECTO / "servicios").glob("*.py")):
        if ruta.stem.startswith("_"):
            continue
        texto = ruta.read_text(encoding="utf-8")
        codigo = re.search(r'^CODIGO\s*=\s*"([^"]+)"', texto, re.M)
        nombre = re.search(r'^SERVICIO\s*=\s*"([^"]+)"', texto, re.M)
        if codigo:
            lista.append({"codigo": codigo.group(1), "archivo": ruta.stem,
                          "nombre": nombre.group(1) if nombre else ruta.stem})
    orden = ["TF", "TFP", "OTF", "PTF", "IF", "PIF", "TVP", "TVPR", "PTV",
             "SMP", "SMC", "SMO", "SMQ", "OS", "AC"]
    return sorted(lista, key=lambda s: orden.index(s["codigo"]) if s["codigo"] in orden else 99)


def _apartados(archivo: str) -> list[tuple[str, str]]:
    ruta = GRABACIONES / archivo / "paso2.json"
    if not ruta.exists() or ruta.stat().st_size == 0:
        return []
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))["data"]
    except Exception:
        return []
    salida = []
    for seccion in datos.get("secciones") or []:
        for apartado in seccion.get("apartados") or []:
            salida.append((apartado["nombre"].strip(), (apartado.get("tipoDato") or {}).get("nombre", "")))
    return salida


def campos_paso_2(servicio: dict, por_archivo: dict) -> list[tuple[str, str]]:
    campos = _apartados(servicio["archivo"])
    if campos or servicio["codigo"] not in RESPALDO:
        return campos
    origen, cambios, *extra = RESPALDO[servicio["codigo"]]
    base = _apartados(origen)
    campos = [cambios.get(n, (n, t)) for n, t in base]
    for n, t in (extra[0] if extra else []):
        if n not in [c[0] for c in campos]:
            campos.insert(2, (n, t))
    return campos


def campos_por_accion(servicio: dict, accion: str) -> list[dict]:
    """Campos de un servicio para una acción: [{nombre, paso, tipo, ok, motivo}]."""
    por_archivo = {}
    filas = [{"nombre": "Otros datos generales", "paso": 1, "tipo": "texto"}]
    if servicio["codigo"] in CON_INSTALACION:
        filas.append({"nombre": "Tarifa de instalación", "paso": 1, "tipo": "numero"})
    sin = {_sin_tildes(n) for n in CAMBIO_SIN.get(servicio["codigo"], [])} if accion == "cambio" else set()
    for nombre, tipo_dato in campos_paso_2(servicio, por_archivo):
        if _sin_tildes(nombre) in sin:
            continue
        filas.append({"nombre": nombre, "paso": 2, "tipo": tipo_de(tipo_dato)})
    admitidos = ACCION_TIPOS.get(accion, set())
    for f in filas:
        f["ok"] = f["tipo"] in admitidos
        f["motivo"] = "" if f["ok"] else (
            "fe de erratas solo corrige textos" if accion == "fe_erratas" and f["tipo"] in ("numero", "lista")
            else NO_SOPORTADO.get(f["tipo"], "no se cambia en esta acción"))
    return filas


def catalogo_campos() -> dict:
    """{accion: {codigo: [campos]}} para fe_erratas, rectificar y cambio."""
    lista = servicios()
    return {accion: {s["codigo"]: campos_por_accion(s, accion) for s in lista}
            for accion in ACCION_TIPOS}


# --- Tarifas conocidas ----------------------------------------------------------

def _leer_csv(ruta: Path) -> list[dict]:
    try:
        with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
            return list(csv.DictReader(entrada, delimiter=";"))
    except Exception:
        return []


def tarifas_conocidas() -> list[dict]:
    """Tarifas con resultado EXITOSO en los CSV: [{nombre, codigo, origen, definitivo}].
    Aplica los renombres exitosos de acciones.csv (el nombre actual es el nuevo)."""
    vistas: dict[str, dict] = {}
    definitivas = set()
    for ruta in RESULTADOS.glob("registro_definitivo_*.csv"):
        definitivas |= {f.get("nombre") for f in _leer_csv(ruta) if f.get("resultado") == "EXITOSO"}
    for ruta in sorted(RESULTADOS.glob("*/resultados_*.csv")):
        carpeta = ruta.parent.name
        for fila in _leer_csv(ruta):
            nombre = fila.get("nombre") or ""
            if fila.get("resultado") in ("EXITOSO", "CONFIRMADA") or (
                    carpeta in ("EMP", "PRM") and "ERROR" not in (fila.get("resultado") or "")):
                if not nombre:
                    continue
                vistas[nombre] = {"nombre": nombre, "codigo": codigo_de(nombre) or carpeta,
                                  "origen": {"EMP": "empaquetado", "PRM": "promoción"}.get(carpeta, "establecida")}
    for fila in _leer_csv(RESULTADOS / "ACCIONES" / "acciones.csv"):
        if fila.get("accion") == "renombre" and fila.get("resultado") == "EXITOSO" and fila.get("nombre_nuevo"):
            anterior = vistas.pop(fila["nombre"], {"codigo": codigo_de(fila["nombre"]), "origen": "establecida"})
            vistas[fila["nombre_nuevo"]] = {**anterior, "nombre": fila["nombre_nuevo"]}
            if fila["nombre"] in definitivas:
                definitivas.add(fila["nombre_nuevo"])
    salida = []
    for t in vistas.values():
        t["definitivo"] = t["nombre"] in definitivas
        salida.append(t)
    return sorted(salida, key=lambda t: t["nombre"])


def codigo_de(nombre: str) -> str:
    """'TE-SMP-PROMO-001' → 'SMP'; 'TP-EMP-...' → 'EMP'; si no, ''."""
    m = re.match(r"^T[EP]-([A-Z]+)-", nombre or "")
    return m.group(1) if m else ""


def lotes_conocidos() -> list[str]:
    lotes = set()
    for ruta in RESULTADOS.glob("*/resultados_*.csv"):
        lotes.add(ruta.stem.replace("resultados_", ""))
    return sorted(lotes)
