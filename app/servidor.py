"""
App web local para usar la herramienta de pruebas de SIRT desde el navegador.

Arranque (desde la carpeta del proyecto, con el entorno virtual activado):
    python -m app.servidor            → abre http://127.0.0.1:8000

Cada botón lanza el script que ya existe (ejecutar.py, empaquetar.py, promocionar.py,
registro_definitivo.py, acciones.py) como un proceso aparte, muestra su salida en vivo
y lee los CSV de resultados. Solo corre una tarea a la vez (hay un solo Edge con SIRT).
"""

import csv
import itertools
import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app import catalogo, datos as datos_excel

PROYECTO = catalogo.PROYECTO
ESTATICOS = Path(__file__).resolve().parent / "static"
CDP_URL = "http://localhost:9222"
PYTHON = sys.executable

app = FastAPI(title="SIRT - Herramienta de pruebas")
app.mount("/static", StaticFiles(directory=ESTATICOS), name="static")


# --- Tareas (procesos en segundo plano) -------------------------------------------

class Tarea:
    def __init__(self, titulo: str, comandos: list[list[str]]):
        self.id = next(_ids)
        self.titulo = titulo
        self.comandos = comandos
        self.estado = "en cola"          # en cola | corriendo | terminada | con errores | cancelada
        self.lineas: list[str] = []
        self.inicio = datetime.now()
        self.fin: datetime | None = None
        self.proceso: subprocess.Popen | None = None
        self.cancelada = False
        self.fallos = 0

    def resumen(self) -> dict:
        return {"id": self.id, "titulo": self.titulo, "estado": self.estado,
                "inicio": self.inicio.strftime("%H:%M:%S"),
                "fin": self.fin.strftime("%H:%M:%S") if self.fin else "",
                "comandos": len(self.comandos), "lineas": len(self.lineas)}

    def escribir(self, texto: str) -> None:
        self.lineas.append(texto.rstrip("\n"))

    def correr(self) -> None:
        self.estado = "corriendo"
        entorno = {**os.environ, "PYTHONUNBUFFERED": "1", "PYTHONIOENCODING": "utf-8"}
        for numero, comando in enumerate(self.comandos, start=1):
            if self.cancelada:
                break
            visible = " ".join(c if " " not in c else f'"{c}"' for c in comando[1:])
            self.escribir(f"\n▶ [{numero}/{len(self.comandos)}] python {visible}")
            try:
                self.proceso = subprocess.Popen(
                    comando, cwd=PROYECTO, env=entorno, stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1)
                for linea in self.proceso.stdout:
                    self.escribir(linea)
                codigo = self.proceso.wait()
            except Exception as error:
                self.escribir(f"✖ No se pudo ejecutar: {error}")
                codigo = 1
            if codigo != 0 and not self.cancelada:
                self.fallos += 1
                self.escribir(f"✖ El comando terminó con código {codigo}")
        self.fin = datetime.now()
        self.estado = "cancelada" if self.cancelada else ("con errores" if self.fallos else "terminada")
        self.escribir(f"\n■ {self.estado.upper()} ({(self.fin - self.inicio).seconds} s)")

    def cancelar(self) -> None:
        self.cancelada = True
        if self.proceso and self.proceso.poll() is None:
            self.proceso.terminate()


_ids = itertools.count(1)
TAREAS: dict[int, Tarea] = {}
_candado = threading.Lock()


def tarea_activa() -> Tarea | None:
    return next((t for t in TAREAS.values() if t.estado in ("en cola", "corriendo")), None)


def edge_conectado() -> bool:
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=1.5):
            return True
    except Exception:
        return False


def lanzar(titulo: str, comandos: list[list[str]]) -> Tarea:
    if not edge_conectado():
        raise HTTPException(409, "Edge no está conectado (puerto 9222). Ábrelo con el perfil de SIRT, "
                                 "inicia sesión y vuelve a intentar.")
    with _candado:
        activa = tarea_activa()
        if activa:
            raise HTTPException(409, f"Ya hay una tarea en curso: «{activa.titulo}». Espera o cancélala.")
        tarea = Tarea(titulo, comandos)
        TAREAS[tarea.id] = tarea
    threading.Thread(target=tarea.correr, daemon=True).start()
    return tarea


# --- Construcción de comandos ------------------------------------------------------

def script(nombre: str, *argumentos) -> list[str]:
    return [PYTHON, "-u", nombre, *[str(a) for a in argumentos if a not in (None, "")]]


def limpiar_lista(valores: list[str]) -> list[str]:
    salida = []
    for v in valores or []:
        for parte in re.split(r"[\n,]+", v):
            parte = parte.strip()
            if parte and parte not in salida:
                salida.append(parte)
    return salida


def validar_lote(lote: str) -> str:
    lote = (lote or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,30}", lote):
        raise HTTPException(400, "El lote solo puede tener letras, números, '-' y '_' (máx. 30).")
    return lote


class Establecidas(BaseModel):
    servicios: list[str]
    lote: str
    limite: int = 1
    diagnostico: bool = True


class Definitivo(BaseModel):
    lote: str
    servicios: list[str] = []
    nombres: list[str] = []
    limite: int = 0
    ubigeos: str = ""
    diagnostico: bool = True


class Empaquetar(BaseModel):
    tarifas: list[str]
    lote: str
    nombre: str = ""
    valores: list[str] = []
    diagnostico: bool = True


class Promocionar(BaseModel):
    tarifas: list[str]
    lote: str
    una_por_tarifa: bool = False
    atributos: str = "probados"
    variacion_porcentual: str = "10"
    variacion_cantidad: str = "1"
    duracion: str = "30"
    dias_comercializacion: int = 60
    sin_valor_plan: bool = False
    sin_instalacion: bool = False
    diagnostico: bool = True


class Accion(BaseModel):
    accion: str
    tarifas: list[str]
    hasta_resumen: bool = True
    diagnostico: bool = True
    repetir: bool = False
    sin_validar: bool = False
    # renombre
    sufijo: str = "_REN"
    nuevo_nombre: str = ""
    # cese / cambio
    fecha: str = ""
    # rectificar / cambio
    delta: str = ""
    modo: str = "defecto"            # defecto | todos | elegir
    campos: list[str] = []
    texto: str = ""


@app.post("/api/tareas/establecidas")
def tarea_establecidas(datos: Establecidas):
    lote = validar_lote(datos.lote)
    por_codigo = {s["codigo"]: s for s in catalogo.servicios()}
    elegidos = [por_codigo[c] for c in datos.servicios if c in por_codigo]
    if not elegidos:
        raise HTTPException(400, "Elige al menos un servicio.")
    comandos = [script("ejecutar.py", s["archivo"], "--lote", lote, "--limite", max(1, datos.limite), "--si",
                       "--diagnostico" if datos.diagnostico else None) for s in elegidos]
    return lanzar(f"Registrar establecidas ({len(elegidos)} servicio(s), lote {lote})", comandos).resumen()


@app.post("/api/tareas/definitivo")
def tarea_definitivo(datos: Definitivo):
    lote = validar_lote(datos.lote)
    argumentos = ["--lote", lote, "--si"]
    for c in datos.servicios:
        argumentos += ["--servicio", c]
    for n in limpiar_lista(datos.nombres):
        argumentos += ["--nombre", n]
    if datos.limite:
        argumentos += ["--limite", datos.limite]
    if datos.ubigeos.strip():
        argumentos += ["--ubigeos", datos.ubigeos.strip()]
    if datos.diagnostico:
        argumentos.append("--diagnostico")
    return lanzar(f"Registro definitivo (lote {lote})", [script("registro_definitivo.py", *argumentos)]).resumen()


@app.post("/api/tareas/empaquetar")
def tarea_empaquetar(datos: Empaquetar):
    tarifas = limpiar_lista(datos.tarifas)
    if len(tarifas) < 2:
        raise HTTPException(400, "Un empaquetado necesita al menos 2 tarifas.")
    argumentos = ["--tarifas", *tarifas, "--lote", validar_lote(datos.lote)]
    if datos.nombre.strip():
        argumentos += ["--nombre", datos.nombre.strip()]
    valores = limpiar_lista(datos.valores)
    if valores:
        argumentos += ["--valores", *valores]
    if datos.diagnostico:
        argumentos.append("--diagnostico")
    return lanzar(f"Empaquetar {len(tarifas)} tarifas", [script("empaquetar.py", *argumentos)]).resumen()


@app.post("/api/tareas/promocionar")
def tarea_promocionar(datos: Promocionar):
    tarifas = limpiar_lista(datos.tarifas)
    if not tarifas:
        raise HTTPException(400, "Indica al menos una tarifa.")
    lote = validar_lote(datos.lote)
    comunes = ["--lote", lote, "--atributos", datos.atributos or "probados",
               "--variacion-porcentual", datos.variacion_porcentual,
               "--variacion-cantidad", datos.variacion_cantidad,
               "--duracion", datos.duracion, "--dias-comercializacion", datos.dias_comercializacion]
    if datos.sin_valor_plan:
        comunes.append("--sin-valor-plan")
    if datos.sin_instalacion:
        comunes.append("--sin-instalacion")
    if datos.diagnostico:
        comunes.append("--diagnostico")
    grupos = [[t] for t in tarifas] if datos.una_por_tarifa else [tarifas]
    comandos = [script("promocionar.py", "--tarifas", *g, *comunes) for g in grupos]
    return lanzar(f"Promocionar ({len(comandos)} promoción(es))", comandos).resumen()


ACCIONES = {"buscar", "renombre", "cese", "fe_erratas", "rectificar", "cambio"}


@app.post("/api/tareas/accion")
def tarea_accion(datos: Accion):
    if datos.accion not in ACCIONES:
        raise HTTPException(400, "Acción desconocida.")
    tarifas = limpiar_lista(datos.tarifas)
    if not tarifas:
        raise HTTPException(400, "Indica al menos una tarifa.")
    argumentos = [datos.accion, "--tarifas", *tarifas]
    if datos.accion != "buscar":
        argumentos.append("--si")
    if datos.diagnostico:
        argumentos.append("--diagnostico")
    if datos.repetir:
        argumentos.append("--repetir")
    if datos.sin_validar and datos.accion in ("cese", "fe_erratas", "rectificar", "cambio"):
        argumentos.append("--sin-validar")
    if datos.accion == "renombre":
        if datos.nuevo_nombre.strip():
            if len(tarifas) != 1:
                raise HTTPException(400, "El nombre exacto solo sirve con una tarifa.")
            argumentos += ["--nuevo-nombre", datos.nuevo_nombre.strip()]
        else:
            argumentos += ["--sufijo", datos.sufijo or "_REN"]
    if datos.accion in ("cese", "cambio") and datos.fecha:
        argumentos += ["--fecha", datos.fecha]
    if datos.accion in ("fe_erratas", "rectificar", "cambio"):
        if datos.hasta_resumen:
            argumentos.append("--hasta-resumen")
        if datos.texto.strip():
            argumentos += ["--texto", datos.texto.strip()]
        if datos.modo == "todos" and datos.accion != "fe_erratas":
            argumentos.append("--todo")
        elif datos.modo == "elegir":
            campos = [c for c in datos.campos if c.strip()]
            if not campos:
                raise HTTPException(400, "Marca al menos un campo o elige otro modo.")
            argumentos += ["--campos", ";".join(campos)]
    if datos.accion in ("rectificar", "cambio") and datos.delta.strip():
        if not re.fullmatch(r"-?\d+(\.\d+)?", datos.delta.strip()):
            raise HTTPException(400, "El delta debe ser un número, p. ej. 1 o -1.")
        argumentos += [f"--delta={datos.delta.strip()}"]
    nombres = {"buscar": "Buscar", "renombre": "Renombrar", "cese": "Cese", "fe_erratas": "Fe de erratas",
               "rectificar": "Rectificar", "cambio": "Cambio de tarifa"}
    sufijo = " (hasta el resumen)" if datos.hasta_resumen and datos.accion in ("fe_erratas", "rectificar", "cambio") else ""
    return lanzar(f"{nombres[datos.accion]}: {len(tarifas)} tarifa(s){sufijo}",
                  [script("acciones.py", *argumentos)]).resumen()


# --- Consultas ---------------------------------------------------------------------

@app.get("/")
def inicio():
    return FileResponse(ESTATICOS / "demo.html")


@app.get("/avanzado")
def avanzado():
    return FileResponse(ESTATICOS / "index.html")


# --- Versión simple: tarifas del Excel (datos/pruebas.xlsx) -------------------------

class Seleccion(BaseModel):
    nombres: list[str]
    diagnostico: bool = True


def filas_elegidas(nombres: list[str]) -> list[dict]:
    try:
        filas = datos_excel.con_estado(datos_excel.leer_filas())
    except FileNotFoundError as error:
        raise HTTPException(404, str(error))
    elegidas = [f for f in filas if f["nombre"] in set(nombres)]
    if not elegidas:
        raise HTTPException(400, "Elige al menos una tarifa.")
    malas = [f"{f['nombre']}: {'; '.join(f['problemas'])}" for f in elegidas if f["problemas"]]
    if malas:
        raise HTTPException(400, "Revisa el Excel — " + " | ".join(malas))
    return elegidas


@app.get("/api/demo/tarifas")
def demo_tarifas():
    try:
        todo = datos_excel.todo_con_estado()
    except FileNotFoundError as error:
        raise HTTPException(404, str(error))
    return {"archivo": str(datos_excel.ARCHIVO.relative_to(PROYECTO)), **todo}


def comandos_definitivo(filas: list[dict], diagnostico: bool) -> list[list[str]]:
    """registro_definitivo.py por lote, para las filas que lo piden y aún no lo tienen."""
    por_lote: dict[str, list[str]] = {}
    for f in filas:
        if f["quiere_definitivo"] and f.get("estado") != "Definitivo":
            por_lote.setdefault(f["lote"], []).append(f["nombre"])
    comandos = []
    for lote, nombres in por_lote.items():
        argumentos = ["--lote", lote, "--si"]
        for n in nombres:
            argumentos += ["--nombre", n]
        if diagnostico:
            argumentos.append("--diagnostico")
        comandos.append(script("registro_definitivo.py", *argumentos))
    return comandos


def _elegidas(lista: list[dict], nombres: list[str]) -> list[dict]:
    elegidas = [f for f in lista if f["nombre"] in set(nombres)]
    if not elegidas:
        raise HTTPException(400, "Elige al menos una.")
    return elegidas


@app.post("/api/demo/empaquetar")
def demo_empaquetar(seleccion: Seleccion):
    todo = datos_excel.todo_con_estado()
    paquetes = _elegidas(todo["paquetes"], seleccion.nombres)
    definitivas = {f["nombre"] for f in todo["filas"] if f["estado"] == "Definitivo"}
    faltan = sorted({t for p in paquetes if p["estado"] in ("Pendiente", "Error")
                     for t in p["tarifas"] if t not in definitivas})
    if faltan:
        raise HTTPException(400, "Estas tarifas base aún no tienen registro definitivo: " + ", ".join(faltan)
                            + ". Regístralas primero con «Registrar tarifas».")
    comandos = [script("empaquetar.py", "--tarifas", *p["tarifas"], "--lote", p["lote"], "--nombre", p["nombre"],
                       "--diagnostico" if seleccion.diagnostico else None)
                for p in paquetes if p["estado"] in ("Pendiente", "Error")]
    comandos += comandos_definitivo(paquetes, seleccion.diagnostico)
    if not comandos:
        raise HTTPException(400, "Los paquetes elegidos ya están en definitivo.")
    return lanzar(f"Registrar {len(paquetes)} empaquetado(s) (+ registro definitivo)", comandos).resumen()


@app.post("/api/demo/promocionar")
def demo_promocionar(seleccion: Seleccion):
    todo = datos_excel.todo_con_estado()
    promos = _elegidas(todo["promos"], seleccion.nombres)
    definitivas = {f["nombre"] for f in todo["filas"] + todo["paquetes"] if f["estado"] == "Definitivo"}
    faltan = sorted({t for p in promos if p["estado"] in ("Pendiente", "Error")
                     for t in p["tarifas"] if t not in definitivas})
    if faltan:
        raise HTTPException(400, "Estas tarifas base aún no tienen registro definitivo: " + ", ".join(faltan))
    comandos = [script("promocionar.py", "--tarifas", *p["tarifas"], "--lote", p["lote"], "--nombre", p["nombre"],
                       "--atributos", p["atributos_promo"], "--variacion-porcentual", p["variacion_porcentual"],
                       "--variacion-cantidad", p["variacion_cantidad"], "--duracion", p["duracion"],
                       "--diagnostico" if seleccion.diagnostico else None)
                for p in promos if p["estado"] in ("Pendiente", "Error")]
    comandos += comandos_definitivo(promos, seleccion.diagnostico)   # promociones: sin disponibilidad
    if not comandos:
        raise HTTPException(400, "Las promociones elegidas ya están en definitivo.")
    return lanzar(f"Registrar {len(promos)} promoción(es) (+ registro definitivo)", comandos).resumen()


class SeleccionAcciones(Seleccion):
    grupo: str                         # cese_renombre | modificar


def _parametro(texto: str, clave: str) -> str:
    m = re.search(rf"{clave}\s*=\s*([^\s;]+)", texto or "")
    return m.group(1) if m else ""


@app.post("/api/demo/acciones")
def demo_acciones(seleccion: SeleccionAcciones):
    todo = datos_excel.todo_con_estado()
    admitidas = {"cese_renombre": {"renombre", "cese"},
                 "modificar": {"fe_erratas", "rectificar", "cambio"}}.get(seleccion.grupo)
    if not admitidas:
        raise HTTPException(400, "Grupo de acciones desconocido.")
    filas = [f for f in _elegidas(todo["filas"] + todo["paquetes"], seleccion.nombres)
             if datos_excel.ACCION_SCRIPT.get(f["accion"]) in admitidas and f.get("estado_accion") != "Aplicada"]
    if not filas:
        raise HTTPException(400, "No hay acciones pendientes entre las elegidas.")
    grupos: dict[tuple, list[str]] = {}
    for f in filas:
        accion = datos_excel.ACCION_SCRIPT[f["accion"]]
        if accion == "renombre":
            clave = (accion, ("--sufijo", _parametro(f["parametros"], "sufijo") or "_REN"))
        elif accion == "cese":
            fecha = _parametro(f["parametros"], "fecha")
            clave = (accion, ("--fecha", fecha) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", fecha) else ())
        elif accion == "fe_erratas":
            clave = (accion, ())
        else:                                        # rectificar / cambio: todos los campos posibles
            delta = _parametro(f["parametros"], "delta") or ("1" if accion == "rectificar" else "-1")
            clave = (accion, (f"--delta={delta}", "--todo"))
        grupos.setdefault(clave, []).append(f["nombre"])
    comandos = []
    for (accion, extra), nombres in grupos.items():
        comandos.append(script("acciones.py", accion, "--tarifas", *nombres, *extra, "--si",
                               "--diagnostico" if seleccion.diagnostico else None))
    titulo = "Ceses y renombres" if seleccion.grupo == "cese_renombre" else "Cambio, fe de erratas y rectificación"
    return lanzar(f"{titulo}: {len(filas)} tarifa(s)", comandos).resumen()


@app.get("/api/demo/excel")
def demo_excel():
    """Descarga el Excel de escenarios (datos/pruebas.xlsx) tal como está en disco."""
    if not datos_excel.ARCHIVO.exists():
        raise HTTPException(404, f"No existe {datos_excel.ARCHIVO}")
    return FileResponse(datos_excel.ARCHIVO, filename=datos_excel.ARCHIVO.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/demo/registrar")
def demo_registrar(seleccion: Seleccion):
    filas = filas_elegidas(seleccion.nombres)
    comandos = []
    for f in filas:
        if f["estado"] in ("Pendiente", "Error"):
            comandos.append(script("ejecutar.py", f["archivo"], "--lote", f["lote"], "--nombre", f["nombre"],
                                   "--alcance", f.get("alcance") or None, "--si",
                                   "--diagnostico" if seleccion.diagnostico else None))
    # Después, el registro definitivo de las que lo piden en el Excel (columna «Registro definitivo»).
    por_lote: dict[str, list[str]] = {}
    for f in filas:
        if f["quiere_definitivo"] and f["estado"] != "Definitivo":
            por_lote.setdefault(f["lote"], []).append(f["nombre"])
    for lote, nombres in por_lote.items():
        argumentos = ["--lote", lote, "--si"]
        for n in nombres:
            argumentos += ["--nombre", n]
        if seleccion.diagnostico:
            argumentos.append("--diagnostico")
        comandos.append(script("registro_definitivo.py", *argumentos))
    if not comandos:
        raise HTTPException(400, "Las tarifas elegidas ya están registradas en definitivo.")
    return lanzar(f"Registrar {len(filas)} tarifa(s) del Excel (+ registro definitivo)", comandos).resumen()


@app.post("/api/demo/definitivo")
def demo_definitivo(seleccion: Seleccion):
    filas = filas_elegidas(seleccion.nombres)
    sin_registro = [f["nombre"] for f in filas if f["estado"] not in ("Registrada", "Error definitivo", "Definitivo")]
    if sin_registro:
        raise HTTPException(400, "Primero hay que registrarlas: " + ", ".join(sin_registro))
    por_lote: dict[str, list[str]] = {}
    for f in filas:
        if f["estado"] != "Definitivo":
            por_lote.setdefault(f["lote"], []).append(f["nombre"])
    if not por_lote:
        raise HTTPException(400, "Las tarifas elegidas ya tienen registro definitivo.")
    comandos = []
    for lote, nombres in por_lote.items():
        argumentos = ["--lote", lote, "--si"]
        for n in nombres:
            argumentos += ["--nombre", n]
        if seleccion.diagnostico:
            argumentos.append("--diagnostico")
        comandos.append(script("registro_definitivo.py", *argumentos))
    total = sum(len(n) for n in por_lote.values())
    return lanzar(f"Registro definitivo de {total} tarifa(s) del Excel", comandos).resumen()


# --- Licitación y Contratos (datos/licitaciones.xlsx) -----------------------------

@app.get("/api/demo/licitaciones")
def demo_licitaciones():
    try:
        filas = datos_excel.leer_licitaciones()
    except ValueError as error:
        raise HTTPException(400, str(error))
    return {"archivo": str(datos_excel.ARCHIVO_LYC.relative_to(PROYECTO)), "licitaciones": filas}


@app.get("/api/demo/excel-licitaciones")
def demo_excel_licitaciones():
    """Descarga datos/licitaciones.xlsx tal como está en disco."""
    if not datos_excel.ARCHIVO_LYC.exists():
        raise HTTPException(404, f"No existe {datos_excel.ARCHIVO_LYC}")
    return FileResponse(datos_excel.ARCHIVO_LYC, filename=datos_excel.ARCHIVO_LYC.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/demo/licitaciones/registrar")
def demo_licitaciones_registrar(seleccion: Seleccion):
    """licitacion.py con las elegidas y, después, su registro definitivo (sin disponibilidad)."""
    try:
        filas = _elegidas(datos_excel.leer_licitaciones(), seleccion.nombres)
    except ValueError as error:
        raise HTTPException(400, str(error))
    comandos = []
    pendientes = [f["nombre"] for f in filas if f["estado"] in ("Pendiente", "Error")]
    if pendientes:
        comandos.append(script("licitacion.py", "--nombres", *pendientes,
                               "--diagnostico" if seleccion.diagnostico else None))
    comandos += comandos_definitivo(filas, seleccion.diagnostico)
    if not comandos:
        raise HTTPException(400, "Las licitaciones elegidas ya están en definitivo.")
    return lanzar(f"Registrar {len(filas)} licitación(es) y contrato(s) (+ registro definitivo)", comandos).resumen()


@app.get("/api/catalogo")
def api_catalogo():
    return {"servicios": catalogo.servicios(), "campos": catalogo.catalogo_campos(),
            "tarifas": catalogo.tarifas_conocidas(), "lotes": catalogo.lotes_conocidos()}


@app.get("/api/estado")
def api_estado():
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=1.5) as respuesta:
            navegador = json.loads(respuesta.read()).get("Browser", "")
        conectado = True
    except Exception:
        navegador, conectado = "", False
    activa = tarea_activa()
    return {"edge": conectado, "navegador": navegador, "tarea": activa.resumen() if activa else None}


@app.get("/api/tareas")
def api_tareas():
    return [t.resumen() for t in sorted(TAREAS.values(), key=lambda t: -t.id)]


@app.post("/api/tareas/{id_tarea}/cancelar")
def api_cancelar(id_tarea: int):
    tarea = TAREAS.get(id_tarea)
    if not tarea:
        raise HTTPException(404, "No existe esa tarea.")
    tarea.cancelar()
    return tarea.resumen()


@app.get("/api/tareas/{id_tarea}/salida")
def api_salida(id_tarea: int):
    """Salida en vivo (Server-Sent Events): cada línea nueva es un evento."""
    tarea = TAREAS.get(id_tarea)
    if not tarea:
        raise HTTPException(404, "No existe esa tarea.")

    def eventos():
        enviadas = 0
        while True:
            nuevas = tarea.lineas[enviadas:]
            for linea in nuevas:
                yield f"data: {json.dumps(linea, ensure_ascii=False)}\n\n"
            enviadas += len(nuevas)
            if tarea.estado not in ("en cola", "corriendo") and enviadas >= len(tarea.lineas):
                yield f"event: fin\ndata: {json.dumps(tarea.resumen(), ensure_ascii=False)}\n\n"
                return
            time.sleep(0.3)

    return StreamingResponse(eventos(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.get("/api/resultados")
def api_resultados():
    base = PROYECTO / "resultados"
    archivos = []
    for ruta in sorted(base.rglob("*.csv"), key=lambda r: r.stat().st_mtime, reverse=True):
        archivos.append({"ruta": ruta.relative_to(base).as_posix(),
                         "modificado": datetime.fromtimestamp(ruta.stat().st_mtime).strftime("%d/%m %H:%M")})
    return archivos


@app.get("/api/resultados/ver")
def api_resultado(ruta: str, ultimas: int = 300):
    base = (PROYECTO / "resultados").resolve()
    archivo = (base / ruta).resolve()
    if base not in archivo.parents or archivo.suffix != ".csv" or not archivo.exists():
        raise HTTPException(404, "Archivo no encontrado.")
    with archivo.open("r", newline="", encoding="utf-8-sig") as entrada:
        filas = list(csv.reader(entrada, delimiter=";"))
    if not filas:
        return {"columnas": [], "filas": [], "total": 0}
    return {"columnas": filas[0], "filas": filas[1:][-ultimas:][::-1], "total": len(filas) - 1}


def main() -> None:
    import uvicorn
    import webbrowser
    puerto = int(os.environ.get("PUERTO", "8000"))
    print(f"\n  Herramienta SIRT → http://127.0.0.1:{puerto}\n  (Ctrl+C para detener)\n")
    threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{puerto}")).start()
    uvicorn.run(app, host="127.0.0.1", port=puerto, log_level="warning")


if __name__ == "__main__":
    main()
