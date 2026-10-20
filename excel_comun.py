"""
Piezas comunes de los comandos que trabajan con los Excel de automatizacion/.

- Lectura de una hoja con cabecera en la fila 5 (formato de todas las plantillas).
- Escritura del resultado en las columnas grises (Estado, Código SIRT, Fecha, Detalle).
- Ejecución de los scripts ya probados (ejecutar, empaquetar, promocionar…) como
  subproceso, mostrando su salida en pantalla.
- Registro definitivo de una tarifa (registro_definitivo.py) y lectura de su resultado.
"""

import csv
import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent
AUTOMATIZACION = PROYECTO / "automatizacion"
RESULTADOS = PROYECTO / "resultados"
FILA_CABECERA = 5
LOTE_EXCEL = "EXCEL"          # lote con que se registra todo lo que viene de los Excel

COLORES = {
    "Registrada": "C6EFCE", "Aplicada": "C6EFCE", "Registro definitivo": "A9D08E",
    "Error": "FFC7CE", "Error de datos": "FFC7CE", "No aplica": "FFEB9C",
    "Ya registrada": "DDEBF7", "Ya aplicada": "DDEBF7", "Omitida": "EDEDED",
    "Hasta resumen": "FFF2CC",
}
SI = ("sí", "si")


class DatoInvalido(ValueError):
    """Un dato del Excel que impide ejecutar la fila (se escribe como «Error de datos»)."""


# --- Valores ------------------------------------------------------------------

def texto(valor) -> str:
    """Celda → texto (50.0 → «50»; None → «»)."""
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return "Sí" if valor else "No"
    if isinstance(valor, float):
        return str(int(valor)) if valor.is_integer() else repr(valor)
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


def es_si(valor) -> bool:
    return texto(valor).lower() in SI


def lista(valor) -> list[str]:
    """«A + B + C» → [A, B, C]. También acepta «;»."""
    t = texto(valor)
    if not t:
        return []
    separador = ";" if ";" in t and "+" not in t else "+"
    return [p.strip() for p in t.split(separador) if p.strip()]


def leer_fecha(valor, campo: str = "fecha") -> date | None:
    if valor in (None, ""):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    t = texto(valor)
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(t, formato).date()
        except ValueError:
            pass
    raise DatoInvalido(f"«{campo}»: fecha «{t}» no reconocida (usa dd/mm/aaaa)")


def numero(valor, campo: str, minimo=None, maximo=None, entero=False) -> str:
    """Valida un número del Excel y lo devuelve como texto para SIRT."""
    t = texto(valor).replace(",", ".")
    try:
        n = float(t)
    except ValueError:
        raise DatoInvalido(f"«{campo}» debe ser un número (tiene «{texto(valor)}»)")
    if entero and not n.is_integer():
        raise DatoInvalido(f"«{campo}» debe ser un número entero (tiene «{t}»)")
    if minimo is not None and n < minimo:
        raise DatoInvalido(f"«{campo}» debe ser al menos {minimo}")
    if maximo is not None and n > maximo:
        raise DatoInvalido(f"«{campo}» no puede pasar de {maximo}")
    return str(int(n)) if n.is_integer() else t


# --- Excel ----------------------------------------------------------------------

def excel_esta_abierto(ruta: Path) -> bool:
    """Windows: Excel bloquea el archivo. Linux/LibreOffice: deja un archivo .~lock."""
    if (ruta.parent / f".~lock.{ruta.name}#").exists() or (ruta.parent / f"~${ruta.name}").exists():
        return True
    try:
        with ruta.open("r+b"):
            return False
    except PermissionError:
        return True


def abrir_libro(ruta: Path, solo_listar: bool):
    from openpyxl import load_workbook
    if not ruta.exists():
        sys.exit(f"No existe {ruta}")
    if not solo_listar and excel_esta_abierto(ruta):
        sys.exit(f"El Excel está abierto: ciérralo y vuelve a correr el comando.\n  {ruta}")
    return load_workbook(ruta)


class Hoja:
    """Una hoja con cabecera en la fila 5 y columnas de resultado al final."""

    def __init__(self, hoja, obligatorias=(), estado="Estado", codigo="Código SIRT",
                 fecha=("Fecha de registro", "Fecha de ejecución"), detalle="Detalle"):
        self.hoja = hoja
        self.col = {}
        for celda in hoja[FILA_CABECERA]:
            if celda.value is not None and texto(celda.value) not in self.col:
                self.col[texto(celda.value)] = celda.column
        faltan = [t for t in obligatorias if t not in self.col]
        if faltan:
            sys.exit(f"Faltan columnas en la hoja «{hoja.title}»: {', '.join(faltan)}")
        fechas = (fecha,) if isinstance(fecha, str) else fecha
        self.c_estado = self.col.get(estado)
        self.c_codigo = self.col.get(codigo)
        self.c_fecha = next((self.col[f] for f in fechas if f in self.col), None)
        self.c_detalle = self.col.get(detalle)

    def filas(self, columna_nombre: str):
        """(número de fila, {título: valor}) de las filas con nombre."""
        c_nombre = self.col[columna_nombre]
        for r in range(FILA_CABECERA + 1, self.hoja.max_row + 1):
            if texto(self.hoja.cell(r, c_nombre).value):
                yield r, {t: self.hoja.cell(r, c).value for t, c in self.col.items()}

    def escribir(self, r: int, estado: str, codigo: str = "", fecha: str = "", detalle: str = "") -> None:
        from openpyxl.styles import Font, PatternFill
        for c, valor in ((self.c_estado, estado), (self.c_codigo, codigo),
                         (self.c_fecha, fecha), (self.c_detalle, detalle[:1000] if detalle else detalle)):
            if c:
                celda = self.hoja.cell(r, c, valor)
                celda.font = Font(name="Arial", size=9, bold=c == self.c_estado)
        if self.c_estado:
            self.hoja.cell(r, self.c_estado).fill = PatternFill("solid", fgColor=COLORES.get(estado, "EDEDED"))

    def estado(self, r: int) -> str:
        return texto(self.hoja.cell(r, self.c_estado).value) if self.c_estado else ""


def guardar(libro, ruta: Path) -> None:
    try:
        libro.save(ruta)
    except PermissionError:
        print("   Aviso: no se pudo guardar el Excel (¿se abrió?); se guardará al final en una copia.")


def guardar_con_copia(libro, ruta: Path, carpeta: str, prefijo: str) -> Path:
    copia = RESULTADOS / carpeta / f"{prefijo}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    copia.parent.mkdir(parents=True, exist_ok=True)
    try:
        libro.save(ruta)
        shutil.copyfile(ruta, copia)
    except PermissionError:
        libro.save(copia)
        print(f"No se pudo escribir en {ruta}; los resultados quedaron solo en la copia.")
    return copia


def confirmar(mensaje: str, si: bool) -> bool:
    if si:
        return True
    print(f"\nATENCIÓN: {mensaje}")
    return input("Escribe 'si' para continuar: ").strip().lower() in SI


def ejecutar_fila(fila: dict) -> str:
    """«Ejecutar»: vacío = Sí. Devuelve 'si', 'no' o el valor inválido."""
    t = texto(fila.get("Ejecutar")).lower()
    if t in ("", "sí", "si"):
        return "si"
    if t == "no":
        return "no"
    return t


# --- Subprocesos y CSV --------------------------------------------------------------

def correr(argumentos: list[str]) -> str:
    """Corre un script del proyecto, muestra su salida (con sangría) y la devuelve."""
    comando = [sys.executable, "-u", *argumentos]
    proceso = subprocess.Popen(comando, cwd=PROYECTO, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                               text=True, encoding="utf-8", errors="replace")
    lineas = []
    for linea in proceso.stdout:
        lineas.append(linea)
        if linea.strip() and not linea.startswith("=" * 20):
            print("   " + linea.rstrip())
    proceso.wait()
    salida = "".join(lineas)
    if "ECONNREFUSED" in salida and "9222" in salida:
        sys.exit("\nNo se pudo conectar con Edge (puerto 9222). Ábrelo en modo controlable, inicia sesión "
                 "en SIRT y vuelve a correr el comando (lo ya hecho quedó anotado en el Excel).")
    return salida


def marca_tiempo() -> str:
    return (datetime.now() - timedelta(seconds=2)).strftime("%Y-%m-%d %H:%M:%S")


def ultimo_resultado(ruta: Path, nombre: str, desde: str, clave_nombre: str = "nombre") -> dict | None:
    if not ruta.exists():
        return None
    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        filas = [f for f in csv.DictReader(entrada, delimiter=";")
                 if f.get(clave_nombre) == nombre and (f.get("fecha_hora") or "") >= desde]
    return filas[-1] if filas else None


def exitosa_en(ruta: Path, nombre: str) -> dict | None:
    if not ruta.exists():
        return None
    with ruta.open("r", newline="", encoding="utf-8-sig") as entrada:
        filas = [f for f in csv.DictReader(entrada, delimiter=";")
                 if f.get("nombre") == nombre and f.get("resultado") == "EXITOSO"]
    return filas[-1] if filas else None


def resumen_salida(salida: str, largo: int = 400) -> str:
    return " | ".join(l.strip() for l in salida.splitlines() if l.strip())[-largo:]


def hacer_registro_definitivo(nombre: str, lote: str, ubigeos: str = "", diagnostico: bool = False) -> tuple[bool, str]:
    """Registro definitivo (con disponibilidad si el servicio la pide) de UNA tarifa.
    Devuelve (éxito, detalle)."""
    ruta = RESULTADOS / f"registro_definitivo_{lote}.csv"
    if exitosa_en(ruta, nombre):
        return True, "ya tenía registro definitivo"
    desde = marca_tiempo()
    argumentos = ["registro_definitivo.py", "--lote", lote, "--nombre", nombre, "--si"]
    if ubigeos:
        argumentos += ["--ubigeos", ubigeos]
    if diagnostico:
        argumentos.append("--diagnostico")
    print(f"   → registro definitivo de {nombre}" + (f" (disponibilidad: {ubigeos})" if ubigeos else ""))
    salida = correr(argumentos)
    fila = ultimo_resultado(ruta, nombre, desde)
    if fila is None:
        return False, "registro definitivo sin resultado: " + resumen_salida(salida)
    if fila.get("resultado") == "EXITOSO":
        disp = fila.get("disponibilidad", "")
        return True, "registro definitivo OK" + (f" (disponibilidad: {disp})" if disp else "")
    return False, "registro definitivo con error: " + fila.get("detalle", "")
