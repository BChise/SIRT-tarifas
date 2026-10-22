"""
Acciones sobre tarifas ya registradas (página "Tarifas registradas").

Subcomandos:
    buscar    muestra código SIRT, versión y estados (comercialización / vigencia)
              de cada tarifa, sin cambiar nada.
    renombre  cambia el nombre de cada tarifa (por defecto: nombre actual + sufijo).
    cese      programa el cese de comercialización (fin de comercialización; por
              defecto mañana). Exige "En comercialización" y vigente; pide confirmación.
    fe_erratas agrega "| Fe de erratas <fecha hora>" a todos los campos de texto de
              los pasos 1 y 2 (cualquier servicio), revisa el resumen y guarda.
              --hasta-resumen se detiene en el paso 3 sin guardar.
    rectificar igual que fe_erratas (textos con "| Rectificación <fecha hora>") y además
              suma --delta (1) al Valor del servicio (paso 2) y a la Tarifa de
              instalación (paso 1, si está habilitada). No toca listas ni tablas.
    cambio    "Cambio de Tarifa": mismo flujo de 3 pasos + "Fecha de ejecución" (paso 1).
              Por defecto BAJA 1 el valor del servicio y la instalación (--delta -1) con
              fecha hoy + DIAS_ANTICIPACION_BAJA (mañana); si sube (--delta 1), hoy +
              DIAS_ANTICIPACION_ALZA. --fecha AAAA-MM-DD fija la fecha.

Flujo grabado (grabaciones/_acciones/):
    Tarifas registradas → Filtro avanzado → "Ingresa nombre de tarifa" → Buscar →
    fila → ⋮ Acciones → Renombre → "Ingresa el nombre de la tarifa" → Guardar →
    Guardar tarifa → Aceptar.

Los estados se leen de la respuesta del filtro (data.items: comercializacion,
vigencia, esUltimaVersion), no de la pantalla. Cese y fe de erratas exigirán
"En comercialización" y vigente; el renombre funciona también con "Por iniciar".

Uso:
    python acciones.py buscar --tarifas TE-SMO-PROMO-001 TE-IF-MUESTRA-001
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --solo-listar
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --diagnostico
    python acciones.py renombre --tarifas TE-SMO-PROMO-002 --nuevo-nombre TE-SMO-PROMO-002_NUEVO
    python acciones.py renombre --tarifas TE-AC-PROMO-001 TE-AC-PROMO-002 --sufijo _R2
    python acciones.py cese --tarifas TE-TF-MUESTRA-001 --solo-listar
    python acciones.py cese --tarifas TE-TF-MUESTRA-001 --fecha 2026-10-20 --diagnostico
    python acciones.py fe_erratas --tarifas TE-IF-MUESTRA-001 --hasta-resumen --diagnostico
    python acciones.py rectificar --tarifas TE-OS-MUESTRA-001 --hasta-resumen --diagnostico
    python acciones.py cambio --tarifas TE-OS-MUESTRA-001 --hasta-resumen --diagnostico
    python acciones.py cambio --tarifas TE-OS-MUESTRA-001 --delta 1 --diagnostico

Resultados (reanudable): resultados/ACCIONES/acciones.csv
"""

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import comun
from empaquetar import texto_dialogo, texto_visible
from ejecutar import Diagnostico


RUTA_RESULTADOS = comun.DIRECTORIO_RESULTADOS / "ACCIONES" / "acciones.csv"
CAMPOS = ["fecha_hora", "accion", "nombre", "codigo_sirt", "version", "nombre_nuevo",
          "comercializacion", "vigencia", "resultado", "detalle"]
TIMEOUT_BUSQUEDA = 20_000
SUFIJO_POR_DEFECTO = "_REN"

# --- Selección de campos (--campos "A;B;C") -----------------------------------------
# Si se pasa, fe de erratas / rectificar / cambio solo tocan los campos cuya etiqueta
# (o el título del bloque que los contiene) empieza por alguno de esos nombres.
# Sin --campos se tocan todos los que la acción cambia por defecto.
import unicodedata

FILTRO_CAMPOS: list[str] | None = None


def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip().lower()


# --- Sentido del cambio y detalle por campo (modificaciones.xlsx, 20/10) -------------
# SENTIDO = "beneficioso": los PRECIOS bajan y los BENEFICIOS (minutos, datos, velocidades…)
# suben; "perjudicial": al revés. None: el delta se aplica igual a todo (como antes).
# DETALLE: [(partes de la ruta normalizadas, regla)], regla = {"valor": X} o {"delta": N}.
#   Ruta («›»):  campo  ·  bloque › campo  ·  tabla  ·  tabla › fila  ·  tabla › fila › campo
#   (fila = el texto de una celda: 4G, ALEMANIA, Bolsa…; campo = etiqueta del número o de la lista).
#   Con DETALLE solo cambia lo que coincide con alguna ruta; el resto queda igual.
SENTIDO: str | None = None
DETALLE: list = []
SEPARADOR_RUTA = re.compile(r"\s*[›>]\s*")
PRECIO = re.compile(r"\b(valor|tarifa|precio|costo|cargo|instalacion|periodo de contrato)", re.I)


def es_precio(etiqueta: str) -> bool:
    return bool(PRECIO.search(normalizar(etiqueta)))


def cargar_detalle(crudo: dict) -> None:
    for ruta, regla in crudo.items():
        partes = [normalizar(p) for p in SEPARADOR_RUTA.split(ruta) if p.strip()]
        if partes:
            DETALLE.append((partes, regla))


def _empieza(texto: str, parte: str) -> bool:
    n = normalizar(texto)
    return bool(n) and n.startswith(parte)


def regla_detalle(etiqueta: str, contexto=(), tabla: str | None = None, celdas=()) -> dict | None:
    """La regla del DETALLE que corresponde a un campo, o None si ninguna ruta lo nombra."""
    for partes, regla in DETALLE:
        if tabla is not None:                            # número / lista dentro de una fila de tabla
            if not _empieza(tabla, partes[0]):
                continue
            resto = partes[1:]
            if not resto:
                return regla
            if len(resto) == 1:
                if any(normalizar(c) == resto[0] for c in celdas) or _empieza(etiqueta, resto[0]):
                    return regla
                continue
            if any(normalizar(c) == resto[0] for c in celdas) and _empieza(etiqueta, resto[1]):
                return regla
            continue
        if len(partes) == 1 and (_empieza(etiqueta, partes[0]) or any(_empieza(c, partes[0]) for c in contexto)):
            return regla
        if len(partes) == 2 and any(_empieza(c, partes[0]) for c in contexto) and _empieza(etiqueta, partes[1]):
            return regla
    return None


def ajuste(etiqueta: str, delta: Decimal, contexto=(), tabla: str | None = None, celdas=()):
    """Qué hacer con un campo: ('valor', '25'), ('delta', Decimal) o None (no tocarlo)."""
    if DETALLE:
        regla = regla_detalle(etiqueta, contexto, tabla, celdas)
        if regla is None:
            return None
        return ("valor", str(regla["valor"])) if "valor" in regla else ("delta", Decimal(str(regla["delta"])))
    if SENTIDO:
        magnitud = abs(delta)
        sube = (not es_precio(etiqueta)) if SENTIDO == "beneficioso" else es_precio(etiqueta)
        return "delta", magnitud if sube else -magnitud
    return "delta", delta


def campo_elegido(*etiquetas: str) -> bool:
    if FILTRO_CAMPOS is None:
        return True
    for etiqueta in etiquetas:
        e = normalizar(etiqueta)
        if e and any(e.startswith(f) for f in FILTRO_CAMPOS):
            return True
    return False


# --- Cambio de tarifa: anticipación de la "Fecha de ejecución" (días desde hoy) ------
# SIRT (16/10, al SUBIR valor del servicio e instalación con fecha de hoy): "La fecha de
# ejecución de la acción no puede ser anterior a 2 días a partir de la fecha actual,
# según los tipos de variación aplicados". Si OSIPTEL cambia el plazo, se cambia aquí.
# 19/10: un cambio con -1 que también editaba tablas (--todo) fue rechazado con fecha de
# mañana («no puede ser anterior a 2 días…»): bajar velocidades o gigas perjudica al
# usuario aunque baje el precio. Por ahora se usan 3 días en ambos casos.
DIAS_ANTICIPACION_ALZA = 3      # antes 2: cambio perjudicial para el usuario (sube precios)
DIAS_ANTICIPACION_BAJA = 3      # antes 1 (el calendario siempre bloquea HOY)
# Si el día calculado está bloqueado en el calendario, se prueban hasta estos días siguientes.
DIAS_EXTRA_SI_BLOQUEADO = 7


# --- Resultados ------------------------------------------------------------------

def registrar(fila: dict) -> None:
    RUTA_RESULTADOS.parent.mkdir(parents=True, exist_ok=True)
    nuevo = not RUTA_RESULTADOS.exists()
    with RUTA_RESULTADOS.open("a", newline="", encoding="utf-8-sig") as salida:
        escritor = csv.DictWriter(salida, fieldnames=CAMPOS, delimiter=";")
        if nuevo:
            escritor.writeheader()
        escritor.writerow({c: fila.get(c, "") for c in CAMPOS})


def ya_hechas(accion: str) -> set[str]:
    if not RUTA_RESULTADOS.exists():
        return set()
    with RUTA_RESULTADOS.open("r", newline="", encoding="utf-8-sig") as entrada:
        return {f["nombre"] for f in csv.DictReader(entrada, delimiter=";")
                if f["accion"] == accion and f["resultado"] == "EXITOSO"}


# --- Búsqueda en "Tarifas registradas" ----------------------------------------------

def abrir_tarifas_registradas(page) -> None:
    origen = re.match(r"https?://[^/]+", page.url).group(0)
    page.keyboard.press("Escape")
    page.goto(f"{origen}/extranet/consulta-tarifa", wait_until="domcontentloaded")
    page.get_by_role("button", name=re.compile("Filtro (avanzado|principal)")).first.wait_for(state="visible")


def abrir_filtro_avanzado(page) -> None:
    campo = page.get_by_role("textbox", name="Ingresa nombre de tarifa")
    if campo.count() and campo.is_visible():
        return
    page.get_by_role("button", name="Filtro avanzado").click()
    campo.wait_for(state="visible")


def _sin_espacios(texto: str) -> str:
    return re.sub(r"\s+", "", texto or "")


def buscar_tarifa(page, nombre: str) -> dict | None:
    """Busca por nombre y devuelve la última versión con ese nombre EXACTO:
    {id, codigo, version, nombre, comercializacion, vigencia, condicion} o None.
    Los datos salen de la respuesta del filtro (la tabla no muestra la vigencia)."""
    items = items_del_filtro(page, nombre)
    exactos = [i for i in items if _sin_espacios(i.get("nombre")) == _sin_espacios(nombre)]
    if not exactos:
        return None
    ultimo = max(exactos, key=lambda i: (bool(i.get("esUltimaVersion")), i.get("version") or 0))
    page.wait_for_timeout(300)        # que la tabla termine de pintarse
    return datos_de_item(ultimo)


def datos_de_item(ultimo: dict) -> dict:
    return {
        "id": ultimo.get("idTarifa"),
        "codigo": ultimo.get("codigoSIRT", ""),
        "version": str(ultimo.get("version", "")),
        "nombre": ultimo.get("nombre", ""),
        "comercializacion": (ultimo.get("comercializacion") or {}).get("nombre", ""),
        "vigencia": (ultimo.get("vigencia") or {}).get("nombre", ""),
        "condicion": (ultimo.get("condicion") or {}).get("nombre", ""),
        "tipo": ((ultimo.get("tipoTarifa") or {}).get("nombre", "")),
    }


def items_del_filtro(page, nombre: str, con_total: bool = False):
    """Filtra «Tarifas registradas» por nombre y devuelve los items del JSON de la respuesta
    (con_total=True → (items, total que dice SIRT))."""
    respuestas = []

    def capturar(respuesta):
        pedido = respuesta.request
        if pedido.method == "POST" and pedido.resource_type in ("xhr", "fetch"):
            respuestas.append(respuesta)

    abrir_filtro_avanzado(page)
    # SIRT a veces deja filtros de la búsqueda anterior (Condición, Tipo de servicio,
    # Modalidad: visto el 16/10 tras un cambio en SMC → "no se encontraron resultados"
    # al buscar una SMP). Se limpian antes de buscar por nombre.
    limpiar = page.get_by_role("button", name="Limpiar filtro")
    if limpiar.count() and limpiar.first.is_visible():
        limpiar.first.click()
        try:
            page.wait_for_load_state("networkidle", timeout=5_000)
        except Exception:
            pass
        page.wait_for_timeout(500)
        abrir_filtro_avanzado(page)
    page.get_by_role("textbox", name="Ingresa nombre de tarifa").fill(nombre)

    page.on("response", capturar)
    try:
        page.get_by_role("button", name="Buscar", exact=True).click()
        items = None
        for _ in range(TIMEOUT_BUSQUEDA // 250):
            for respuesta in list(respuestas):
                try:
                    datos = respuesta.json().get("data")
                except Exception:
                    respuestas.remove(respuesta)
                    continue
                if isinstance(datos, dict) and isinstance(datos.get("items"), list):
                    items = datos["items"]
                    total = datos.get("total") or len(items)
                    break
                respuestas.remove(respuesta)
            if items is not None:
                break
            page.wait_for_timeout(250)
    finally:
        page.remove_listener("response", capturar)

    if items is None:
        raise RuntimeError(f"No llegó la respuesta del filtro al buscar '{nombre}'.")
    return (items, total) if con_total else items


def fila_en_tabla(page, tarifa: dict):
    """La fila de la tabla con ese código SIRT y versión (columnas leídas por su título)."""
    js = """
        ([codigo, version]) => {
            const titulos = Array.from(document.querySelectorAll("thead th"))
                .map(th => th.innerText.normalize("NFD").replace(/[^A-Za-z ]/g, "").replace(/\\s+/g, " ").trim());
            const col = t => titulos.findIndex(x => x.startsWith(t));
            const cCod = col("Codigo SIRT"), cVer = col("Version");
            return Array.from(document.querySelectorAll("tbody tr")).findIndex(fila => {
                const celdas = Array.from(fila.querySelectorAll("td")).map(td => td.innerText.trim());
                const esCodigo = cCod >= 0 ? celdas[cCod] === codigo : celdas.includes(codigo);
                return esCodigo && (cVer < 0 || celdas[cVer] === version);
            });
        }
    """
    argumentos = [tarifa["codigo"], tarifa["version"]]
    try:   # la tabla se pinta un instante después de llegar la respuesta
        page.wait_for_function(f"a => ({js})(a) >= 0", arg=argumentos, timeout=TIMEOUT_BUSQUEDA)
    except Exception:
        raise RuntimeError(f"La tabla no muestra {tarifa['codigo']} versión {tarifa['version']}.")
    indice = page.evaluate(js, argumentos)
    return page.locator("tbody tr").nth(indice)


def abrir_accion(page, fila, accion: str) -> None:
    fila.get_by_label("Acciones").click()
    opcion = page.get_by_role("menuitem", name=accion)
    opcion.wait_for(state="visible")
    if not opcion.is_enabled():
        page.keyboard.press("Escape")
        raise RuntimeError(f"La acción '{accion}' está deshabilitada para esta tarifa.")
    opcion.click()


def confirmar_con_aceptar(page) -> str:
    """Lee el mensaje final, pulsa Aceptar y avisa si es un error."""
    aceptar = page.get_by_role("button", name="Aceptar", exact=True)
    aceptar.first.wait_for(state="visible")
    mensaje = texto_dialogo(page)
    aceptar.first.click()
    # Esperar a que se cierre ESTA ventana: o desaparece el botón, o aparece otra
    # ventana distinta (SIRT a veces muestra dos "Aceptar" seguidos).
    for _ in range(40):                                  # hasta ~10 s
        page.wait_for_timeout(250)
        if not aceptar.first.is_visible() or texto_dialogo(page) != mensaje:
            break
    if re.search(r"acci[oó]n de cambio de tarifa pendiente", mensaje, re.I):
        # 19/10: no se puede programar otro cambio mientras haya uno pendiente de ejecutar.
        raise NoAplica("la tarifa ya tiene un CAMBIO DE TARIFA programado y pendiente; hasta que se "
                       "ejecute (en su fecha) no se puede programar otro. Nada se guardó.")
    if re.search(r"error|no se pudo|no se cumplieron|object reference", mensaje, re.I):
        raise RuntimeError(f"SIRT respondió: {mensaje[:400]}")
    return mensaje


# --- Renombre --------------------------------------------------------------------

def renombrar(page, tarifa: dict, nombre_nuevo: str) -> str:
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, "Renombre")

    campo = page.get_by_role("textbox", name="Ingresa el nombre de la tarifa")
    campo.wait_for(state="visible")
    campo.fill(nombre_nuevo)
    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    mensaje = confirmar_con_aceptar(page)

    # Verificación: la tarifa debe encontrarse con el nombre nuevo.
    abrir_tarifas_registradas(page)
    renombrada = buscar_tarifa(page, nombre_nuevo)
    if renombrada is None or renombrada["codigo"] != tarifa["codigo"]:
        raise RuntimeError(f"Tras el renombre, '{nombre_nuevo}' no aparece con el código {tarifa['codigo']}.")
    return mensaje


# --- Cese de comercialización ------------------------------------------------------

def _renovacion_ya_guardada(nombre: str, tarifa: dict) -> bool:
    """¿Una renovación anterior que quedó en ERROR sí se guardó? Se reconoce porque la promoción
    tiene una versión > 1 «Por iniciar comercialización» (la que crea la renovación)."""
    try:
        version = int(tarifa.get("version") or 1)
    except ValueError:
        return False
    if version < 2 or "por iniciar comercializ" not in tarifa.get("comercializacion", "").lower():
        return False
    if not RUTA_RESULTADOS.exists():
        return False
    with RUTA_RESULTADOS.open("r", newline="", encoding="utf-8-sig") as entrada:
        previas = [f for f in csv.DictReader(entrada, delimiter=";")
                   if f.get("accion") == "renovar" and f.get("nombre") == nombre]
    resultados = [f.get("resultado") for f in previas]
    return "ERROR" in resultados and "EXITOSO" not in resultados


class NoAplica(Exception):
    """La tarifa no está en el estado que exige la acción: no se tocó SIRT."""


def validar_para_cese(tarifa: dict, sin_validar: bool, accion: str = "el cese") -> None:
    """SIRT da error si la tarifa no está En comercialización y vigente."""
    problemas = []
    # Licitación y Contratos no se comercializa: SIRT muestra siempre «No aplica» en
    # comercialización y aun así permite las acciones (19/10). Basta que esté vigente.
    licitacion = re.search(r"licitaci", tarifa.get("tipo", ""), re.I) and \
        re.search(r"no aplica", tarifa["comercializacion"], re.I)
    if "en comercializ" not in tarifa["comercializacion"].lower() and not licitacion:
        problemas.append(f"comercialización '{tarifa['comercializacion']}'")
    if re.search(r"por iniciar|no vigente|fin", tarifa["vigencia"], re.I):
        problemas.append(f"vigencia '{tarifa['vigencia']}'")
    if problemas and not sin_validar:
        raise NoAplica(f"No se puede aplicar {accion}: " + " y ".join(problemas)
                           + " (debe estar En comercialización y vigente; pasa a ese estado el día"
                             " siguiente al registro definitivo). Usa --sin-validar para intentarlo igual.")


def elegir_fecha_calendario(page, fecha: date, dias_extra: int = 0) -> date:
    """Abre el único calendario del formulario y avanza mes a mes hasta la fecha.
    Si el día está bloqueado y `dias_extra` > 0, prueba los días siguientes (hasta
    `dias_extra`) y devuelve la fecha que eligió."""
    page.get_by_role("button", name="Open calendar").first.click()
    for extra in range(dias_extra + 1):
        dia_buscado = fecha + timedelta(days=extra)
        texto = f"{dia_buscado.day} de {comun.MESES[dia_buscado.month - 1]} de {dia_buscado.year}"
        for _ in range(24):
            dia = page.get_by_role("button", name=texto, exact=True)
            if dia.count():
                break
            siguiente = page.get_by_role("button", name="Next month")
            if not siguiente.is_enabled():
                break
            siguiente.click()
        if dia.count() and dia.first.is_enabled():
            dia.click()
            page.locator("mat-calendar").wait_for(state="hidden")
            return dia_buscado
    page.keyboard.press("Escape")
    raise RuntimeError(f"El calendario no permite elegir el {fecha:%d/%m/%Y}"
                       + (f" ni los {dias_extra} días siguientes" if dias_extra else "") + " (día bloqueado).")


def cesar(page, tarifa: dict, fecha: date) -> str:
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, "Cese de Comercialización")
    page.get_by_text("Fin de comercialización").first.wait_for(state="visible")
    # Si el día pedido está bloqueado en el calendario, se usa el siguiente permitido (19/10).
    elegida = elegir_fecha_calendario(page, fecha, DIAS_EXTRA_SI_BLOQUEADO)
    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    mensaje = confirmar_con_aceptar(page)
    return (f"fecha usada {elegida:%d/%m/%Y} (pedida {fecha:%d/%m/%Y}, estaba bloqueada) | " if elegida != fecha
            else "") + mensaje


# --- Fe de erratas ----------------------------------------------------------------

def _texto_nuevo(actual: str, sello: str, prefijo: str, maximo: int | None) -> str:
    """Quita un sello anterior DEL MISMO TIPO ('| <prefijo> dd/mm/aaaa hh:mm') y agrega el
    nuevo al final, sin pasar del máximo de caracteres del campo. Los sellos de otras
    acciones (p. ej. una fe de erratas previa) se conservan."""
    patron = rf"\s*\|\s*{re.escape(prefijo)} \d{{2}}/\d{{2}}/\d{{4}} \d{{2}}:\d{{2}}\s*$"
    base = re.sub(patron, "", actual or "").rstrip()
    nuevo = f"{base} {sello}".strip()
    if maximo and len(nuevo) > maximo:
        nuevo = (base[: max(0, maximo - len(sello) - 1)] + " " + sello)[:maximo]
    return nuevo


# Marca las textarea de la pestaña ACTIVA: SIRT deja en la página las del paso 2
# (pestaña oculta) y a veces se ven un instante mientras carga (SMQ 16/10).
MARCAR_TEXTOS = """() => {
  let n = 0; const etiquetas = [];
  for (const t of document.querySelectorAll('textarea')) {
    t.removeAttribute('data-fe');
    const r = t.getBoundingClientRect();
    if (t.disabled || t.readOnly || r.width === 0 || r.height === 0) continue;
    let ok = true;
    for (let a = t.parentElement; a && ok; a = a.parentElement) {
      if (a.tagName === 'MAT-TAB-BODY' && !a.classList.contains('mat-mdc-tab-body-active')) ok = false;
      if (a.getAttribute('aria-hidden') === 'true') ok = false;
      const s = getComputedStyle(a);
      if (s.display === 'none' || s.visibility === 'hidden') ok = false;
    }
    if (ok) {
      t.setAttribute('data-fe', String(n++));
      const ff = t.closest('mat-form-field');
      const et = ff && ff.querySelector('mat-label, label');
      let lab = (et && et.innerText) || t.getAttribute('aria-label') || '';
      if (!lab) {                       // la etiqueta suele estar antes del mat-form-field
        let prev = (ff || t).previousElementSibling;
        lab = prev ? prev.innerText : (t.placeholder || '').replace(/^Ingrese\\s*/i, '');
      }
      etiquetas.push(lab.replace(/\\s+/g, ' ').trim());
    }
  }
  return etiquetas;
}"""


def corregir_textos(page, sello: str, prefijo: str) -> list[str]:
    """Agrega el sello a los campos de texto (textarea) visibles y editables de la
    pestaña activa. Sirve para cualquier servicio: solo cambia cuántos campos hay."""
    corregidos = []
    etiquetas = page.evaluate(MARCAR_TEXTOS)
    for i, etiqueta in enumerate(etiquetas):
        campo = page.locator(f'textarea[data-fe="{i}"]')
        if not campo_elegido(etiqueta, campo.get_attribute("placeholder") or ""):
            continue
        try:
            if not (campo.is_visible() and campo.is_editable()):
                continue
            maximo = campo.get_attribute("maxlength")
            nuevo = _texto_nuevo(campo.input_value(), sello, prefijo, int(maximo) if maximo else None)
            campo.fill(nuevo, timeout=5_000)
        except Exception:
            if campo.count() and campo.is_visible():
                raise                                    # visible y aun así falló: error real
            continue                                     # se ocultó (no era de este paso)
        corregidos.append(nuevo)
    return corregidos


# Campos de texto de UNA línea (input, no textarea) que también se corrigen.
# Licitación y Contratos (grabación 19/10): «Nombre de empresa» recibe el sello y
# «Documento de Empresa» (solo dígitos, máx. 15) cambia su última cifra.
# Promociones (19/10): «Otras condiciones tarifa» (si es input; si es textarea ya la
# corrige corregir_textos).
INPUTS_TEXTO = [("Nombre de empresa", "texto"), ("Otras condiciones tarifa", "texto"),
                ("Documento de Empresa", "digitos")]


# ¿El elemento está en la pestaña/paso ACTIVO? SIRT deja en la página los paneles de los
# otros pasos y Playwright los ve como visibles: el 19/10 «Documento de Empresa» (paso 1)
# se volvió a cambiar en cada pestaña del paso 2 y el resumen no cuadró.
EN_PANEL_ACTIVO = """e => {
  for (let a = e; a; a = a.parentElement) {
    if (a.tagName === 'MAT-TAB-BODY' && !a.classList.contains('mat-mdc-tab-body-active')) return false;
    if (a.getAttribute && a.getAttribute('aria-hidden') === 'true') return false;
    const s = getComputedStyle(a);
    if (s.display === 'none' || s.visibility === 'hidden') return false;
  }
  const r = e.getBoundingClientRect();
  return r.width > 0 && r.height > 0;
}"""


# Igual, pero sin exigir tamaño (las casillas de Angular Material son input invisibles).
EN_PANEL_ACTIVO_SIN_TAMANO = EN_PANEL_ACTIVO.split("  const r = ")[0] + "  return true;\n}"


def corregir_inputs(page, sello: str, prefijo: str, hechos: set | None = None) -> list[str]:
    """Corrige los INPUTS_TEXTO visibles y editables del paso/pestaña activa, cada uno
    una sola vez por acción (`hechos`). Devuelve los valores nuevos (para el resumen)."""
    nuevos = []
    hechos = set() if hechos is None else hechos
    for etiqueta, tipo in INPUTS_TEXTO:
        if etiqueta in hechos or not campo_elegido(etiqueta):
            continue
        campo = page.get_by_role("textbox", name=etiqueta, exact=True)
        if campo.count() == 0:
            continue
        campo = campo.first
        try:
            if not (campo.is_visible() and campo.is_editable()) or not campo.evaluate(EN_PANEL_ACTIVO):
                continue
            if campo.evaluate("e => e.tagName") == "TEXTAREA":
                continue                                     # ya la trató corregir_textos
            actual = campo.input_value()
            if tipo == "digitos":
                digitos = re.sub(r"\D", "", actual)
                nuevo = (digitos[:-1] + str((int(digitos[-1]) + 1) % 10)) if digitos else "1"
            else:
                maximo = campo.get_attribute("maxlength")
                nuevo = _texto_nuevo(actual, sello, prefijo, int(maximo) if maximo else None)
            campo.fill(nuevo, timeout=5_000)
            campo.press("Tab")
            quedo = campo.input_value()
            if quedo != nuevo:                               # máscara o validación del campo
                print(f"    Aviso: {etiqueta} quedó como «{quedo}» (se escribió «{nuevo}»)")
                nuevo = quedo
        except Exception:
            if campo.is_visible():
                raise
            continue
        hechos.add(etiqueta)
        print(f"    {etiqueta}: {actual} → {nuevo}")
        nuevos.append(nuevo)
    return nuevos


def _numero_nuevo(actual: str, delta: Decimal) -> str:
    """'50' + 1 → '51'; '0.5' + 1 → '1.5' (mantiene los decimales que tenía)."""
    from decimal import Decimal as D
    texto = (actual or "").strip().replace(",", "")
    valor = D(texto) if texto else D(0)
    decimales = len(texto.split(".")[1]) if "." in texto else 0
    return f"{valor + delta:.{decimales}f}"


def _releer(page, campo) -> str:
    """Valor que quedó en el campo tras salir de él (Tab)."""
    campo.press("Tab")
    page.wait_for_timeout(250)
    return (campo.input_value() or "").strip().replace(",", "")


def _mismo_numero(a: str, b: str) -> bool:
    try:
        return Decimal(a) == Decimal(b)
    except Exception:
        return a == b


def cambiar_numero(page, etiqueta: re.Pattern, nombre: str, delta: Decimal,
                   si_vacio: Decimal | None = None) -> str | None:
    """Cambia el campo numérico visible y habilitado cuya etiqueta coincide:
    - con número: le suma `delta` (50 → 51);
    - vacío de verdad (tras esperar hasta ~5 s a que SIRT lo llene): pone `si_vacio`
      (por defecto 0 + delta, o sea 1).
    Devuelve 'Etiqueta: antes → después' o None si el campo no existe o está bloqueado."""
    if not campo_elegido(nombre):
        return None
    resultado = ajuste(nombre, delta)
    if resultado is None:
        return None
    tipo, regla = resultado
    for rol in ("textbox", "spinbutton"):
        for campo in page.get_by_role(rol, name=etiqueta).all():
            if not (campo.is_visible() and campo.is_editable()):
                continue
            antes = campo.input_value()
            for _ in range(10):                          # ¿aún no cargó?
                if antes.strip():
                    break
                page.wait_for_timeout(500)
                antes = campo.input_value()
            if tipo == "valor":
                despues = regla                          # valor pedido en la hoja «Campos»
            elif antes.strip():
                despues = _numero_nuevo(antes, regla)
                if Decimal(despues) < 0:
                    return None                          # no se puede bajar de 0
            elif si_vacio is not None:
                despues = str(si_vacio)
            elif regla > 0:
                despues = str(regla)
            else:
                return None                              # vacío y se quería bajar: no se toca
            campo.fill(despues)
            # Lo que el campo REALMENTE guardó (algunos tienen máscara y reescriben el número: visto
            # el 22/10 con «Valor de contrato» de TL-LYC-CM-123, escrito 1773.00 y mostrado 1773.03).
            # El resumen se compara con este valor.
            nota = ""
            try:
                quedo = _releer(page, campo)
                if quedo and not _mismo_numero(quedo, despues):
                    # 2.º intento: borrar todo y escribir tecla por tecla (la máscara lo acepta mejor).
                    campo.click()
                    campo.press("Control+A")
                    campo.press("Delete")
                    campo.press_sequentially(despues, delay=40)
                    quedo = _releer(page, campo)
                if quedo and not _mismo_numero(quedo, despues):
                    nota = f" (se pidió {despues}; la máscara de SIRT dejó otro valor)"
                    despues = quedo                      # el resumen se compara con lo que quedó
            except Exception:
                pass
            return f"{nombre}: {antes.strip() or '(vacío)'}{nota} → {despues}"
    return None


# Listas de selección múltiple que la fe de erratas / rectificación también cambia
# (hoy solo PTV: "Tipo de Contratación" con Postpago / Prepago).
LISTAS_PASO_2 = [(re.compile(r"^Tipo de Contrataci[oó]n", re.I), "Tipo de Contratación")]
# Con --todo se cambian además estas listas múltiples (móviles).
LISTAS_TODO = [(re.compile(r"^Apps sin consumo", re.I), "Apps sin consumo de datos")]

# --- --todo: TODOS los campos numéricos del paso 2 ----------------------------------
# Marca los input con número de la pestaña activa (como MARCAR_TEXTOS) y devuelve su
# etiqueta y valor. Excluye fechas, listas y campos bloqueados.
MARCAR_NUMEROS = """() => {
  const out = []; let n = 0;
  for (const i of document.querySelectorAll('input')) {
    i.removeAttribute('data-num');
    if (i.disabled || i.readOnly) continue;
    if (!['text', 'number', ''].includes(i.type || '')) continue;
    if (i.getAttribute('role') === 'combobox') continue;
    const r = i.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) continue;
    let ok = true;
    for (let a = i.parentElement; a && ok; a = a.parentElement) {
      if (a.tagName === 'MAT-TAB-BODY' && !a.classList.contains('mat-mdc-tab-body-active')) ok = false;
      if (a.getAttribute('aria-hidden') === 'true') ok = false;
      // Paneles de OTROS pasos que SIRT deja en la página (22/10: «Valor de contrato» del
      // paso 1 subía 0.01 en cada pestaña del paso 2 → 1773.00 quedó 1773.03 con 3 servicios).
      if (a.getAttribute('aria-expanded') === 'false') ok = false;
      const s = getComputedStyle(a);
      if (s.display === 'none' || s.visibility === 'hidden') ok = false;
    }
    if (!ok) continue;
    const v = (i.value || '').trim().replace(/,/g, '');
    if (!/^-?\\d+(\\.\\d+)?$/.test(v)) continue;
    const ff = i.closest('mat-form-field');
    const et = ff && ff.querySelector('mat-label, label');
    const lab = ((et && et.innerText) || i.getAttribute('aria-label') || i.placeholder || '').trim();
    const ctx = [];                     // primera línea de los contenedores (título del bloque)
    for (let a = i.parentElement, k = 0; a && k < 8; a = a.parentElement, k++) {
      const linea = (a.innerText || '').split('\\n')[0].trim();
      if (linea && linea.length < 90) ctx.push(linea);
    }
    i.setAttribute('data-num', String(n));
    out.push({n: n++, lab: lab.replace(/\\s+/g, ' '), v, ctx});
  }
  return out;
}"""
# Campos del paso 1 que «--todo» (paso 2) nunca debe tocar aunque SIRT los deje en la página.
NUNCA_EN_TODO = [re.compile(r"^Valor de contrato", re.I), re.compile(r"^Duraci[oó]n de contrato", re.I),
                 re.compile(r"^Documento de Empresa", re.I)]
# Topes de los que dependen otros campos (tethering, traslada datos y datos roaming
# incluidos deben ser <= "Datos tope descarga"): al SUBIR se cambian primero, al BAJAR al final.
TOPES = re.compile(r"tope", re.I)
PORCENTAJE = re.compile(r"%|porcentaje|degradaci", re.I)


def cambiar_todos_los_numeros(page, delta: Decimal, excluir: list[re.Pattern]) -> list[str]:
    """Suma `delta` a todos los campos numéricos visibles y editables de la pestaña activa
    (datos, velocidades, minutos, SMS, tarifas por minuto, roaming...). No toca los que
    coinciden con `excluir` (se cambian aparte). Cambia 1 unidad del último decimal
    (0.10 → 0.11, 12 → 13); no toca los que valen 0, no baja de 0 ni pasa de 100 en %.
    Las tablas (filas ya registradas) NO se tocan: no son campos editables en pantalla."""
    campos = page.evaluate(MARCAR_NUMEROS)
    campos = [c for c in campos if not any(e.search(c["lab"]) for e in excluir)
              and campo_elegido(c["lab"], *c.get("ctx", []))]
    for c in campos:
        c["regla"] = ajuste(c["lab"], delta, c.get("ctx", []))
    campos = [c for c in campos if c["regla"] is not None]
    # Topes: si suben, primero; si bajan, al final (los que dependen de ellos deben ser <= tope).
    campos.sort(key=lambda c: 0 if TOPES.search(c["lab"]) and c["regla"][0] == "delta" and c["regla"][1] > 0
                else 2 if TOPES.search(c["lab"]) else 1)
    hechos = []
    for c in campos:
        tipo, regla = c["regla"]
        if tipo == "valor":
            nuevo = regla
        else:
            if Decimal(c["v"]) == 0:
                continue                                 # 0 suele ser "no aplica": no se toca
            decimales = len(c["v"].split(".")[1]) if "." in c["v"] else 0
            paso = regla * (Decimal(1).scaleb(-decimales))   # 0.10 → 0.11 (no 1.10); 12 → 13
            nuevo = _numero_nuevo(c["v"], paso)
        if Decimal(nuevo) <= 0 or (PORCENTAJE.search(c["lab"]) and Decimal(nuevo) > 100):
            continue
        campo = page.locator(f'input[data-num="{c["n"]}"]')
        try:
            campo.fill(nuevo, timeout=5_000)
        except Exception:
            continue
        hechos.append(f'{c["lab"] or "(sin etiqueta)"}: {c["v"]} → {nuevo}')
    return hechos


def cambiar_lista_multiple(page, etiqueta: re.Pattern, nombre: str) -> tuple[str, list[str]] | None:
    """Cambia una lista de casillas (mat-select multiple) sin dejarla vacía:
    con 2 o más marcadas desmarca la última; con 1 marca la primera que falte.
    Devuelve ('Nombre: antes → después', opciones marcadas al final) o None si no existe."""
    if not campo_elegido(nombre):
        return None
    for combo in page.get_by_role("combobox", name=etiqueta).all():
        if not (combo.is_visible() and combo.is_enabled()):
            continue
        antes = " ".join(combo.inner_text().split()) or "(vacío)"
        combo.click()
        lista = page.get_by_role("listbox").last
        lista.wait_for(state="visible")
        if lista.get_attribute("aria-multiselectable") == "false":
            page.keyboard.press("Escape")                # lista de una sola opción: no se toca
            lista.wait_for(state="hidden")
            return None
        opciones = lista.get_by_role("option").all()
        marcadas = [o for o in opciones if o.get_attribute("aria-selected") == "true"]
        libres = [o for o in opciones if o.get_attribute("aria-selected") != "true"]
        if len(marcadas) >= 2:
            marcadas[-1].click()
        elif libres:
            libres[0].click()
        page.wait_for_timeout(300)
        finales = [" ".join(o.inner_text().split()) for o in lista.get_by_role("option").all()
                   if o.get_attribute("aria-selected") == "true"]
        page.keyboard.press("Escape")                    # la grabación cerraba con clic fuera
        lista.wait_for(state="hidden")
        despues = ", ".join(finales) or "(vacío)"
        if not finales:
            raise RuntimeError(f"{nombre} quedó vacío.")
        return f"{nombre}: {antes} → {despues}", finales
    return None


# --- Rectificación de Licitación y Contratos (grabación 19/10) ----------------------
# Paso 1 editable: Inicio y Fin de vigencia (solo con el calendario: escribir la fecha
# falla), Ámbito, Valor de contrato, Moneda y los textos. «Duración de contrato» e
# «Identificador de contrato» quedan bloqueados (la duración la recalcula SIRT).

# Los campos de fecha de esta pantalla no tienen nombre accesible (la grabación usa
# #mat-input-14 y «Open calendar» .first / .nth(1)): se ubican por orden entre los
# botones «Open calendar» del paso activo: 0 = Inicio de vigencia, 1 = Fin de vigencia.
ORDEN_FECHAS = {"Inicio de vigencia": 0, "Fin de vigencia": 1,
                # Renovación de promociones: las dos fechas de «Datos a renovar de la tarifa».
                "Fecha inicio de la comercialización": 0, "Fecha fin de la comercialización": 1}


def _boton_calendario(page, etiqueta: str):
    activos = [b for b in page.get_by_role("button", name="Open calendar").all()
               if b.is_visible() and b.evaluate(EN_PANEL_ACTIVO)]
    indice = ORDEN_FECHAS[etiqueta]
    return activos[indice] if len(activos) > indice else None


def _leer_fecha(page, etiqueta: str) -> date | None:
    boton = _boton_calendario(page, etiqueta)
    if boton is None:
        return None
    texto = boton.evaluate("""b => { const f = b.closest('mat-form-field');
                                     const i = f && f.querySelector('input');
                                     return i ? i.value : ''; }""")
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y").date()
    except ValueError:
        print(f"    Aviso: no se pudo leer «{etiqueta}» ({texto!r}); no se cambia")
        return None


def elegir_fecha_campo(page, etiqueta: str, fecha: date, dias_extra: int = 7) -> date | None:
    """Calendario del campo `etiqueta`: avanza (o retrocede) mes a mes hasta el día.
    Si está bloqueado prueba los siguientes. None si el campo no se puede editar."""
    boton = _boton_calendario(page, etiqueta)
    if boton is None or not boton.is_enabled():
        print(f"    Aviso: «{etiqueta}» no tiene calendario habilitado; no se cambia")
        return None
    boton.click()
    page.locator("mat-calendar").wait_for(state="visible")
    for extra in range(dias_extra + 1):
        dia_buscado = fecha + timedelta(days=extra)
        texto = f"{dia_buscado.day} de {comun.MESES[dia_buscado.month - 1]} de {dia_buscado.year}"
        dia = page.get_by_role("button", name=texto, exact=True)
        for direccion, veces in (("Next month", 40), ("Previous month", 80)):
            for _ in range(veces):
                if dia.count():
                    break
                mover = page.get_by_role("button", name=direccion)
                if not mover.is_enabled():
                    break
                mover.click()
            if dia.count():
                break
        if dia.count() and dia.first.is_enabled():
            dia.first.click()
            page.locator("mat-calendar").wait_for(state="hidden")
            return dia_buscado
    page.keyboard.press("Escape")
    raise RuntimeError(f"El calendario de «{etiqueta}» no permite elegir el {fecha:%d/%m/%Y}"
                       f" ni los {dias_extra} días siguientes.")


def cambiar_lista_unica(page, etiqueta: re.Pattern, nombre: str) -> tuple[str, str] | None:
    """Lista de una sola opción (Ámbito, Moneda): elige la primera opción distinta."""
    if not campo_elegido(nombre):
        return None
    for combo in page.get_by_role("combobox", name=etiqueta).all():
        if not (combo.is_visible() and combo.is_enabled()):
            continue
        antes = " ".join(combo.inner_text().split())
        combo.click()
        lista = page.get_by_role("listbox").last
        lista.wait_for(state="visible")
        otras = [o for o in lista.get_by_role("option").all()
                 if o.get_attribute("aria-selected") != "true"
                 and " ".join(o.inner_text().split()) not in ("", antes)
                 and not o.inner_text().strip().lower().startswith("seleccione")]
        if not otras:
            page.keyboard.press("Escape")
            return None
        despues = " ".join(otras[0].inner_text().split())
        otras[0].click()
        lista.wait_for(state="hidden")
        return f"{nombre}: {antes} → {despues}", despues
    return None


def rectificar_licitacion_paso_1(page, delta: Decimal) -> tuple[list[str], list[str]]:
    """Solo si la pantalla es de Licitación (tiene «Valor de contrato»). Devuelve
    (cambios, valores a buscar en el resumen)."""
    cambios, valores = [], []
    if page.get_by_role("textbox", name=re.compile(r"^Valor de contrato")).count() == 0:
        return cambios, valores
    manana = date.today() + timedelta(days=1)            # hoy siempre está bloqueado

    inicio, fin = _leer_fecha(page, "Inicio de vigencia"), _leer_fecha(page, "Fin de vigencia")
    if inicio and campo_elegido("Inicio de vigencia"):
        pedido = max(inicio + timedelta(days=1), manana)
        if fin and pedido >= fin:
            pedido = None                                # no se puede mover sin pasar el fin
        elegida = elegir_fecha_campo(page, "Inicio de vigencia", pedido) if pedido else None
        if elegida:
            cambios.append(f"Inicio de vigencia: {inicio:%d/%m/%Y} → {elegida:%d/%m/%Y}")
            valores.append((f"{elegida:%d/%m/%Y}", f"{elegida:%Y-%m-%d}"))   # el resumen usa 2026-11-29
    if fin and campo_elegido("Fin de vigencia"):
        elegida = elegir_fecha_campo(page, "Fin de vigencia", fin + timedelta(days=1))
        if elegida:
            cambios.append(f"Fin de vigencia: {fin:%d/%m/%Y} → {elegida:%d/%m/%Y}")
            valores.append((f"{elegida:%d/%m/%Y}", f"{elegida:%Y-%m-%d}"))
    duracion = page.get_by_role("textbox", name=re.compile(r"^Duraci[oó]n de contrato"))
    if duracion.count():
        cambios.append(f"Duración (la calcula SIRT): {duracion.first.input_value()}")

    hecho = cambiar_lista_unica(page, re.compile(r"^Ámbito"), "Ámbito")
    if hecho:
        cambios.append(hecho[0]); valores.append(hecho[1])
    hecho = cambiar_numero(page, re.compile(r"^Valor de contrato"), "Valor de contrato", delta)
    if hecho:
        cambios.append(hecho); valores.append(hecho.split("→")[-1].strip())
    hecho = cambiar_lista_unica(page, re.compile(r"^Moneda"), "Moneda")
    if hecho:
        cambios.append(hecho[0]); valores.append(hecho[1])
    for c in cambios:
        print(f"    {c}")
    return cambios, valores


# --- Renovación de tarifas promocionales (capturas y JSON del 19/10) -----------------
# Paso 1: «Datos a renovar de la tarifa»: Fecha inicio / fin de la comercialización
#         (vacías, con calendario) y la «Duración» del valor del plan (la de tarifa de
#         instalación solo si la promoción la tiene).
# Paso 2: una pestaña «Servicio N» por servicio con los MISMOS atributos que se
#         promocionaron; solo la «Duración» de cada uno es editable.
# Regla de la duración: si vale 25 → 31; cualquier otro valor → 25.
# El JSON que SIRT pide al abrir «Renovar» (modalidades → apartados) se captura para
# verificar que en cada pestaña haya tantas duraciones como atributos promocionados.

# Fechas (acordado el 19/10): el calendario de inicio abre en el día siguiente al fin de
# comercialización actual (p. ej. fin 14/12 → abre en 15/12). Inicio = ese día + 1
# (16/12) y fin = inicio + 19 (04/01) — antes 3; cambiado el 19/10. Si cae en otro mes o año se calcula igual.
DIAS_DESPUES_DEL_DEFECTO = 1
DIAS_COMERCIALIZACION_RENOVACION = 19

MARCAR_DURACIONES = """() => {
  let n = 0; const valores = [];
  for (const e of document.querySelectorAll('input')) {
    e.removeAttribute('data-dur');
    if (e.disabled || e.readOnly || e.type === 'hidden') continue;
    const r = e.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    let ok = true;
    for (let a = e; a && ok; a = a.parentElement) {
      if (a.tagName === 'MAT-TAB-BODY' && !a.classList.contains('mat-mdc-tab-body-active')) ok = false;
      if (a.getAttribute && a.getAttribute('aria-hidden') === 'true') ok = false;
      const s = getComputedStyle(a);
      if (s.display === 'none' || s.visibility === 'hidden') ok = false;
    }
    if (!ok) continue;
    // Contenedor más cercano que solo tiene ESTE input: ¿dice «Duración»?
    let es = /Duraci[oó]n/i.test((e.getAttribute('aria-label') || '') + ' ' +
              (e.labels ? [...e.labels].map(l => l.innerText).join(' ') : ''));
    for (let a = e.parentElement, k = 0; !es && a && k < 6; a = a.parentElement, k++) {
      if (a.querySelectorAll('input').length > 1) break;
      if (/Duraci[oó]n/i.test(a.innerText || '')) es = true;
    }
    if (es) { e.setAttribute('data-dur', String(n++)); valores.push(e.value); }
  }
  return valores;
}"""


DURACION_FIJA: str | None = None          # --duracion-renovacion: todas las duraciones a este valor


def duracion_nueva(actual: str) -> str:
    if DURACION_FIJA:
        return DURACION_FIJA
    return "31" if actual.strip() == "25" else "25"


def cambiar_duraciones(page, donde: str) -> list[str]:
    """Cambia todas las «Duración» editables del paso/pestaña activa. Devuelve los cambios."""
    cambios = []
    valores = page.evaluate(MARCAR_DURACIONES)
    for i, antes in enumerate(valores):
        campo = page.locator(f'input[data-dur="{i}"]')
        nuevo = duracion_nueva(antes)
        campo.fill(nuevo)
        campo.press("Tab")
        quedo = campo.input_value()
        if quedo != nuevo:
            raise RuntimeError(f"{donde}: la duración {i + 1} quedó en «{quedo}» (se escribió {nuevo}).")
        cambios.append(f"{antes or '(vacío)'}→{nuevo}")
    return cambios


def _json_renovacion(respuestas) -> dict | None:
    """Entre las respuestas capturadas, la que trae modalidades y diasRenovacionAparatado."""
    for respuesta in reversed(respuestas):
        try:
            if respuesta.request.resource_type not in ("xhr", "fetch") or respuesta.status != 200:
                continue
            texto = respuesta.text()
            if "diasRenovacion" not in texto:
                continue
            datos = json.loads(texto).get("data") or {}
            if "modalidades" in datos:
                return datos
        except Exception:
            continue
    return None


def _atributos_por_servicio(datos: dict) -> list[tuple[str, list[str]]]:
    salida = []
    for modalidad in datos.get("modalidades") or []:
        nombres = [a.get("nombre", "?") for sec in modalidad.get("seccionesServicio") or []
                   for a in sec.get("apartados") or []]
        salida.append((modalidad.get("nombre", "?"), nombres))
    return salida


def _fin_comercial_actual(page, datos: dict | None) -> date | None:
    """Fin de comercialización actual: del JSON (finComercial) o de la pantalla."""
    texto = ((datos or {}).get("tarifa") or {}).get("finComercial") or ""
    if texto:
        try:
            return date.fromisoformat(texto[:10])
        except ValueError:
            pass
    encontrado = re.search(r"Fin comercializaci[oó]n:?\s*(\d{2}/\d{2}/\d{4})", page.locator("body").inner_text())
    return datetime.strptime(encontrado.group(1), "%d/%m/%Y").date() if encontrado else None


def renovar(page, tarifa: dict, hasta_resumen: bool, inicio: date | None) -> tuple[list[str], str]:
    respuestas = []
    escucha = lambda r: respuestas.append(r)
    page.on("response", escucha)
    try:
        abrir_accion(page, fila_en_tabla(page, tarifa), "Renovar")
        page.get_by_role("button", name="Continuar con el paso 2").wait_for(state="visible")
        try:
            page.wait_for_load_state("networkidle", timeout=5_000)
        except Exception:
            pass
        page.wait_for_timeout(500)
    finally:
        page.remove_listener("response", escucha)
    cambios = []

    datos = _json_renovacion(respuestas)
    servicios_json = _atributos_por_servicio(datos) if datos else []
    if datos:
        for nombre, atributos in servicios_json:
            print(f"    JSON: {nombre}: {len(atributos)} atributo(s) promocionado(s): {', '.join(atributos)}")
        maximo = datos.get("diasRenovacionFecha")
    else:
        print("    Aviso: no se pudo leer el JSON de la renovación; se sigue solo con la pantalla.")
        maximo = None

    # Paso 1: fechas de comercialización y duraciones (valor del plan / instalación).
    dias = DIAS_COMERCIALIZACION_RENOVACION
    if maximo and dias > int(maximo):
        dias = int(maximo)
    if inicio is None:
        fin_actual = _fin_comercial_actual(page, datos)
        if fin_actual is None:
            raise RuntimeError("No se pudo leer el fin de comercialización actual; usa --fecha AAAA-MM-DD.")
        por_defecto = fin_actual + timedelta(days=1)
        inicio = por_defecto + timedelta(days=DIAS_DESPUES_DEL_DEFECTO)
        print(f"    Fin de comercialización actual {fin_actual:%d/%m/%Y} → el calendario abre en "
              f"{por_defecto:%d/%m/%Y}; inicio {inicio:%d/%m/%Y}, fin {inicio + timedelta(days=dias):%d/%m/%Y}")
    elegido_inicio = elegir_fecha_campo(page, "Fecha inicio de la comercialización", inicio)
    if elegido_inicio is None:
        raise RuntimeError("No se pudo abrir el calendario de «Fecha inicio de la comercialización».")
    elegido_fin = elegir_fecha_campo(page, "Fecha fin de la comercialización",
                                     elegido_inicio + timedelta(days=dias))
    if elegido_fin is None:
        raise RuntimeError("No se pudo abrir el calendario de «Fecha fin de la comercialización».")
    cambios.append(f"Comercialización: {elegido_inicio:%d/%m/%Y} → {elegido_fin:%d/%m/%Y}")
    paso1 = cambiar_duraciones(page, "Paso 1")
    if not paso1:
        raise RuntimeError("Paso 1: no se encontró ninguna «Duración» editable (valor del plan).")
    cambios.append(f"Paso 1 duración(es): {', '.join(paso1)}")

    continuar3 = page.get_by_role("button", name="Continuar con el paso 3")
    avanzar(page, page.get_by_role("button", name="Continuar con el paso 2"), continuar3, "paso 2")
    page.wait_for_timeout(500)

    # Paso 2: cada pestaña «Servicio N».
    pestanas = page.get_by_role("tab", name=re.compile(r"^Servicio \d+$"))
    total = max(1, pestanas.count())
    if servicios_json and len(servicios_json) != total:
        raise RuntimeError(f"El JSON trae {len(servicios_json)} servicio(s) y la pantalla {total} pestaña(s).")
    pendientes = list(servicios_json)
    for i in range(total):
        if pestanas.count():
            pestanas.nth(i).click()
            page.wait_for_timeout(500)
        # ¿Qué servicio del JSON muestra esta pestaña? Por su nombre; si no, por orden.
        del_json = next((s for s in pendientes if texto_visible(page, s[0])), None) \
            or (pendientes[0] if pendientes else None)
        if del_json:
            pendientes.remove(del_json)
        hechos = cambiar_duraciones(page, f"Servicio {i + 1}")
        nombre = del_json[0] if del_json else f"Servicio {i + 1}"
        # Un atributo con subapartados (p. ej. velocidades de Internet fijo) muestra una duración por
        # subapartado, así que en pantalla puede haber MÁS duraciones que atributos en el JSON. Solo es
        # error si hay MENOS (quedaría algún atributo promocionado sin renovar).
        if del_json and len(hechos) < len(del_json[1]):
            raise RuntimeError(f"Servicio {i + 1} ({nombre}): el JSON tiene {len(del_json[1])} atributo(s) "
                               f"promocionado(s) y en pantalla solo hay {len(hechos)} duración(es) editables.")
        detalle = ", ".join(f"{a} {c}" for a, c in zip(del_json[1], hechos)) if del_json else ", ".join(hechos)
        cambios.append(f"Servicio {i + 1} – {nombre}: {detalle or 'sin atributos'}")

    # Paso 3: resumen y guardado.
    final = page.get_by_role("button", name=re.compile(r"^Guardar cambios$"))     # grabación 19/10
    avanzar(page, continuar3, final.first, "paso 3 (resumen)")
    boton = " ".join(final.first.inner_text().split())
    resumen = _sin_espacios(page.locator("body").inner_text())
    fechas = [(f"{d:%d/%m/%Y}", f"{d:%Y-%m-%d}") for d in (elegido_inicio, elegido_fin)]
    faltan = [f[0] for f in fechas if not any(_sin_espacios(x) in resumen for x in f)]
    if faltan:
        raise RuntimeError(f"El resumen no muestra la(s) fecha(s) {', '.join(faltan)}.")
    if hasta_resumen:
        return cambios, f"detenido en el resumen (botón final: «{boton}», sin guardar)"
    return cambios, terminar_guardando(page, boton)


# --- Canales de TV de paga (carga de Excel) en rectificación y cambio (19/10) -----------
# El número de canales lo calcula SIRT desde el Excel subido (campo bloqueado). Para
# cambiarlo se sube otro Excel: si la fila tiene el original, se sube el «_cambio»
# (3 → 5 canales) y viceversa, así cada acción alterna entre los dos.
CARPETA_CANALES = Path(__file__).parent / "automatizacion" / "canales"
if not CARPETA_CANALES.exists():
    CARPETA_CANALES = Path(__file__).parent / "datos" / "canales"
ARCHIVOS_CANALES = ("Canales analógicos_1.xlsx", "Canales analógicos_1_cambio.xlsx")
FILAS_CANALES = ["Canales analógicos", "Canales digitales definición estándar",
                 "Canales digitales alta definición"]


def _numero_de_fila(fila) -> str:
    for campo in fila.locator("input:not([type='file'])").all():
        valor = (campo.input_value() or "").strip()
        if valor:
            return valor
    return ""


def cambiar_canales(page) -> list[str]:
    """Sube el otro Excel de canales en cada fila de canales del paso/pestaña activa."""
    cambios = []
    for etiqueta in FILAS_CANALES:
        if not campo_elegido(etiqueta):
            continue
        filas = page.locator(f"xpath=//*[normalize-space(text())='{etiqueta}']"
                             "/ancestor::*[.//input[@type='file'] or count(.//button) >= 2][1]")
        fila = next((f for f in filas.all() if f.is_visible() and f.evaluate(EN_PANEL_ACTIVO)), None)
        if fila is None:
            continue
        campo_archivo = fila.locator("input[type='file']")
        if campo_archivo.count() == 0 or campo_archivo.first.is_disabled():
            print(f"    {etiqueta}: la carga de archivo no está habilitada; no se cambia")
            continue
        texto = fila.inner_text()
        nuevo = ARCHIVOS_CANALES[0] if ARCHIVOS_CANALES[1] in texto else ARCHIVOS_CANALES[1]
        ruta = CARPETA_CANALES / nuevo
        if not ruta.exists():
            raise FileNotFoundError(f"No existe {ruta} (Excel de canales para el cambio).")
        antes = _numero_de_fila(fila)

        def subir() -> bool:
            campo_archivo.first.set_input_files(str(ruta))
            try:
                fila.get_by_text(nuevo, exact=True).wait_for(state="visible", timeout=8_000)
                return True
            except Exception:
                return False

        if not subir():
            # Respaldo: quitar el archivo actual con la ✕ (último botón de la fila) y volver a subir.
            fila.locator("button").last.click()
            page.wait_for_timeout(500)
            for texto_boton in ("Aceptar", "Sí", "Si", "Confirmar"):
                boton = page.locator("mat-dialog-container").get_by_role("button", name=texto_boton, exact=True)
                if boton.count() and boton.first.is_visible():
                    boton.first.click()
                    break
            if not subir():
                raise RuntimeError(f"{etiqueta}: SIRT no mostró el archivo «{nuevo}» después de subirlo.")
        despues = antes
        for _ in range(20):                                    # hasta ~5 s a que SIRT recalcule
            despues = _numero_de_fila(fila)
            if despues != antes:
                break
            page.wait_for_timeout(250)
        aviso = "" if despues != antes else " (el número no cambió)"
        cambios.append(f"{etiqueta}: {antes or '?'} → {despues or '?'} canales ({nuevo}){aviso}")
    return cambios


# --- Filas de tablas en rectificación y cambio (grabación tablas_movil, 19/10) -------
# Cada fila tiene en su última columna dos botones: lápiz (editar, el 1.º) y papelera.
# El lápiz abre la misma ventana que «Agregar»; se cambian SOLO sus números (Gigas,
# las 4 velocidades, Tarifa) y se pulsa «Guardar». No se tocan Tipo, Tecnología,
# Países ni Plataformas: SIRT rechaza países duplicados en una tabla («no debe contener
# países duplicados») y repetir tecnología/plataformas podría pasar lo mismo.
# Si algún número quedaría por debajo de 0 (cambio −1 sobre 0.5), esa fila no se toca.
MARCAR_FILAS = """() => {
  const salida = []; let t = 0;
  for (const tabla of document.querySelectorAll('table')) {
    tabla.removeAttribute('data-tabla');
    let ok = true;
    for (let a = tabla; a && ok; a = a.parentElement) {
      if (a.tagName === 'MAT-TAB-BODY' && !a.classList.contains('mat-mdc-tab-body-active')) ok = false;
      if (a.getAttribute && a.getAttribute('aria-hidden') === 'true') ok = false;
      const s = getComputedStyle(a);
      if (s.display === 'none' || s.visibility === 'hidden') ok = false;
    }
    if (!ok || tabla.closest('mat-dialog-container')) continue;
    const filas = [...tabla.querySelectorAll('tbody tr')].filter(f => {
      const celdas = f.querySelectorAll('td');
      return celdas.length && celdas[celdas.length - 1].querySelectorAll('button').length >= 2;
    });
    if (!filas.length) continue;
    // Título: primera línea del contenedor más cercano que tenga texto antes de la tabla.
    let titulo = '';
    for (let a = tabla.parentElement, k = 0; a && k < 6 && !titulo; a = a.parentElement, k++) {
      const linea = (a.innerText || '').split('\\n').map(x => x.trim()).find(x => x);
      if (linea && !/^No\\.?$/i.test(linea)) titulo = linea;
    }
    tabla.setAttribute('data-tabla', String(t));
    salida.push({t: t++, titulo, filas: filas.length});
  }
  return salida;
}"""


def _boton_editar(page, t: int, i: int):
    filas = page.locator(f'table[data-tabla="{t}"] tbody tr').filter(
        has=page.locator("td:last-child button"))
    return filas.nth(i).locator("td").last.locator("button").first


def cambiar_listas_de_fila(page, dialogo, tabla: str, celdas: list) -> list[tuple[str, str]]:
    """Detalle con «Valor nuevo» para una lista de la fila (p. ej. «Tecnologías y velocidades › 4G ›
    Tecnología» = 4.5G): elige esa opción. SIRT rechaza repetir tecnología / país en la tabla."""
    hechos = []
    for lista in dialogo.get_by_role("combobox").all():
        if not lista.is_visible():
            continue
        etiqueta = lista.evaluate("""e => { const f = e.closest('mat-form-field');
            const l = f && f.querySelector('mat-label, label');
            return ((l && l.innerText) || e.getAttribute('aria-label') || '').trim(); }""")
        # Solo rutas que terminan en el nombre de ESTA lista: «tabla › lista» o «tabla › fila › lista».
        regla = next((r for partes, r in DETALLE if "valor" in r and len(partes) >= 2
                      and _empieza(tabla, partes[0]) and _empieza(etiqueta, partes[-1])
                      and (len(partes) == 2 or any(normalizar(c) == partes[1] for c in celdas))), None)
        if regla is None:
            continue
        antes = (lista.inner_text() or "").strip()
        lista.click()
        opcion = page.locator(".cdk-overlay-container").get_by_role("option", name=str(regla["valor"]), exact=True)
        if opcion.count() == 0:
            page.keyboard.press("Escape")
            raise RuntimeError(f"{tabla}: «{etiqueta}» no tiene la opción «{regla['valor']}».")
        opcion.first.click()
        fondo = page.locator(".cdk-overlay-backdrop.cdk-overlay-backdrop-showing")
        if fondo.count():                                # lista múltiple: se cierra sin cerrar la ventana
            fondo.last.click(position={"x": 5, "y": 5})
        hechos.append((f"{etiqueta}: {antes}", str(regla["valor"])))
    return hechos


def editar_filas_de_tablas(page, delta: Decimal) -> list[str]:
    """Edita todas las filas de las tablas visibles del paso/pestaña activa."""
    cambios = []
    tablas = page.evaluate(MARCAR_FILAS)
    for tabla in tablas:
        if not campo_elegido(tabla["titulo"]):
            continue
        for i in range(tabla["filas"]):
            page.evaluate(MARCAR_FILAS)                      # la tabla se vuelve a pintar tras guardar
            celdas = [x.strip() for x in page.locator(f'table[data-tabla="{tabla["t"]}"] tbody tr').filter(
                has=page.locator("td:last-child button")).nth(i).locator("td").all_inner_texts()]
            if DETALLE and not any(_empieza(tabla["titulo"], partes[0]) and
                                   (len(partes) == 1 or any(normalizar(c) == partes[1] for c in celdas)
                                    or len(partes) == 2) for partes, _ in DETALLE):
                continue                                     # fila no nombrada en el detalle
            _boton_editar(page, tabla["t"], i).click()
            dialogo = page.locator("mat-dialog-container").last
            dialogo.wait_for(state="visible")
            page.wait_for_timeout(300)
            campos = [c for c in dialogo.get_by_role("spinbutton").all() if c.is_visible() and c.is_editable()]
            nuevos = []
            for campo in campos:
                antes = (campo.input_value() or "").strip()
                if not antes:
                    continue
                etiqueta = campo.evaluate("""e => { const f = e.closest('mat-form-field');
                    const l = f && f.querySelector('mat-label, label');
                    return ((l && l.innerText) || e.getAttribute('aria-label') || e.placeholder || '').trim(); }""")
                resultado = ajuste(etiqueta, delta, tabla=tabla["titulo"], celdas=celdas)
                if resultado is None:
                    continue
                tipo, regla = resultado
                despues = regla if tipo == "valor" else _numero_nuevo(antes, regla)
                nuevos.append((campo, antes, despues))
            listas = cambiar_listas_de_fila(page, dialogo, tabla["titulo"], celdas) if DETALLE else []
            if (not nuevos and not listas) or any(Decimal(d) < 0 for _, _, d in nuevos):
                motivo = "sin números" if not nuevos else "quedaría por debajo de 0"
                silencioso = DETALLE and not nuevos and not listas    # fila abierta, pero sin nada pedido
                cerrar = dialogo.get_by_role("button", name=re.compile(r"Cancelar|Cerrar|close", re.I))
                if cerrar.count():
                    cerrar.first.click()
                else:
                    page.keyboard.press("Escape")
                dialogo.wait_for(state="hidden")
                if not silencioso:
                    cambios.append(f"{tabla['titulo']} fila {i + 1}: no se cambió ({motivo})")
                continue
            for campo, _, despues in nuevos:
                campo.fill(despues)
            dialogo.get_by_role("button", name="Guardar", exact=True).click()
            nuevos = nuevos + [(None, a, d) for a, d in listas]
            try:
                dialogo.wait_for(state="hidden", timeout=10_000)
            except Exception:
                errores = " ".join(dialogo.locator("mat-error").all_inner_texts()) or texto_dialogo(page)
                raise RuntimeError(f"{tabla['titulo']} fila {i + 1}: la ventana no se cerró al guardar. {errores[:200]}")
            cambios.append(f"{tabla['titulo']} fila {i + 1}: " + ", ".join(f"{a}→{d}" for _, a, d in nuevos))
    return cambios


# --- Agregar filas a las tablas (--agregar-filas, grabación tablas_movil 19/10) -------
# Una fila nueva por tabla con «Agregar». Lo que no puede repetirse se elige entre las
# opciones que NO están ya en la tabla: País (SIRT rechaza países duplicados),
# Tecnología y Plataformas. Si no queda ninguna libre, la tabla se deja como está.
VALORES_FILA_NUEVA = {"Tarifa": "1", "Gigas": "1", "Velocidad": "1"}

USADOS_EN_TABLA = """b => {
  for (let a = b.parentElement; a; a = a.parentElement) {
    const t = a.querySelector('table');
    if (t) {
      const usados = [];
      t.querySelectorAll('tbody td').forEach(td => (td.innerText || '').split(/[,\\n]/)
        .forEach(x => { x = x.trim().toUpperCase(); if (x) usados.push(x); }));
      let titulo = (a.innerText || '').split('\\n').map(x => x.trim()).find(x => x) || '';
      return {usados, titulo};
    }
  }
  return {usados: [], titulo: ''};
}"""


def _cerrar_lista(page) -> None:
    """Cierra una lista múltiple abierta sin cerrar la ventana (Escape la cerraría)."""
    fondo = page.locator(".cdk-overlay-backdrop.cdk-overlay-transparent-backdrop")
    if fondo.count() and fondo.last.is_visible():
        fondo.last.click()
        page.wait_for_timeout(300)


def _elegir_libre(page, combo, usados: set, preferida: str | None = None) -> str | None:
    """Abre la lista y elige `preferida` o la primera opción que no esté en `usados`."""
    combo.click()
    lista = page.get_by_role("listbox").last
    lista.wait_for(state="visible")
    elegida = None
    opciones = lista.get_by_role("option").all()
    nombres = [" ".join(o.inner_text().split()) for o in opciones]
    orden = ([nombres.index(preferida)] if preferida in nombres else []) + list(range(len(opciones)))
    for k in orden:
        if nombres[k] and nombres[k].upper() not in usados and opciones[k].get_attribute("aria-selected") != "true":
            opciones[k].click()
            elegida = nombres[k]
            break
    page.wait_for_timeout(300)
    if lista.is_visible():
        _cerrar_lista(page)
    return elegida


def agregar_filas_a_tablas(page) -> list[str]:
    cambios = []
    total = len([b for b in page.get_by_role("button", name="Agregar", exact=True).all()
                 if b.is_visible() and b.evaluate(EN_PANEL_ACTIVO)])
    for k in range(total):
        botones = [b for b in page.get_by_role("button", name="Agregar", exact=True).all()
                   if b.is_visible() and b.evaluate(EN_PANEL_ACTIVO)]
        if k >= len(botones):
            break
        info = botones[k].evaluate(USADOS_EN_TABLA)
        usados, titulo = set(info["usados"]), info["titulo"] or f"tabla {k + 1}"
        if not campo_elegido(titulo):
            continue
        botones[k].click()
        dialogo = page.locator("mat-dialog-container").last
        dialogo.wait_for(state="visible")
        page.wait_for_timeout(300)
        elegido, motivo = [], ""

        def combo(nombre):
            busc = dialogo.get_by_role("combobox", name=re.compile(rf"^{nombre}"))
            return busc.first if busc.count() and busc.first.is_visible() else None

        # «Tipo» primero (de él depende qué campos aparecen): Bolsa en las bolsas de gigas
        # (con Ilimitado no hay gigas); si no, la primera opción. Solo si está vacío.
        c = combo("Tipo")
        if c is not None and c.evaluate("e => e.classList.contains('mat-mdc-select-empty')"):
            valor = _elegir_libre(page, c, set(), preferida="Bolsa")
            if valor:
                elegido.append(f"Tipo: {valor}")
                page.wait_for_timeout(300)
        # Listas que deben ser distintas a las de la tabla.
        for nombre in ("Países", "Tecnología", "Plataformas"):
            c = combo(nombre)
            if c is None:
                continue
            valor = _elegir_libre(page, c, usados)
            if valor is None:
                motivo = f"no queda {nombre.lower()} libre"
                break
            elegido.append(f"{nombre}: {valor}")
        if not motivo:
            for campo in dialogo.get_by_role("spinbutton").all():
                if not (campo.is_visible() and campo.is_editable()):
                    continue
                nombre = re.sub(r"\s+", " ", campo.get_attribute("aria-label") or campo.evaluate(
                    "e => ((e.closest('mat-form-field') || e.parentElement).innerText || '').split('\\n')[0]")).strip()
                valor = next((v for k2, v in VALORES_FILA_NUEVA.items() if k2.lower() in nombre.lower()), "1")
                campo.fill(valor)
                elegido.append(f"{(nombre or 'número')[:30]}={valor}")
        if motivo:
            cerrar = dialogo.get_by_role("button", name=re.compile(r"Cancelar|Cerrar|close", re.I))
            if cerrar.count():
                cerrar.first.click()
            else:
                page.keyboard.press("Escape")
            dialogo.wait_for(state="hidden")
            cambios.append(f"{titulo}: no se agregó fila ({motivo})")
            continue
        dialogo.get_by_role("button", name="Guardar", exact=True).click()
        try:
            dialogo.wait_for(state="hidden", timeout=10_000)
        except Exception:
            errores = " ".join(dialogo.locator("mat-error").all_inner_texts()) or texto_dialogo(page)
            raise RuntimeError(f"{titulo}: no se pudo agregar la fila. {errores[:200]}")
        cambios.append(f"{titulo}: fila agregada ({'; '.join(elegido)})")
    return cambios


# --- Condicionales: de «ilimitado» a limitado (--quitar-ilimitado) -------------------
# Grabación tablas_movil 19/10: se desmarca la casilla y se llenan sus dos campos
# (cantidad y tarifa adicional; en la grabación 10 y 15).
VALORES_LIMITADO = ("10", "15")


def quitar_ilimitado(page) -> list[str]:
    cambios = []
    for casilla in page.locator("input[type='checkbox']").all():
        try:
            if not casilla.is_checked() or casilla.is_disabled() or not casilla.evaluate(EN_PANEL_ACTIVO_SIN_TAMANO):
                continue
        except Exception:
            continue
        rotulo = casilla.evaluate("""e => { const l = e.closest('mat-checkbox, label, .mdc-form-field');
                                            return l ? l.innerText.trim() : ''; }""")
        if not re.search(r"ilimitad", rotulo, re.I):
            continue
        titulo = casilla.evaluate("""e => { for (let a = e.parentElement, k = 0; a && k < 6; a = a.parentElement, k++) {
                  const l = (a.innerText || '').split('\\n').map(x => x.replace(/es ilimitad[oa]/ig, '').trim()).find(x => x);
                  if (l) return l; } return ''; }""")
        if not campo_elegido(titulo):
            continue
        casilla.uncheck(force=True)
        page.wait_for_timeout(300)
        campos = [casilla.locator(f"xpath=following::input[not(@type='checkbox')][{n}]") for n in (1, 2)]
        llenos = []
        for campo, valor in zip(campos, VALORES_LIMITADO):
            if campo.count() and campo.is_visible() and campo.is_editable():
                campo.fill(valor)
                llenos.append(valor)
        cambios.append(f"{titulo}: ilimitado → limitado ({' / '.join(llenos) or 'sin campos'})")
    return cambios


def terminar_guardando(page, boton: str = "Guardar cambios") -> str:
    """'Guardar cambios' → confirmaciones → todos los 'Aceptar' seguidos (SIRT muestra
    dos: '¿Está seguro…?' y 'Registro guardado correctamente')."""
    page.get_by_role("button", name=boton, exact=True).first.click()
    for _ in range(60):                                  # hasta ~15 s
        if page.get_by_role("button", name="Aceptar", exact=True).is_visible():
            mensajes = []
            aceptar = page.get_by_role("button", name="Aceptar", exact=True)
            while len(mensajes) < 5:
                mensajes.append(confirmar_con_aceptar(page))
                try:
                    aceptar.first.wait_for(state="visible", timeout=4_000)
                except Exception:
                    break
            final = " / ".join(mensajes)
            if re.search(r"no se cumplieron|no puede|error", final, re.I) and \
                    not re.search(r"correctamente", final, re.I):
                raise RuntimeError(f"SIRT rechazó el guardado: {final[:400]}")
            return final
        dialogo = page.locator("mat-dialog-container")
        if dialogo.count():
            for texto in ("Guardar tarifa", "Guardar cambios", "Confirmar", "Guardar", "Sí", "Si"):
                boton = dialogo.last.get_by_role("button", name=texto, exact=True)
                if boton.count() and boton.is_visible():
                    boton.click()
                    break
        page.wait_for_timeout(250)
    raise RuntimeError(f"No apareció el mensaje final tras 'Guardar cambios'. Pantalla: {texto_dialogo(page)[:200]}")


# Campos numéricos que se cambian en la rectificación (los demás no se tocan: hay
# reglas entre campos, p. ej. tethering ≤ tope de descarga).
NUMERICOS_PASO_1 = [(re.compile(r"^Tarifa (de )?instalaci[oó]n", re.I), "Tarifa de instalación")]
NUMERICOS_PASO_2 = [(re.compile(r"^Valor del servicio", re.I), "Valor del servicio")]


def esperar_datos_cargados(page) -> None:
    """El botón 'Continuar' aparece antes que los datos: el 15-16/10 el paso 1 se leyó
    vacío (0 textos y 'Tarifa de instalación: (vacío) → 1'). Espera a que haya cajas de
    texto visibles y con contenido, y a que la red se calme (máx. ~10 s, sin fallar)."""
    try:
        page.locator("textarea:visible").first.wait_for(state="visible", timeout=8_000)
        page.wait_for_function(
            """() => [...document.querySelectorAll('textarea')]
                     .filter(t => t.offsetParent !== null)
                     .every(t => t.value.trim() !== '')""", timeout=5_000)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=3_000)
    except Exception:
        pass
    page.wait_for_timeout(300)


def avanzar(page, boton, siguiente, nombre: str, intentos: int = 3) -> None:
    """Pulsa `boton` hasta que aparezca `siguiente`. A veces SIRT ignora el primer clic
    (p. ej. PIF el 16/10: se quedó en el paso 1 sin mostrar ningún error)."""
    for intento in range(1, intentos + 1):
        boton.click()
        try:
            siguiente.wait_for(state="visible", timeout=10_000)
            return
        except Exception:
            aviso = texto_dialogo(page)
            if aviso:
                raise RuntimeError(f"No se pudo pasar al {nombre}: {aviso[:200]}")
            print(f"    {nombre} no cargó (intento {intento}); se vuelve a pulsar")
            page.wait_for_timeout(1_000)
    raise RuntimeError(f"No se pudo pasar al {nombre} tras {intentos} intentos.")


def editar_tres_pasos(page, tarifa: dict, opcion_menu: str, sello: str, prefijo: str,
                      hasta_resumen: bool, delta: Decimal | None = None,
                      cambiar_listas: bool = True,
                      si_vacio: Decimal | None = None,
                      fecha_ejecucion: date | None = None,
                      todo: bool = False) -> tuple[list[str], str]:
    """Flujo común de Fe de erratas y Rectificar: paso 1 → paso 2 → resumen → guardar.
    Textos: se les agrega el sello. Números (solo con `delta`): valor del servicio y
    tarifa de instalación, si están habilitados."""
    fila = fila_en_tabla(page, tarifa)
    abrir_accion(page, fila, opcion_menu)
    cambios, textos, numeros = [], [], []
    inputs_hechos: set = set()

    def corregir_pantalla(numericos, listas=()):
        nonlocal textos
        nuevos = corregir_textos(page, sello, prefijo)
        nuevos += corregir_inputs(page, sello, prefijo, inputs_hechos)
        textos += nuevos
        cambios.append(f"{len(nuevos)} texto(s)")
        if cambiar_listas:
            for etiqueta, nombre in listas:
                hecho = cambiar_lista_multiple(page, etiqueta, nombre)
                if hecho:
                    cambios.append(hecho[0])
                    numeros.extend(hecho[1])             # se buscan también en el resumen
        if delta is not None:
            for etiqueta, nombre in numericos:
                hecho = cambiar_numero(page, etiqueta, nombre, delta, si_vacio)
                if hecho:
                    cambios.append(hecho)
                    numeros.append(hecho.split("→")[-1].strip())
            if todo and listas:                          # solo en el paso 2
                otros = cambiar_todos_los_numeros(page, delta, [e for e, _ in NUMERICOS_PASO_2 + NUMERICOS_PASO_1]
                                                  + NUNCA_EN_TODO)
                cambios.append(f"--todo: {len(otros)} campo(s) numérico(s)")
                cambios.extend(otros)
        if delta is not None and listas and todo:        # filas de tablas (con --todo / --campos)
            for hecho in editar_filas_de_tablas(page, delta):
                cambios.append(hecho)
                print(f"    {hecho}")
        if delta is not None and listas and EXTRAS["agregar_filas"]:
            for hecho in agregar_filas_a_tablas(page):
                cambios.append(hecho)
                print(f"    {hecho}")
        if delta is not None and listas and EXTRAS["quitar_ilimitado"]:
            for hecho in quitar_ilimitado(page):
                cambios.append(hecho)
                print(f"    {hecho}")
        if delta is not None and listas and EXTRAS["canales"]:   # paso 2 de rectificación / cambio
            for hecho in cambiar_canales(page):
                cambios.append(hecho)
                print(f"    {hecho}")
        if todo and cambiar_listas and listas:
            extra = LISTAS_TODO
            if FILTRO_CAMPOS is not None:                # listas elegidas en --campos
                extra = [(re.compile("^" + re.escape(n), re.I), n) for n in NOMBRES_CAMPOS]
            for etiqueta, nombre in extra:
                if any(nombre == n for _, n in listas):
                    continue                             # ya se trató arriba
                hecho = cambiar_lista_multiple(page, etiqueta, nombre)
                if hecho:
                    cambios.append(hecho[0])
                    numeros.extend(hecho[1])

    # Paso 1: datos de la tarifa.
    page.get_by_role("button", name="Continuar con el paso 2").wait_for(state="visible")
    esperar_datos_cargados(page)
    if fecha_ejecucion:                                   # solo Cambio de Tarifa
        elegida = elegir_fecha_calendario(page, fecha_ejecucion, DIAS_EXTRA_SI_BLOQUEADO)
        cambios.append(f"Fecha de ejecución: {elegida:%d/%m/%Y}"
                       + (f" (pedida {fecha_ejecucion:%d/%m/%Y}, estaba bloqueada)" if elegida != fecha_ejecucion else ""))
    corregir_pantalla(NUMERICOS_PASO_1)
    if opcion_menu == MENUS["rectificar"] and delta is not None:   # Licitación y Contratos
        extra_cambios, extra_valores = rectificar_licitacion_paso_1(page, delta)
        cambios.extend(extra_cambios)
        numeros.extend(extra_valores)
    continuar3 = page.get_by_role("button", name="Continuar con el paso 3")
    avanzar(page, page.get_by_role("button", name="Continuar con el paso 2"), continuar3, "paso 2")

    # Paso 2: datos del servicio (si hay pestañas "Servicio N", se corrige cada una).
    esperar_datos_cargados(page)
    pestanas = page.get_by_role("tab", name=re.compile(r"^Servicio \d+$"))
    total = pestanas.count()
    for i in range(max(1, total)):
        if total:
            pestanas.nth(i).click()
            page.wait_for_timeout(500)
        numericos = NUMERICOS_PASO_2
        if total > 1 and delta is not None and ajuste("Valor del servicio", delta) == ("delta", abs(delta)) \
                and abs(delta) > 0 and "empaquet" in tarifa.get("condicion", "").lower():
            # Empaquetado: cada servicio debe ser <= el valor de SU tarifa base y el plan (suma)
            # MENOR a la suma de las bases. Un paquete creado con «base − 1» (o con una base que
            # vale 0) no admite subidas: SIRT lo rechazó el 16/10 (suma) y el 22/10 con
            # TE-EMP-CM-113 / 126 (servicio > su base). Por eso en empaquetados el valor de los
            # servicios NO sube (sí se cambian textos, instalación y lo demás).
            numericos = []
            cambios.append(f"Servicio {i + 1}: valor sin subir (regla del empaquetado: cada servicio ≤ su "
                           "tarifa base y plan < suma de bases)")
        corregir_pantalla(numericos, LISTAS_PASO_2)
    if not textos and not numeros and not any("→" in c for c in cambios):
        raise RuntimeError("No se encontró ningún campo para modificar"
                           + (" entre los elegidos en --campos." if FILTRO_CAMPOS is not None else "."))
    avanzar(page, continuar3, page.get_by_role("button", name="Guardar cambios"), "paso 3 (resumen)")

    # Paso 3: el resumen debe mostrar los textos y números nuevos.
    page.get_by_role("button", name="Guardar cambios").wait_for(state="visible")
    resumen = _sin_espacios(page.locator("body").inner_text())
    # Un valor puede ser una tupla de formatos equivalentes (fechas: 29/11/2026 o 2026-11-29).
    faltan = [t for t in textos + numeros
              if not any(_sin_espacios(x) in resumen for x in (t if isinstance(t, tuple) else (t,)))]
    if faltan:
        primero = faltan[0] if isinstance(faltan[0], str) else " o ".join(faltan[0])
        raise RuntimeError(f"El resumen no muestra {len(faltan)} de los cambios: {primero[:80]}")
    if hasta_resumen:
        return cambios, "detenido en el resumen (sin guardar)"
    return cambios, terminar_guardando(page)


# --- Principal -------------------------------------------------------------------

def leer_argumentos():
    parser = argparse.ArgumentParser(description="Acciones sobre tarifas registradas.")
    parser.add_argument("accion", choices=["buscar", "renombre", "cese", "fe_erratas", "rectificar", "cambio", "renovar"])
    parser.add_argument("--tarifas", nargs="+", required=True, help="nombres actuales de las tarifas")
    parser.add_argument("--sufijo", default=SUFIJO_POR_DEFECTO, help="renombre: se agrega al nombre actual")
    parser.add_argument("--nuevo-nombre", help="renombre: nombre exacto (solo con una tarifa)")
    parser.add_argument("--fecha", help="cese: fin de comercialización AAAA-MM-DD (por defecto, mañana); "
                                        "cambio: fecha de ejecución (por defecto según --delta); "
                                        "renovar: inicio de la comercialización (por defecto, el día por defecto del calendario + 1)")
    parser.add_argument("--sin-validar", action="store_true",
                        help="cese / fe_erratas: intentar aunque la tarifa no esté En comercialización")
    parser.add_argument("--texto", default=None,
                        help="fe_erratas / rectificar: texto del sello (por defecto 'Fe de erratas' / 'Rectificación')")
    parser.add_argument("--delta", default=None,
                        help="rectificar / cambio: cuánto se suma al valor del servicio y a la tarifa de "
                             "instalación (por defecto: rectificar 1, cambio -1)")
    parser.add_argument("--valor-si-vacio", default=None,
                        help="rectificar: valor para un campo numérico vacío (por defecto, el delta: 1)")
    parser.add_argument("--hasta-resumen", action="store_true",
                        help="fe_erratas / rectificar / cambio / renovar: llega al paso 3 y se detiene SIN guardar")
    parser.add_argument("--sin-listas", action="store_true",
                        help="fe_erratas / rectificar: no cambiar listas como 'Tipo de Contratación'")
    parser.add_argument("--todo", action="store_true",
                        help="rectificar / cambio: cambia también TODOS los demás campos numéricos del "
                             "paso 2 (datos, velocidades, minutos, tarifas...), 'Apps sin consumo de datos' y "
                             "los números de cada fila de las tablas (lápiz → Guardar).")
    parser.add_argument("--dias-renovacion", type=int, default=None,
                        help="renovar: días de comercialización (fin = inicio + N; por defecto 19)")
    parser.add_argument("--duracion-renovacion", default=None,
                        help="renovar: duración (días, máx. 180) para todas las duraciones; por defecto 25 (o 31 si ya era 25)")
    parser.add_argument("--agregar-filas", action="store_true",
                        help="rectificar / cambio: agrega una fila a cada tabla del paso 2 (país, tecnología "
                             "o plataforma que no esté ya en la tabla)")
    parser.add_argument("--quitar-ilimitado", action="store_true",
                        help="rectificar / cambio: desmarca «ilimitado» en los condicionales y llena "
                             "cantidad y tarifa adicional (10 y 15)")
    parser.add_argument("--sentido", choices=["beneficioso", "perjudicial"], default=None,
                        help="cambio: beneficioso = precios bajan y beneficios suben; perjudicial = al revés "
                             "(usa el valor absoluto de --delta)")
    parser.add_argument("--detalle", default=None,
                        help="JSON {campo: {\"valor\": X} o {\"delta\": N}}: solo esos campos, con su valor o delta")
    parser.add_argument("--sin-canales", action="store_true",
                        help="rectificar / cambio (TV): no cambiar los Excel de canales")
    parser.add_argument("--campos", default=None,
                        help="fe_erratas / rectificar / cambio: solo estos campos, separados por ';' "
                             "(p. ej. 'Otros datos generales;Valor del servicio;Datos para tethering')")
    parser.add_argument("--repetir", action="store_true",
                        help="hacer la acción aunque la tarifa ya la tenga EXITOSA en el CSV")
    parser.add_argument("--si", action="store_true", help="no pedir confirmación")
    parser.add_argument("--solo-listar", action="store_true", help="muestra lo que haría, sin tocar SIRT")
    parser.add_argument("--diagnostico", action="store_true")
    parser.add_argument("--al-frente", action="store_true")
    return parser.parse_args()


MENUS = {"fe_erratas": "Fe de Erratas", "rectificar": "Rectificar", "cambio": "Cambio de Tarifa"}
PREFIJOS = {"fe_erratas": "Fe de erratas", "rectificar": "Rectificación", "cambio": "Cambio"}
NOMBRES_ACCION = {"fe_erratas": "la fe de erratas", "rectificar": "la rectificación",
                  "cambio": "el cambio de tarifa"}


def nombre_nuevo_de(nombre: str, args) -> str:
    return args.nuevo_nombre if args.nuevo_nombre else f"{nombre}{args.sufijo}"


NOMBRES_CAMPOS: list[str] = []
EXTRAS = {"agregar_filas": False, "quitar_ilimitado": False, "canales": True}   # se fijan en main()


def main() -> None:
    global FILTRO_CAMPOS, NOMBRES_CAMPOS
    args = leer_argumentos()
    global DIAS_COMERCIALIZACION_RENOVACION, DURACION_FIJA
    if args.dias_renovacion:
        DIAS_COMERCIALIZACION_RENOVACION = args.dias_renovacion
    if args.duracion_renovacion:
        if not (args.duracion_renovacion.isdigit() and 1 <= int(args.duracion_renovacion) <= 180):
            sys.exit("--duracion-renovacion debe ser un número de días entre 1 y 180.")
        DURACION_FIJA = args.duracion_renovacion
    EXTRAS["agregar_filas"] = args.agregar_filas
    EXTRAS["quitar_ilimitado"] = args.quitar_ilimitado
    EXTRAS["canales"] = not args.sin_canales
    global SENTIDO
    SENTIDO = args.sentido
    if args.detalle:
        crudo = json.loads(Path(args.detalle).read_text(encoding="utf-8"))
        cargar_detalle(crudo)
        if not args.campos:                              # se recorren los bloques / tablas nombrados
            args.campos = ";".join(dict.fromkeys(SEPARADOR_RUTA.split(k)[0].strip() for k in crudo))
    if args.campos:
        NOMBRES_CAMPOS = [c.strip() for c in args.campos.split(";") if c.strip()]
        FILTRO_CAMPOS = [normalizar(c) for c in NOMBRES_CAMPOS]
        args.todo = True                                 # los números elegidos se buscan en todo el paso 2
        print(f"  Campos elegidos ({len(NOMBRES_CAMPOS)}): {'; '.join(NOMBRES_CAMPOS)}")
    if args.nuevo_nombre and len(args.tarifas) != 1:
        sys.exit("--nuevo-nombre solo se puede usar con una tarifa.")

    hechas = ya_hechas(args.accion) if args.accion not in ("buscar",) and not args.repetir else set()
    pendientes = [t for t in args.tarifas if t not in hechas]
    for t in args.tarifas:
        if t in hechas:
            print(f"  {t}: ya tiene {args.accion} EXITOSO en {RUTA_RESULTADOS} (se omite)")

    if args.accion == "renombre":
        for t in pendientes:
            print(f"  {t}  →  {nombre_nuevo_de(t, args)}")
    fecha_cese = date.fromisoformat(args.fecha) if args.fecha else date.today() + timedelta(days=1)
    if args.accion == "cese":
        if fecha_cese <= date.today():
            sys.exit("La fecha de cese debe ser posterior a hoy.")
        for t in pendientes:
            print(f"  {t}  →  fin de comercialización {fecha_cese:%d/%m/%Y}")
    if args.delta is None:
        args.delta = "-1" if args.accion == "cambio" else "1"
    fecha_ejecucion = None
    if args.accion == "cambio":
        alza = Decimal(args.delta) > 0
        minimo = date.today() + timedelta(days=DIAS_ANTICIPACION_ALZA if alza else DIAS_ANTICIPACION_BAJA)
        fecha_ejecucion = date.fromisoformat(args.fecha) if args.fecha else minimo
        if fecha_ejecucion < minimo:
            print(f"  AVISO: fecha {fecha_ejecucion:%d/%m/%Y} antes del mínimo {minimo:%d/%m/%Y} "
                  f"({'alza' if alza else 'baja'}); SIRT podría rechazarla.")
    if args.accion in ("fe_erratas", "rectificar", "cambio"):
        prefijo = args.texto or PREFIJOS[args.accion]
        extra = "" if args.accion == "fe_erratas" else \
            (f"; {SENTIDO}: precios {'−' if SENTIDO == 'beneficioso' else '+'}{abs(Decimal(args.delta))}, "
             f"beneficios {'+' if SENTIDO == 'beneficioso' else '−'}{abs(Decimal(args.delta))}" if SENTIDO else
             f"; valor del servicio y tarifa de instalación {Decimal(args.delta):+}")
        if DETALLE:   # en modo detalle solo cambia lo de la hoja «Campos»
            extra = f"; detalle: solo {len(DETALLE)} campo(s), cada uno con su valor/delta"
        if fecha_ejecucion:
            extra += f"; fecha de ejecución {fecha_ejecucion:%d/%m/%Y}"
        for t in pendientes:
            print(f"  {t}  →  {args.accion}: textos + '| {prefijo} <fecha hora>'{extra}"
                  + ("  (se detiene en el paso 3, sin guardar)" if args.hasta_resumen else ""))
    inicio_renovacion = date.fromisoformat(args.fecha) if args.fecha else None
    if args.accion == "renovar":
        desde = f"{inicio_renovacion:%d/%m/%Y}" if inicio_renovacion else \
            f"fecha por defecto del calendario + {DIAS_DESPUES_DEL_DEFECTO} día"
        for t in pendientes:
            print(f"  {t}  →  renovar: comercialización desde {desde} "
                  f"(fin = inicio + {DIAS_COMERCIALIZACION_RENOVACION} días); duraciones "
                  + (DURACION_FIJA if DURACION_FIJA else "25 (o 31 si ya eran 25)")
                  + ("  (se detiene en el paso 3, sin guardar)" if args.hasta_resumen else ""))
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT.")
        return
    if not pendientes:
        print("No hay tarifas pendientes.")
        return
    if args.accion in ("cese", "fe_erratas", "rectificar", "cambio", "renovar") and not args.si \
            and not args.hasta_resumen:
        accion_txt = {"cese": "el CESE", "fe_erratas": "una FE DE ERRATAS",
                      "rectificar": "una RECTIFICACIÓN", "cambio": "un CAMBIO DE TARIFA",
                      "renovar": "una RENOVACIÓN"}[args.accion]
        print(f"\nATENCIÓN: se registrará {accion_txt} en {len(pendientes)} tarifa(s) (probablemente irreversible).")
        if input("Escribe 'si' para continuar: ").strip().lower() not in ("si", "sí"):
            print("Cancelado.")
            return

    from playwright.sync_api import sync_playwright

    diag = Diagnostico(args.diagnostico, False, RUTA_RESULTADOS.parent / "diagnostico")
    exitosas = errores = 0

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        diag.iniciar(context)
        if args.al_frente:
            page.bring_to_front()

        for numero, nombre in enumerate(pendientes, start=1):
            print("\n" + "=" * 70)
            print(f"[{numero}/{len(pendientes)}] {nombre}")
            estado = {"etapa": ""}
            fila = {"fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "accion": args.accion, "nombre": nombre}
            diag.iniciar_caso({"correlativo": numero, "nombre": nombre})

            def etapa(nombre_etapa, accion):
                estado["etapa"] = nombre_etapa
                resultado = accion()
                diag.captura_paso(page, nombre_etapa)
                return resultado

            try:
                etapa("abrir_tarifas_registradas", lambda: abrir_tarifas_registradas(page))
                tarifa = etapa("buscar", lambda: buscar_tarifa(page, nombre))
                if tarifa is None:
                    raise RuntimeError(f"'{nombre}' no aparece en Tarifas registradas.")
                fila.update(codigo_sirt=tarifa["codigo"], version=tarifa["version"],
                            comercializacion=tarifa["comercializacion"], vigencia=tarifa["vigencia"])
                print(f"  {tarifa['codigo']} v{tarifa['version']} | {tarifa['tipo']} {tarifa['condicion']}"
                      f" | {tarifa['comercializacion']} | {tarifa['vigencia']}")

                if args.accion == "buscar":
                    fila["resultado"] = "ENCONTRADA"
                elif args.accion == "renombre":
                    nuevo = nombre_nuevo_de(nombre, args)
                    fila["nombre_nuevo"] = nuevo
                    mensaje = etapa("renombre", lambda: renombrar(page, tarifa, nuevo))
                    fila["resultado"] = "EXITOSO"
                    fila["detalle"] = mensaje[:300]
                    print(f"  Renombrada: {nuevo}")
                elif args.accion == "cese":
                    if re.search(r"promocional", tarifa.get("tipo", ""), re.I):
                        raise NoAplica("una tarifa PROMOCIONAL no se puede cesar (solo monoproductos y empaquetados).")
                    validar_para_cese(tarifa, args.sin_validar)
                    mensaje = etapa("cese", lambda: cesar(page, tarifa, fecha_cese))
                    fila["resultado"] = "EXITOSO"
                    fila["detalle"] = f"fin comercialización {fecha_cese:%d/%m/%Y} | {mensaje[:250]}"
                    print(f"  Cese programado: fin de comercialización {fecha_cese:%d/%m/%Y}")
                elif args.accion in ("fe_erratas", "rectificar", "cambio"):
                    con_numeros = args.accion != "fe_erratas"
                    validar_para_cese(tarifa, args.sin_validar, NOMBRES_ACCION[args.accion])
                    prefijo = args.texto or PREFIJOS[args.accion]
                    sello = f"| {prefijo} {datetime.now():%d/%m/%Y %H:%M}"
                    cambios, mensaje = etapa(args.accion, lambda: editar_tres_pasos(
                        page, tarifa, MENUS[args.accion], sello, prefijo,
                        args.hasta_resumen, Decimal(args.delta) if con_numeros else None,
                        not args.sin_listas,
                        Decimal(args.valor_si_vacio) if args.valor_si_vacio else None,
                        fecha_ejecucion, args.todo and con_numeros))
                    fila["resultado"] = "EN_RESUMEN" if args.hasta_resumen else "EXITOSO"
                    fila["detalle"] = f"{' | '.join(cambios)} | {sello} | {mensaje[:250]}"
                    for c in cambios:
                        print(f"  {c}")
                    print("  Detenido en el paso 3 SIN guardar." if args.hasta_resumen else "  Guardado.")
                elif args.accion == "renovar" and _renovacion_ya_guardada(nombre, tarifa):
                    # La corrida anterior falló JUSTO después de guardar (p. ej. sesión vencida,
                    # 22/10 TP-PIF-P5-106): la renovación ya creó su versión «Por iniciar».
                    fila["resultado"] = "EXITOSO"
                    fila["detalle"] = (f"renovación ya registrada: SIRT tiene la versión {tarifa['version']} "
                                       "«Por iniciar comercialización» (el intento anterior se cortó después de guardar)")
                    print("  " + fila["detalle"])
                elif args.accion == "renovar":
                    validar_para_cese(tarifa, args.sin_validar, "la renovación")
                    cambios, mensaje = etapa("renovar", lambda: renovar(
                        page, tarifa, args.hasta_resumen, inicio_renovacion))
                    fila["resultado"] = "EN_RESUMEN" if args.hasta_resumen else "EXITOSO"
                    fila["detalle"] = f"{' | '.join(cambios)} | {mensaje[:250]}"
                    for c in cambios:
                        print(f"  {c}")
                    print("  Detenido en el paso 3 SIN guardar." if args.hasta_resumen else "  Guardado.")
                exitosas += 1

            except NoAplica as aviso:
                fila["resultado"] = "NO_APLICA"
                fila["detalle"] = str(aviso)[:500]
                print(f"  NO APLICA (no se tocó SIRT): {aviso}")

            except Exception as error:
                diag.registrar_fallo(page, error, estado["etapa"])
                fila["resultado"] = "ERROR"
                fila["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
                errores += 1
                print(f"  ERROR en '{estado['etapa']}': {error}")
                page.keyboard.press("Escape")

            diag.cerrar_caso(fila["resultado"] == "ERROR")
            if args.accion != "buscar":
                registrar(fila)

        diag.finalizar()

    print("\n" + "=" * 70)
    print(f"Correctas: {exitosas} | Errores: {errores}")
    if args.accion != "buscar":
        print(f"Resultados: {RUTA_RESULTADOS.resolve()}")


if __name__ == "__main__":
    main()
