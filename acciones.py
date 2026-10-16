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
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import comun
from empaquetar import texto_dialogo
from ejecutar import Diagnostico


RUTA_RESULTADOS = comun.DIRECTORIO_RESULTADOS / "ACCIONES" / "acciones.csv"
CAMPOS = ["fecha_hora", "accion", "nombre", "codigo_sirt", "version", "nombre_nuevo",
          "comercializacion", "vigencia", "resultado", "detalle"]
TIMEOUT_BUSQUEDA = 20_000
SUFIJO_POR_DEFECTO = "_REN"

# --- Cambio de tarifa: anticipación de la "Fecha de ejecución" (días desde hoy) ------
# SIRT (16/10, al SUBIR valor del servicio e instalación con fecha de hoy): "La fecha de
# ejecución de la acción no puede ser anterior a 2 días a partir de la fecha actual,
# según los tipos de variación aplicados". Si OSIPTEL cambia el plazo, se cambia aquí.
DIAS_ANTICIPACION_ALZA = 2      # cambio perjudicial para el usuario (sube precios)
DIAS_ANTICIPACION_BAJA = 1      # cambio beneficioso (baja precios): el calendario siempre
                                # bloquea HOY (visto 16/10), lo antes posible es mañana
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
                    break
                respuestas.remove(respuesta)
            if items is not None:
                break
            page.wait_for_timeout(250)
    finally:
        page.remove_listener("response", capturar)

    if items is None:
        raise RuntimeError(f"No llegó la respuesta del filtro al buscar '{nombre}'.")

    exactos = [i for i in items if _sin_espacios(i.get("nombre")) == _sin_espacios(nombre)]
    if not exactos:
        return None
    ultimo = max(exactos, key=lambda i: (bool(i.get("esUltimaVersion")), i.get("version") or 0))
    page.wait_for_timeout(300)        # que la tabla termine de pintarse
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

class NoAplica(Exception):
    """La tarifa no está en el estado que exige la acción: no se tocó SIRT."""


def validar_para_cese(tarifa: dict, sin_validar: bool, accion: str = "el cese") -> None:
    """SIRT da error si la tarifa no está En comercialización y vigente."""
    problemas = []
    if "en comercializ" not in tarifa["comercializacion"].lower():
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
    elegir_fecha_calendario(page, fecha)
    page.get_by_role("button", name="Guardar", exact=True).click()
    page.get_by_role("button", name="Guardar tarifa", exact=True).click()
    return confirmar_con_aceptar(page)


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
  let n = 0;
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
    if (ok) { t.setAttribute('data-fe', String(n++)); }
  }
  return n;
}"""


def corregir_textos(page, sello: str, prefijo: str) -> list[str]:
    """Agrega el sello a los campos de texto (textarea) visibles y editables de la
    pestaña activa. Sirve para cualquier servicio: solo cambia cuántos campos hay."""
    corregidos = []
    total = page.evaluate(MARCAR_TEXTOS)
    for i in range(total):
        campo = page.locator(f'textarea[data-fe="{i}"]')
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


def _numero_nuevo(actual: str, delta: Decimal) -> str:
    """'50' + 1 → '51'; '0.5' + 1 → '1.5' (mantiene los decimales que tenía)."""
    from decimal import Decimal as D
    texto = (actual or "").strip().replace(",", "")
    valor = D(texto) if texto else D(0)
    decimales = len(texto.split(".")[1]) if "." in texto else 0
    return f"{valor + delta:.{decimales}f}"


def cambiar_numero(page, etiqueta: re.Pattern, nombre: str, delta: Decimal,
                   si_vacio: Decimal | None = None) -> str | None:
    """Cambia el campo numérico visible y habilitado cuya etiqueta coincide:
    - con número: le suma `delta` (50 → 51);
    - vacío de verdad (tras esperar hasta ~5 s a que SIRT lo llene): pone `si_vacio`
      (por defecto 0 + delta, o sea 1).
    Devuelve 'Etiqueta: antes → después' o None si el campo no existe o está bloqueado."""
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
            if antes.strip():
                despues = _numero_nuevo(antes, delta)
                if Decimal(despues) < 0:
                    return None                          # no se puede bajar de 0
            elif si_vacio is not None:
                despues = str(si_vacio)
            elif delta > 0:
                despues = str(delta)
            else:
                return None                              # vacío y se quería bajar: no se toca
            campo.fill(despues)
            return f"{nombre}: {antes.strip() or '(vacío)'} → {despues}"
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
    }
    if (!ok) continue;
    const v = (i.value || '').trim().replace(/,/g, '');
    if (!/^-?\\d+(\\.\\d+)?$/.test(v)) continue;
    const ff = i.closest('mat-form-field');
    const et = ff && ff.querySelector('mat-label, label');
    const lab = ((et && et.innerText) || i.getAttribute('aria-label') || i.placeholder || '').trim();
    i.setAttribute('data-num', String(n));
    out.push({n: n++, lab: lab.replace(/\\s+/g, ' '), v});
  }
  return out;
}"""
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
    campos = [c for c in campos if not any(e.search(c["lab"]) for e in excluir)]
    campos.sort(key=lambda c: (not TOPES.search(c["lab"])) if delta > 0 else bool(TOPES.search(c["lab"])))
    hechos = []
    for c in campos:
        if Decimal(c["v"]) == 0:
            continue                                     # 0 suele ser "no aplica": no se toca
        decimales = len(c["v"].split(".")[1]) if "." in c["v"] else 0
        paso = delta * (Decimal(1).scaleb(-decimales))   # 0.10 → 0.11 (no 1.10); 12 → 13
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
    for combo in page.get_by_role("combobox", name=etiqueta).all():
        if not (combo.is_visible() and combo.is_enabled()):
            continue
        antes = " ".join(combo.inner_text().split()) or "(vacío)"
        combo.click()
        lista = page.get_by_role("listbox").last
        lista.wait_for(state="visible")
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


def terminar_guardando(page) -> str:
    """'Guardar cambios' → confirmaciones → todos los 'Aceptar' seguidos (SIRT muestra
    dos: '¿Está seguro…?' y 'Registro guardado correctamente')."""
    page.get_by_role("button", name="Guardar cambios").click()
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

    def corregir_pantalla(numericos, listas=()):
        nonlocal textos
        nuevos = corregir_textos(page, sello, prefijo)
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
                otros = cambiar_todos_los_numeros(page, delta, [e for e, _ in NUMERICOS_PASO_2 + NUMERICOS_PASO_1])
                cambios.append(f"--todo: {len(otros)} campo(s) numérico(s)")
                cambios.extend(otros)
        if todo and cambiar_listas and listas:
            for etiqueta, nombre in LISTAS_TODO:
                hecho = cambiar_lista_multiple(page, etiqueta, nombre)
                if hecho:
                    cambios.append(hecho[0])

    # Paso 1: datos de la tarifa.
    page.get_by_role("button", name="Continuar con el paso 2").wait_for(state="visible")
    esperar_datos_cargados(page)
    if fecha_ejecucion:                                   # solo Cambio de Tarifa
        elegida = elegir_fecha_calendario(page, fecha_ejecucion, DIAS_EXTRA_SI_BLOQUEADO)
        cambios.append(f"Fecha de ejecución: {elegida:%d/%m/%Y}"
                       + (f" (pedida {fecha_ejecucion:%d/%m/%Y}, estaba bloqueada)" if elegida != fecha_ejecucion else ""))
    corregir_pantalla(NUMERICOS_PASO_1)
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
        if total > 1 and delta is not None and delta > 0 and i == total - 1:
            # Empaquetado: el valor del plan (suma de los servicios) debe ser MENOR a la suma
            # de las tarifas base, y cada servicio <= su tarifa base. Como el paquete se creó
            # con base - 1, subir todos daría suma = suma de bases → SIRT lo rechaza (16/10).
            # Se deja el último servicio sin subir.
            numericos = []
            cambios.append(f"Servicio {i + 1}: valor sin subir (regla del empaquetado: plan < suma de bases)")
        corregir_pantalla(numericos, LISTAS_PASO_2)
    if not textos and not numeros:
        raise RuntimeError("No se encontró ningún campo para modificar.")
    avanzar(page, continuar3, page.get_by_role("button", name="Guardar cambios"), "paso 3 (resumen)")

    # Paso 3: el resumen debe mostrar los textos y números nuevos.
    page.get_by_role("button", name="Guardar cambios").wait_for(state="visible")
    resumen = _sin_espacios(page.locator("body").inner_text())
    faltan = [t for t in textos + numeros if _sin_espacios(t) not in resumen]
    if faltan:
        raise RuntimeError(f"El resumen no muestra {len(faltan)} de los cambios: {faltan[0][:80]}")
    if hasta_resumen:
        return cambios, "detenido en el resumen (sin guardar)"
    return cambios, terminar_guardando(page)


# --- Principal -------------------------------------------------------------------

def leer_argumentos():
    parser = argparse.ArgumentParser(description="Acciones sobre tarifas registradas.")
    parser.add_argument("accion", choices=["buscar", "renombre", "cese", "fe_erratas", "rectificar", "cambio"])
    parser.add_argument("--tarifas", nargs="+", required=True, help="nombres actuales de las tarifas")
    parser.add_argument("--sufijo", default=SUFIJO_POR_DEFECTO, help="renombre: se agrega al nombre actual")
    parser.add_argument("--nuevo-nombre", help="renombre: nombre exacto (solo con una tarifa)")
    parser.add_argument("--fecha", help="cese: fin de comercialización AAAA-MM-DD (por defecto, mañana); "
                                        "cambio: fecha de ejecución (por defecto según --delta)")
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
                        help="fe_erratas / rectificar: llega al paso 3 y se detiene SIN guardar")
    parser.add_argument("--sin-listas", action="store_true",
                        help="fe_erratas / rectificar: no cambiar listas como 'Tipo de Contratación'")
    parser.add_argument("--todo", action="store_true",
                        help="rectificar / cambio: cambia también TODOS los demás campos numéricos del "
                             "paso 2 (datos, velocidades, minutos, tarifas...) y 'Apps sin consumo de datos'. "
                             "No toca filas de tablas.")
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


def main() -> None:
    args = leer_argumentos()
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
            f"; valor del servicio y tarifa de instalación {Decimal(args.delta):+}"
        if fecha_ejecucion:
            extra += f"; fecha de ejecución {fecha_ejecucion:%d/%m/%Y}"
        for t in pendientes:
            print(f"  {t}  →  {args.accion}: textos + '| {prefijo} <fecha hora>'{extra}"
                  + ("  (se detiene en el paso 3, sin guardar)" if args.hasta_resumen else ""))
    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT.")
        return
    if not pendientes:
        print("No hay tarifas pendientes.")
        return
    if args.accion in ("cese", "fe_erratas", "rectificar", "cambio") and not args.si and not args.hasta_resumen:
        accion_txt = {"cese": "el CESE", "fe_erratas": "una FE DE ERRATAS",
                      "rectificar": "una RECTIFICACIÓN", "cambio": "un CAMBIO DE TARIFA"}[args.accion]
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
