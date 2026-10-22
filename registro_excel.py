"""
Registra TARIFAS ESTABLECIDAS (monoproducto) según automatizacion/registro_tarifas_establecidas.xlsx.

Cada hoja de servicio (TF, IF, SMP…) = una fila por tarifa. Solo se registran las hojas que
estén en el archivo. Las tablas (países, velocidades móviles, bolsas) se toman de las hojas
«Tablas – …», unidas por el nombre de la tarifa.

Qué hace con cada fila (Ejecutar = Sí o vacío):
1. Valida los datos (nombre, alcance, fechas, números, condicionales, archivos de canales).
2. Registra la tarifa en SIRT con el mismo código que ejecutar.py (paso 1, paso 2, confirmar),
   pero con los valores de la fila en lugar de los fijos del código.
   Lo que el Excel deja vacío y SIRT exige se completa con el valor de siempre del código
   (se avisa en «Detalle»).
3. Si «Registro definitivo» = Sí: hace el registro definitivo (con la disponibilidad comercial
   de la fila si el servicio la pide; vacío = Lima).
4. Escribe Estado, Código SIRT (n.º de registro), Fecha y Detalle en la fila, y guarda una copia
   del Excel en resultados/REGISTRO_EXCEL/.

Las tarifas se anotan en el lote EXCEL (resultados/<código>/resultados_EXCEL.csv): una fila ya
registrada no se repite (estado «Ya registrada»), salvo con --repetir.

Uso (desde la carpeta del proyecto, Edge abierto y sesión de SIRT iniciada en «Tarifas registradas»):
    python registro_excel.py --solo-listar                 muestra lo que haría, sin tocar SIRT
    python registro_excel.py --hojas IF --limite 1         solo la 1.ª fila pendiente de la hoja IF
    python registro_excel.py --nombres TE-IF-CM-001 TE-TF-CM-001
    python registro_excel.py --sin-definitivo              registra pero no hace el registro definitivo
    python registro_excel.py --si --diagnostico

El Excel debe estar CERRADO mientras corre.
"""

import argparse
import copy
import importlib
import sys
from datetime import date, datetime
from pathlib import Path

import excel_comun as xc
from excel_comun import DatoInvalido, texto

EXCEL = xc.AUTOMATIZACION / "registro_tarifas_establecidas.xlsx"
CANALES = xc.AUTOMATIZACION / "canales"
HOJAS_TABLAS = {"paises": "Tablas – Países", "velocidades": "Tablas – Velocidades", "bolsas": "Tablas – Bolsas"}
DISPONIBILIDAD_POR_DEFECTO = "Lima"

# Columna del Excel → lista del código (DIMENSIONES del servicio).
DIMENSIONES = {
    "Alcance de la tarifa": "alcance",
    "Periodicidad": "periodicidad",
    "Tipo de acceso": "tipo_acceso", "Tipo de Acceso (IF)": "tipo_acceso", "Tipo de Acceso (TV)": "tipo_acceso",
    "Tecnologia": "tecnologia", "Tecnología (IF)": "tecnologia_if",
    "Internet Ilimitado": "internet_ilimitado",
    "Tipo de moneda": "moneda",
    "Tipo de línea": "tipo_linea", "Tipo de línea móvil": "tipo_linea_movil",
    "Tipo de Contratación": "tipo_contratacion",
    "Componentes del plan": "componentes",
}
# Columna → valor simple del código (VALORES del servicio): (clave, tipo).
SIMPLES = {
    "Valor del servicio": ("valor_servicio", "numero"),
    "Duración": ("duracion_dias", "entero"),
    "Cantidad de datos": ("cantidad_datos_gb", "numero"),
    "Repetidores y/o equipos similares": ("repetidores", "entero"),
    "Decodificadores o similares": ("decodificadores", "entero"),
    "Tarifa por minuto": ("tarifa_por_minuto", "numero"),
    "Periodo de contrato (meses)": ("periodo_contrato_meses", "entero"),
    "Tecnología de transmisión": ("tecnologia_transmision", "texto"),
    "Velocidad de transmisión": ("velocidad_transmision", "numero"),
    "Vel. transmisión garant.": ("velocidad_garantizada", "numero"),
    "Datos para tethering": ("datos_tethering_gb", "numero"),
    "Traslada datos": ("traslada_datos_gb", "numero"),
    "Datos tope descarga": ("datos_tope_descarga_gb", "numero"),
    "Tarifa por MB adicional": ("tarifa_mb_adicional", "numero"),
    "Tarifa minuto móvil a fijo": ("tarifa_movil_a_fijo", "numero"),
    "Tarifa minuto móvil todo destino": ("tarifa_movil_todo_destino", "numero"),
    "Datos Roaming incluidos en los datos del plan": ("datos_roaming_incluidos_gb", "numero"),
    "Datos Roaming adicionales a los datos del plan": ("datos_roaming_adicionales_gb", "numero"),
}
# Textos: en fijos van a claves propias; en móviles, al diccionario «textos» con el mismo título.
TEXTOS_FIJOS = {
    "Condiciones generales": "condiciones_generales",
    "Restricciones técnicas": "restricciones_tecnicas",
    "Otras condiciones tarifa": "otras_condiciones",
    "Otras restricciones": "otras_restricciones",
}
PASO_1 = ("Inicio de vigencia", "Inicio comercialización", "Tarifa instalación", "Otros datos generales")
CONTACTO = {  # se llenan solo si la fila los trae (SIRT los precarga con los datos del operador)
    "Enlace Web": "enlace", "Fono Atención al cliente": "fono", "Email de atención al cliente": "email",
}
META = {"N°", "Ejecutar", "Nombre de la tarifa", "Registro definitivo", "Disponibilidad comercial",
        "Estado", "Código SIRT", "Fecha de registro", "Detalle", "ID de registro",
        "Qué se prueba", "Resultado esperado"}           # las dos últimas: Excel de prueba funcional
VELOCIDAD_IF = {"Tecnología": "tecnologia",
                "Descarga contratada (Mbps)": "descarga_contratada",
                "Descarga garantizada (Mbps)": "descarga_garantizada",
                "Carga contratada (Mbps)": "carga_contratada",
                "Carga garantizada (Mbps)": "carga_garantizada"}
VELOCIDAD_TABLA = {"Tecnología / Tipo": "tecnologia",
                   "Descarga máx. contratada (Mbps)": "descarga_contratada",
                   "Descarga mín. garantizada (Mbps)": "descarga_garantizada",
                   "Carga máx. contratada (Mbps)": "carga_contratada",
                   "Carga mín. garantizada (Mbps)": "carga_garantizada"}
LDI_FIJO = {"Tarifa por minuto a fijo LDI": "ldi_fijo", "Tarifa por minuto a móvil LDI": "ldi_movil"}


def leer_argumentos():
    p = argparse.ArgumentParser(description="Registro de tarifas establecidas desde el Excel.")
    p.add_argument("--excel", default=str(EXCEL))
    p.add_argument("--hojas", nargs="+", help="solo estas hojas (códigos: IF TF SMP…)")
    p.add_argument("--nombres", nargs="+", help="solo estas tarifas")
    p.add_argument("--limite", type=int, default=0, help="como máximo N tarifas en esta corrida")
    p.add_argument("--repetir", action="store_true", help="registra aunque ya figure como registrada")
    p.add_argument("--sin-definitivo", action="store_true", help="no hace el registro definitivo")
    p.add_argument("--definitivo-en-bloque", action="store_true",
                   help="registro definitivo de varias tarifas a la vez (lo que falle sigue una por una)")
    p.add_argument("--tamano-bloque", type=int, default=10)
    p.add_argument("--solo-listar", action="store_true")
    p.add_argument("--si", action="store_true")
    p.add_argument("--diagnostico", action="store_true")
    p.add_argument("--al-frente", action="store_true")
    return p.parse_args()


def modulos_por_codigo() -> dict:
    sys.path.insert(0, str(xc.PROYECTO))
    modulos = {}
    for archivo in sorted((xc.PROYECTO / "servicios").glob("*.py")):
        if archivo.stem.startswith("_"):
            continue
        modulo = importlib.import_module(f"servicios.{archivo.stem}")
        modulos[modulo.CODIGO] = modulo
    return modulos


# --- Tablas ----------------------------------------------------------------------

def leer_tablas(libro) -> dict:
    """{nombre de tarifa: {"paises": {tabla: [filas]}, "velocidades": {...}, "bolsas": {...}}}"""
    tablas = {"_hojas": set()}
    for tipo, titulo in HOJAS_TABLAS.items():
        if titulo not in libro.sheetnames:
            continue
        tablas["_hojas"].add(tipo)
        hoja = xc.Hoja(libro[titulo], obligatorias=("Nombre de la tarifa", "Tabla"))
        for r, f in hoja.filas("Nombre de la tarifa"):
            nombre, tabla = texto(f["Nombre de la tarifa"]), texto(f.get("Tabla"))
            if tipo == "paises":
                fila = {"tipo": texto(f.get("Tipo")) or "Bloque", "tarifa": texto(f.get("Tarifa")),
                        "paises": xc.lista(f.get("Países"))}
            elif tipo == "bolsas":
                fila = {"tipo": texto(f.get("Tipo")) or "Bolsa", "gigas": texto(f.get("Gigas")),
                        "plataformas": xc.lista(f.get("Plataformas"))}
            else:
                fila = {clave: texto(f.get(col)) for col, clave in VELOCIDAD_TABLA.items()}
            fila["_origen"] = f"{titulo} fila {r}"
            tablas.setdefault(nombre, {}).setdefault(tipo, {}).setdefault(tabla, []).append(fila)
    return tablas


# --- Fila → valores del código ---------------------------------------------------

def construir(servicio, f: dict, tablas: dict) -> dict:
    """Convierte una fila del Excel en lo que necesita el registro. Lanza DatoInvalido."""
    nombre = texto(f["Nombre de la tarifa"])
    if len(nombre) > 100:
        raise DatoInvalido("el nombre pasa de 100 caracteres")
    avisos, usados = [], set(META)
    dims = {}
    for columna, clave in DIMENSIONES.items():
        if columna not in f:
            continue
        usados.add(columna)
        if clave not in servicio.DIMENSIONES:
            if texto(f[columna]):
                avisos.append(f"«{columna}» no se usa en este servicio")
            continue
        partes = xc.lista(f[columna])
        if not partes:
            if clave == "alcance":
                raise DatoInvalido("falta «Alcance de la tarifa»")
            dims[clave] = servicio.DIMENSIONES[clave]["opciones"][0]
            if servicio.DIMENSIONES[clave]["multiple"]:
                dims[clave] = (dims[clave],)
            avisos.append(f"«{columna}» vacío: se usó {xc.texto(dims[clave]) if isinstance(dims[clave], str) else ' + '.join(dims[clave])}")
            continue
        if clave == "internet_ilimitado":
            partes = ["Si" if p.lower() in xc.SI else p for p in partes]
        # Alcance: en el código están TODAS sus opciones (en otras listas, solo las de por defecto).
        opciones = {xc.texto(o).lower(): o for o in servicio.DIMENSIONES[clave].get("opciones") or []} \
            if clave == "alcance" else {}
        if opciones:
            malas = [p for p in partes if xc.texto(p).lower() not in opciones]
            if malas:
                raise DatoInvalido(f"«{columna}»: «{', '.join(malas)}» no es una opción válida "
                                   f"({' / '.join(opciones.values())})")
            partes = [opciones[xc.texto(p).lower()] for p in partes]
        if servicio.DIMENSIONES[clave]["multiple"]:
            dims[clave] = tuple(partes)
        elif len(partes) > 1:
            raise DatoInvalido(f"«{columna}» admite una sola opción (tiene «{texto(f[columna])}»)")
        else:
            dims[clave] = partes[0]
    for clave, definicion in servicio.DIMENSIONES.items():   # listas del código que el Excel no trae
        if clave not in dims:
            primera = definicion["opciones"][0]
            dims[clave] = (primera,) if definicion["multiple"] else primera

    # Paso 1
    hoy = date.today()
    paso1 = {}
    for columna, clave in (("Inicio de vigencia", "fecha_inicio_vigencia"),
                           ("Inicio comercialización", "fecha_inicio_comercializacion")):
        fecha = xc.leer_fecha(f.get(columna), columna)
        if fecha and fecha < hoy:
            raise DatoInvalido(f"«{columna}» {fecha:%d/%m/%Y} es anterior a hoy")
        paso1[clave] = fecha
    if paso1["fecha_inicio_vigencia"] and paso1["fecha_inicio_comercializacion"] \
            and paso1["fecha_inicio_comercializacion"] < paso1["fecha_inicio_vigencia"]:
        raise DatoInvalido("la comercialización no puede empezar antes de la vigencia")
    if texto(f.get("Tarifa instalación")):
        paso1["tarifa_instalacion"] = xc.numero(f["Tarifa instalación"], "Tarifa instalación", minimo=0)
    if texto(f.get("Otros datos generales")):
        paso1["otros_datos"] = texto(f["Otros datos generales"])
    contacto = {clave: texto(f[c]) for c, clave in CONTACTO.items() if texto(f.get(c))}
    usados.update(PASO_1, CONTACTO)

    # Paso 2: se parte de los valores del código y se reemplaza lo que trae el Excel
    v = copy.deepcopy(servicio.VALORES)
    movil = "textos" in v
    del_excel = set()
    for columna, (clave, tipo) in SIMPLES.items():
        if columna not in f:
            continue
        usados.add(columna)
        if not texto(f[columna]):
            continue
        if clave not in v:
            avisos.append(f"«{columna}» no se usa en este servicio")
            continue
        if tipo == "texto":
            v[clave] = texto(f[columna])
        else:
            v[clave] = xc.numero(f[columna], columna, minimo=0, entero=tipo == "entero")
        del_excel.add(clave)
    if "duracion_dias" in v and texto(f.get("Duración")):
        if not 1 <= int(v["duracion_dias"]) <= 180:
            avisos.append("duración fuera de 1–180 días: SIRT podría rechazarla")

    for columna in [c for c in f if c.startswith(("Condiciones", "Restricciones", "Otras "))]:
        usados.add(columna)
        if not texto(f[columna]):
            continue
        if movil:
            v["textos"][columna] = texto(f[columna])
        elif columna in TEXTOS_FIJOS and TEXTOS_FIJOS[columna] in v:
            v[TEXTOS_FIJOS[columna]] = texto(f[columna])
        else:
            avisos.append(f"«{columna}» no se usa en este servicio")

    # Velocidades con degradación (móvil): mismas etiquetas que en SIRT
    for columna in [c for c in f if c.endswith("con degradación")]:
        usados.add(columna)
        if texto(f[columna]) and movil:
            v["degradacion"][columna] = xc.numero(f[columna], columna, minimo=0)

    # Apps sin consumo de datos (móvil, múltiple)
    if "Apps sin consumo de datos" in f:
        usados.add("Apps sin consumo de datos")
        if "apps_sin_consumo" in v:
            v["apps_sin_consumo"] = tuple(xc.lista(f["Apps sin consumo de datos"]))

    # Condicionales: «Título: ¿Es ilimitado?» + «Título: Cantidad…» + «Título: Tarifa…»
    destino = "minutos" if "minutos" in v else "condicionales" if "condicionales" in v else None
    titulos = [c.split(": ¿Es ilimitado?")[0] for c in f if c.endswith(": ¿Es ilimitado?")]
    if titulos and destino:
        nuevos = {}
        for t in titulos:
            cols = [c for c in f if c.startswith(t + ": ")]
            usados.update(cols)
            ilimitado = texto(f[f"{t}: ¿Es ilimitado?"])
            cantidad = next((f[c] for c in cols if ": Cantidad" in c), None)
            tarifa = next((f[c] for c in cols if ": Tarifa" in c), None)
            if xc.es_si(ilimitado):
                nuevos[t] = {"modo": "ilimitado"}
            elif ilimitado.lower() == "no" or texto(cantidad) or texto(tarifa):
                if not texto(cantidad) or not texto(tarifa):
                    raise DatoInvalido(f"«{t}»: si no es ilimitado, faltan la cantidad y la tarifa adicional")
                nuevos[t] = {"modo": "limitado", "minutos": xc.numero(cantidad, f"{t}: cantidad", minimo=0),
                             "tarifa": xc.numero(tarifa, f"{t}: tarifa", minimo=0)}
            else:
                nuevos[t] = {"modo": "omitir"}
        v[destino] = nuevos

    # Velocidades de Internet fijo (en línea): navegación (obligatoria) y degradación (opcional)
    for prefijo, clave in (("Velocidades de navegación", "velocidad_navegacion"),
                           ("Velocidades con degradación", "velocidad_degradacion")):
        cols = {c: c.split(": ", 1)[1] for c in f if c.startswith(prefijo + ": ")}
        usados.update(cols)
        if not cols or clave not in v:
            continue
        datos = {VELOCIDAD_IF[sub]: texto(f[c]) for c, sub in cols.items() if sub in VELOCIDAD_IF}
        if any(datos.values()):
            for k, valor in datos.items():
                if not valor:
                    raise DatoInvalido(f"«{prefijo}»: falta «{k}»")
                if k != "tecnologia":
                    datos[k] = xc.numero(valor, f"{prefijo}: {k}", minimo=0)
            v[clave] = datos
            if clave == "velocidad_degradacion":
                v["agregar_degradacion"] = True
        elif clave == "velocidad_degradacion" and "agregar_degradacion" in v:
            v["agregar_degradacion"] = False
        else:
            avisos.append(f"«{prefijo}» vacío: se usó la del código")

    # Canales (TV): archivos de automatizacion/canales/
    for columna in [c for c in f if c.endswith(": archivo Excel")]:
        usados.add(columna)
        etiqueta = columna.split(": archivo Excel")[0]
        archivo = texto(f[columna])
        if "archivos" not in v:
            continue
        if not archivo:
            avisos.append(f"«{etiqueta}» sin archivo: se usó el del código")
            continue
        ruta = CANALES / archivo
        if not ruta.exists():
            raise DatoInvalido(f"no existe el archivo de canales «{archivo}» en automatizacion/canales/")
        v["archivos"][etiqueta] = ruta

    # Tablas de las hojas «Tablas – …». Países y bolsas son opcionales en SIRT: si la hoja está
    # en el archivo, solo se registran las filas que trae para esta tarifa (ninguna = tabla vacía).
    # Velocidades móviles es obligatoria: sin filas se usa la del código.
    propias = tablas.get(nombre, {})
    if "paises" in tablas.get("_hojas", ()):
        for clave in LDI_FIJO.values():
            if clave in v:
                v[clave] = []
        for tabla in v.get("tablas_paises", {}):
            v["tablas_paises"][tabla] = []
    if "bolsas" in tablas.get("_hojas", ()):
        for tabla in v.get("bolsas", {}):
            v["bolsas"][tabla] = []
    if "velocidades" in v and not propias.get("velocidades"):
        avisos.append("sin filas en «Tablas – Velocidades»: se usaron las velocidades del código")
    for tabla, filas in propias.get("paises", {}).items():
        filas = [_sin_origen(x) for x in filas]
        for x in filas:
            if not x["paises"] or not x["tarifa"]:
                raise DatoInvalido(f"tabla «{tabla}»: cada fila necesita tarifa y al menos un país")
            x["tarifa"] = xc.numero(x["tarifa"], f"{tabla}: tarifa", minimo=0)
        if tabla in LDI_FIJO and LDI_FIJO[tabla] in v:
            v[LDI_FIJO[tabla]] = filas
        elif "tablas_paises" in v and tabla in v["tablas_paises"]:
            v["tablas_paises"][tabla] = filas
        else:
            avisos.append(f"tabla «{tabla}» no existe en este servicio (se ignoró)")
    for tabla, filas in propias.get("velocidades", {}).items():
        filas = [_sin_origen(x) for x in filas]
        if "velocidades" not in v:
            avisos.append(f"tabla «{tabla}» no existe en este servicio (se ignoró)")
            continue
        tecnologias = [x["tecnologia"] for x in filas]
        if len(set(tecnologias)) != len(tecnologias):
            raise DatoInvalido(f"tabla «{tabla}»: solo puede haber una fila por tecnología")
        for x in filas:
            for k in ("descarga_contratada", "descarga_garantizada", "carga_contratada", "carga_garantizada"):
                x[k] = xc.numero(x[k], f"{tabla}: {k}", minimo=0)
        v["velocidades"] = filas
    for tabla, filas in propias.get("bolsas", {}).items():
        filas = [_sin_origen(x) for x in filas]
        if "bolsas" not in v or tabla not in v["bolsas"]:
            avisos.append(f"tabla «{tabla}» no existe en este servicio (se ignoró)")
            continue
        for x in filas:
            x["gigas"] = xc.numero(x["gigas"], f"{tabla}: gigas", minimo=0)
        v["bolsas"][tabla] = filas

    # Regla de SIRT en móvil: tethering, traslada y roaming incluidos <= tope de descarga
    if movil and v.get("datos_tope_descarga_gb"):
        tope = float(v["datos_tope_descarga_gb"])
        for clave, titulo in (("datos_tethering_gb", "Datos para tethering"), ("traslada_datos_gb", "Traslada datos"),
                              ("datos_roaming_incluidos_gb", "Datos Roaming incluidos")):
            if v.get(clave) and float(v[clave]) > tope and clave not in del_excel:
                avisos.append(f"«{titulo}» vacío: se usó {tope:g} (no puede pasar del tope de descarga)")
                v[clave] = v["datos_tope_descarga_gb"]
            elif v.get(clave) and float(v[clave]) > tope:
                raise DatoInvalido(f"«{titulo}» ({v[clave]}) no puede ser mayor que «Datos tope descarga» ({tope:g})")
    if dims.get("internet_ilimitado") == "No" and "cantidad_datos_gb" in v and not texto(f.get("Cantidad de datos")):
        avisos.append(f"Internet ilimitado = No sin «Cantidad de datos»: se usó {v['cantidad_datos_gb']}")

    sin_uso = [c for c in f if c not in usados and texto(f[c])]
    if sin_uso:
        avisos.append("columnas sin uso todavía: " + ", ".join(sin_uso))

    return {"nombre": nombre, "dims": dims, "paso1": paso1, "contacto": contacto, "valores": v,
            "avisos": avisos, "definitivo": xc.es_si(f.get("Registro definitivo")),
            "disponibilidad": texto(f.get("Disponibilidad comercial")) or DISPONIBILIDAD_POR_DEFECTO}


def _sin_origen(fila: dict) -> dict:
    return {k: v for k, v in fila.items() if k != "_origen"}


# --- SIRT ------------------------------------------------------------------------

def registrar_en_sirt(page, servicio, datos: dict, estado: dict, diag) -> list[str]:
    import comun
    avisos = []
    paso1 = {**comun.PASO1_POR_DEFECTO, **getattr(servicio, "PASO1", {}), **datos["paso1"]}
    modulo_paso2 = sys.modules[servicio.completar_paso_2.__module__]
    originales = (servicio.VALORES, modulo_paso2.VALORES)

    def etapa(nombre, accion):
        estado["etapa"] = nombre
        resultado = accion()
        diag.captura_paso(page, nombre)
        return resultado

    etapa("abrir_nueva_tarifa", lambda: comun.abrir_nueva_tarifa(page))
    instalacion = etapa("paso_1_completo", lambda: comun.completar_paso_1(
        page, datos["nombre"], servicio.SERVICIO, datos["dims"]["alcance"], paso1))
    if "tarifa_instalacion" in datos["paso1"] and not instalacion:
        avisos.append("SIRT no habilitó «Tarifa instalación» para este servicio (se ignoró)")
    avisos += comun.llenar_contacto(page, datos["contacto"])
    estado["id_registro"] = etapa("paso_1_guardado", lambda: comun.guardar_paso_1(page))
    try:
        servicio.VALORES = modulo_paso2.VALORES = datos["valores"]
        etapa("paso_2_completo", lambda: servicio.completar_paso_2(page, datos["dims"]))
    finally:
        servicio.VALORES, modulo_paso2.VALORES = originales
    etapa("paso_2_guardado", lambda: comun.guardar_paso_2(page))
    etapa("confirmada", lambda: comun.confirmar_tarifa(page))
    return avisos


# --- Principal -------------------------------------------------------------------------

def main() -> None:
    args = leer_argumentos()
    ruta = Path(args.excel).resolve()
    libro = xc.abrir_libro(ruta, args.solo_listar)
    modulos = modulos_por_codigo()
    tablas = leer_tablas(libro)

    hojas = [n for n in libro.sheetnames if n in modulos]
    if args.hojas:
        pedidas = [h.upper() for h in args.hojas]
        desconocidas = [h for h in pedidas if h not in hojas]
        if desconocidas:
            print(f"Aviso: no hay hoja de servicio para {', '.join(desconocidas)}")
        hojas = [h for h in hojas if h in pedidas]

    trabajos, con_error, omitidas, ya = [], [], [], []
    lectores = {}
    for codigo in hojas:
        servicio = modulos[codigo]
        hoja = xc.Hoja(libro[codigo], obligatorias=("Nombre de la tarifa", "Alcance de la tarifa"))
        lectores[codigo] = hoja
        for r, f in hoja.filas("Nombre de la tarifa"):
            nombre = texto(f["Nombre de la tarifa"])
            if args.nombres and nombre not in args.nombres:
                continue
            ejecutar = xc.ejecutar_fila(f)
            item = {"codigo": codigo, "fila": r, "nombre": nombre, "servicio": servicio}
            if ejecutar == "no":
                omitidas.append(item)
                continue
            if ejecutar != "si":
                con_error.append((item, f"«Ejecutar» debe ser Sí o No (tiene «{ejecutar}»)"))
                continue
            try:
                item["datos"] = construir(servicio, f, tablas)
            except DatoInvalido as problema:
                con_error.append((item, str(problema)))
                continue
            previo = xc.exitosa_en(xc.RESULTADOS / codigo / f"resultados_{xc.LOTE_EXCEL}.csv", nombre)
            if previo and not args.repetir:
                item["previo"] = previo
                ya.append(item)
                continue
            trabajos.append(item)

    nombres_vistos = {}
    for item in trabajos:
        if item["nombre"] in nombres_vistos:
            con_error.append((item, f"nombre repetido (también en {nombres_vistos[item['nombre']]})"))
        nombres_vistos[item["nombre"]] = f"{item['codigo']} fila {item['fila']}"
    trabajos = [t for t in trabajos if all(t is not e[0] for e in con_error)]
    if args.limite:
        trabajos = trabajos[: args.limite]

    # Ya registradas: solo falta (quizá) el registro definitivo
    definitivos_pendientes = [i for i in ya if i["datos"]["definitivo"] and not args.sin_definitivo
                              and not xc.exitosa_en(xc.RESULTADOS / f"registro_definitivo_{xc.LOTE_EXCEL}.csv", i["nombre"])]

    print(f"Excel: {ruta}")
    print(f"Hojas de servicio: {', '.join(hojas) or '(ninguna)'}")
    print(f"A registrar: {len(trabajos)} · ya registradas: {len(ya)} · omitidas (Ejecutar = No): {len(omitidas)}"
          f" · con errores de datos: {len(con_error)}")
    for item, problema in con_error:
        print(f"  ✖ {item['codigo']} fila {item['fila']} ({item['nombre']}): {problema}")
    for k, item in enumerate(trabajos, start=1):
        d = item["datos"]
        listas = " · ".join(f"{c}: {xc.texto(' + '.join(x) if isinstance(x, tuple) else x)}" for c, x in d["dims"].items())
        print(f"  {k:>2}. {item['codigo']} fila {item['fila']}: {item['nombre']}  [{listas}]"
              f"  valor {d['valores'].get('valor_servicio', '?')}"
              + ("  + registro definitivo" + (f" ({d['disponibilidad']})" if d["disponibilidad"] else "") if d["definitivo"] and not args.sin_definitivo else ""))
        for aviso in d["avisos"]:
            print(f"        aviso: {aviso}")
    for item in definitivos_pendientes:
        print(f"   · {item['nombre']}: ya registrada; falta su registro definitivo")

    if args.solo_listar:
        print("\n--solo-listar: no se tocó SIRT ni el Excel.")
        return

    for item in omitidas:
        lectores[item["codigo"]].escribir(item["fila"], "Omitida")
    for item, problema in con_error:
        lectores[item["codigo"]].escribir(item["fila"], "Error de datos", detalle=problema)
    for item in ya:
        p = item["previo"]
        if lectores[item["codigo"]].estado(item["fila"]) not in ("Registrada", "Registro definitivo"):
            lectores[item["codigo"]].escribir(item["fila"], "Ya registrada", p.get("id_registro", ""), p.get("fecha_hora", ""),
                                              "Ya figura como registrada en el lote EXCEL (no se repitió).")
    if not trabajos and not definitivos_pendientes:
        xc.guardar(libro, ruta)
        print("\nNo hay tarifas para registrar.")
        return
    total_def = sum(1 for t in trabajos if t["datos"]["definitivo"]) + len(definitivos_pendientes)
    if not xc.confirmar(f"se registrarán {len(trabajos)} tarifa(s) CONFIRMADAS en SIRT"
                        + (f" y {total_def} registro(s) definitivo(s) (irreversibles)" if total_def and not args.sin_definitivo else ""),
                        args.si):
        print("Cancelado.")
        return

    import comun
    from ejecutar import Diagnostico
    from playwright.sync_api import sync_playwright

    resumen = {}
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(comun.CDP_URL)
        context = browser.contexts[0]
        page = comun.obtener_pagina_sirt(context)
        page.set_default_timeout(comun.TIMEOUT)
        if args.al_frente:
            page.bring_to_front()
        if trabajos:
            page.get_by_role("button", name="Nueva Tarifa").wait_for(state="visible")
        url_listado = page.url

        for k, item in enumerate(trabajos, start=1):
            servicio, datos, hoja = item["servicio"], item["datos"], lectores[item["codigo"]]
            rutas = comun.rutas_lote(servicio.CODIGO, xc.LOTE_EXCEL)
            from ejecutar import siguiente_correlativo
            correlativo = siguiente_correlativo(rutas["resultados"])
            diag = Diagnostico(args.diagnostico, False, rutas["base"] / f"diagnostico_{xc.LOTE_EXCEL}")
            diag.iniciar(context)
            diag.iniciar_caso({"correlativo": correlativo, "nombre": datos["nombre"]})
            print("\n" + "=" * 70)
            print(f"[{k}/{len(trabajos)}] {item['codigo']} fila {item['fila']}: {datos['nombre']}")
            estado = {"etapa": "", "id_registro": ""}
            resultado = {"correlativo": correlativo, "nombre": datos["nombre"],
                         "clave": " | ".join(comun.texto_valor(x) for x in datos["dims"].values()),
                         "id_registro": "", "resultado": "", "detalle": "", "captura": "",
                         "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            avisos = list(datos["avisos"])
            try:
                avisos += registrar_en_sirt(page, servicio, datos, estado, diag)
                resultado["resultado"] = "EXITOSO"
                resultado["detalle"] = "registrada desde Excel"
                print(f"   CONFIRMADA (registro {estado['id_registro']})")
                fallo = False
            except Exception as error:
                fallo = True
                rutas["capturas"].mkdir(parents=True, exist_ok=True)
                captura = rutas["capturas"] / f"{correlativo:03d}_error.png"
                try:
                    page.screenshot(path=str(captura), full_page=True)
                except Exception:
                    pass
                diag.registrar_fallo(page, error, estado["etapa"])
                resultado["resultado"] = "ERROR"
                resultado["detalle"] = f"[{estado['etapa']}] " + str(error).replace("\n", " ")[:500]
                resultado["captura"] = str(captura)
                print(f"   ERROR en '{estado['etapa']}': {error}")
            diag.cerrar_caso(fallo)
            diag.finalizar()
            resultado["id_registro"] = estado["id_registro"]
            comun.registrar_resultado(rutas["resultados"], resultado)

            detalle = " | ".join(avisos + [resultado["detalle"]])
            if fallo:
                hoja.escribir(item["fila"], "Error", estado["id_registro"], resultado["fecha_hora"], detalle)
            else:
                hoja.escribir(item["fila"], "Registrada", estado["id_registro"], resultado["fecha_hora"], detalle)
                item["registrada"] = True
                if datos["valores"].get("valor_servicio"):
                    comun.guardar_valor_tarifa(datos["nombre"], datos["valores"]["valor_servicio"])
            resumen["Error" if fallo else "Registrada"] = resumen.get("Error" if fallo else "Registrada", 0) + 1
            xc.guardar(libro, ruta)
            try:
                comun.volver_al_listado(page, url_listado)
            except Exception as error:
                print(f"No se pudo volver al listado ({error}). Se detiene; vuelve a correr para continuar.")
                break
            page.wait_for_timeout(1_000)

    # Registro definitivo (después, en su propio proceso)
    if not args.sin_definitivo:
        para_definitivo = [t for t in trabajos if t.get("registrada") and t["datos"]["definitivo"]] + definitivos_pendientes
        if args.definitivo_en_bloque:
            xc.precargar_definitivos([(i["nombre"], i["datos"]["disponibilidad"]) for i in para_definitivo],
                                     xc.LOTE_EXCEL, args.tamano_bloque, args.diagnostico)
        for item in para_definitivo:
            hoja = lectores[item["codigo"]]
            ubigeos = item["datos"]["disponibilidad"]
            ok, detalle = xc.hacer_registro_definitivo(item["nombre"], xc.LOTE_EXCEL, ubigeos, args.diagnostico)
            actual = xc.texto(hoja.hoja.cell(item["fila"], hoja.c_detalle).value) if hoja.c_detalle else ""
            codigo_sirt = xc.texto(hoja.hoja.cell(item["fila"], hoja.c_codigo).value) if hoja.c_codigo else ""
            if not codigo_sirt and item.get("previo"):
                codigo_sirt = item["previo"].get("id_registro", "")
            hoja.escribir(item["fila"], "Registro definitivo" if ok else "Error", codigo_sirt,
                          datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                          (actual + " | " if actual else "") + detalle)
            clave = "Registro definitivo" if ok else "Error en registro definitivo"
            resumen[clave] = resumen.get(clave, 0) + 1
            xc.guardar(libro, ruta)

    copia = xc.guardar_con_copia(libro, ruta, "REGISTRO_EXCEL", "registro_excel")
    print("\n" + "=" * 70)
    print("Resumen: " + (" · ".join(f"{k}: {v}" for k, v in resumen.items()) or "sin cambios"))
    print(f"Tiempo total: {xc.tiempo_total(len(trabajos))}")
    print(f"Excel actualizado: {ruta}")
    print(f"Copia de esta corrida: {copia}")


if __name__ == "__main__":
    main()
