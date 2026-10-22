# Contexto del proyecto (para continuar en una conversación nueva)

Resumen de lo trabajado hasta el 13/10/2026. En una conversación nueva, pedir a
Claude que lea este archivo y `LEEME.md` antes de continuar.

## Objetivo

Automatizar con Playwright (Python) el registro de tarifas en **SIRT** (OSIPTEL),
ambiente QA (`sirt-qa.be1business.com`), para el operador **FLOR TV E.I.R.L.**
(RUC 20601544629).

- Tipo de tarifa: **Tarifa establecida** (fijo).
- Condición: **Monoproducto** (fijo; existe también "Empaquetada", no cubierta).
- Para cada servicio se generan **todas las combinaciones** de los valores de sus
  listas y se crea una tarifa **confirmada** por cada una.

Antecedente: un proyecto anterior (otro operador, START-UP COMPANY) automatizó
"Licitación y Contratos" con 511 combinaciones de servicios y operaciones
posteriores (renombre, rectificación, fe de erratas, ficha informativa).
Sus scripts sirvieron de base, pero ese formulario era distinto.

## Entorno

- Ubuntu 24.04, entorno virtual de Python con `playwright` instalado.
- Edge en modo controlable (el script se conecta a él; el login es manual):
  `microsoft-edge --remote-debugging-port=9222 --user-data-dir="$HOME/perfiles/sirt" &`
- Para usar el equipo mientras corre (14/10): `ejecutar.py` ya no hace
  `bring_to_front` (salvo `--al-frente`); abrir Edge además con
  `--disable-backgrounding-occluded-windows --disable-renderer-backgrounding
  --disable-background-timer-throttling` y moverlo a otro escritorio (ver LEEME).
- `probar_todos.sh [lote] [limite] [opciones]`: corre los 15 servicios seguidos
  (por defecto lote MUESTRA, 1 caso, --diagnostico) y da un resumen + log.
- El login (usuario, contraseña, reCAPTCHA y código por correo) lo hace la persona.
  Los scripts asumen la sesión iniciada y la pantalla en **Tarifas registradas**.
- Grabación de flujos: `grabar_sirt.py` (se conecta al Edge y abre el Inspector;
  botón Record).

## Estructura del proyecto (`sirt_tarifas/`)

- `comun.py`: conexión, listas, fechas, paso 1, bloques reutilizables, confirmación,
  combinaciones, plan, resultados y reanudación.
- `ejecutar.py`: ejecutor único.
  `python ejecutar.py <servicio> [--solo-calcular] [--limite N] [--lote X] [--diagnostico] [--trace] [--si]`
  Sin `--limite` pide confirmación (escribir `si`); `--si` la omite.
- `inspeccionar_json.py <servicio>`: resume `grabaciones/<servicio>/paso2.json`
  (campos, opciones exactas, borrador de DIMENSIONES).
- `grabaciones/<servicio>/`: material de cada servicio (paso2.json,
  grabacion_paso2.py, capturas .png, notas.txt). Ya existen las carpetas de los
  13 servicios pendientes, con archivos vacíos para llenar.
- `servicios/<servicio>.py`: SERVICIO, CODIGO, DIMENSIONES, VALORES y
  `completar_paso_2(page, valores)`. Plantilla en `servicios/_plantilla.py`.
- `resultados/<CODIGO>/`: plan, resultados (con `id_registro`) y capturas de error.

## Flujo en SIRT

1. Nueva Tarifa → Tipo de Tarifa → Condición → Seleccionar servicio (uno).
2. Paso 1: nombre, Alcance (múltiple: Residencial / Comercial), Inicio de vigencia,
   Inicio comercialización (hoy), Tarifa Instalación (solo si está habilitada),
   Otros datos generales → "Guardar y continuar con el paso"
   (crea un borrador; la URL pasa a `nueva-tarifa/<id>?step=2`).
3. Paso 2: depende del servicio. Termina con Condiciones generales y
   Restricciones técnicas → "Guardar y continuar con el paso".
4. Resumen → "Confirmar Tarifa" → "Confirmar tarifa" → "Aceptar".

## Servicios terminados y probados (funcionan)

| Archivo            | Servicio                | Casos | Combinaciones                                     |
|--------------------|-------------------------|-------|---------------------------------------------------|
| `internet_fijo.py` | Internet fijo postpago  | 45    | alcance 3 × periodicidad 5 × tipo de acceso 3     |
| `telefonia_fija.py`| Telefonía fija postpago | 315   | alcance 3 × periodicidad 5 × tipo de acceso 7 × tecnología 3 |
| `telefonia_fija_prepago.py` | Telefonía fija prepago | 63 | alcance 3 × tipo de acceso 7 × tecnología 3 |

Telefonía fija prepago (código TFP): como postpago pero sin Periodicidad y con
"Duración" (días, fijo en 30); sí tiene Tarifa Instalación. Probado el 13/10/2026:
casos 1–7 del lote 20261013 confirmados (registros 26198–26204). El lote se
cortó a propósito: por ahora solo se quieren pruebas de 1–2 casos por servicio,
no lotes completos. Revisar si quedó un borrador `TE-TFP-20261013-008`.
JSON marca Tecnologia como "listado múltiple"; se dejó como única (criterio de
postpago), pendiente de confirmar.

Televisión de paga Postpago (`tv_paga_postpago.py`, código TVP, 90 casos =
alcance 3 × periodicidad 5 × tipo de acceso 3 × tecnología 2). Nuevo: 3 campos
"Carga y descarga" (Canales analógicos / SD / HD), obligatorios, cargados con
el mismo Excel `grabaciones/tv_paga_postpago/Canales analógicos_1.xlsx` mediante
`comun.cargar_archivo` (busca la fila por su título; el selector del grabador
`nameInputRamdom...` cambia en cada carga). Probado en SIRT el 13/10/2026: OK.

Creados el 13/10/2026 (pendientes de probar con `--limite 1 --diagnostico`,
salvo que se indique):

| Archivo | Servicio | Código | Casos | Notas |
|---|---|---|---|---|
| `tv_paga_prepago.py` | Televisión de paga Prepago | TVPR | 18 | como TV postpago, Duración en vez de Periodicidad |
| `otros_servicios.py` | Otros servicios | OS | 45 | nuevo: Tipo de moneda (Soles/Dolares/Euro) + "Otras condiciones tarifa" y "Otras restricciones" |
| `paquetes_internet_fijo.py` | Paquetes adicionales y otros servicios de internet fijo | PIF | 21 | Tecnología (IF) con sus 7 opciones; 1 tabla de velocidades |
| `paquetes_tv_paga.py` | Paquetes adicionales y otras tarifas de televisión de paga | PTV | 9 | nuevo: Tipo de Contratación (múltiple); Excel de `grabaciones/tv_paga_postpago/` |
| `paquetes_telefonia_fija.py` | Paquetes adicionales de telefonía fija | PTF | 9 | nuevo: Tipo de línea (múltiple); 2 tablas LDI, sin minutos |
| `otras_tarifas_telefonia_fija.py` | Otras tarifas de telefonía fija | OTF | 45 | nuevo: Tarifa por minuto; sin Tarifa Instalación |
| `movil_postpago.py` | Servicio público móvil postpago | SMP | 105 | ver abajo |
| `movil_control.py` | Servicio público móvil control | SMC | 105 | paso 2 idéntico a postpago (mismo JSON): reutiliza `movil_postpago.py` |
| `movil_paquete.py` | Servicio público móvil paquete | SMQ | 21 | Duración + Tipo de línea móvil (Postpago/Control/Prepago, múltiple); el resto con `movil_postpago.completar_resto_movil` |
| `movil_otros.py` | Servicio móvil otros | SMO | 21 | formulario móvil completo (campos de postpago + de paquete, casi todo opcional). Periodicidad fija "Mensual" y Tipo de acceso = los tres; combinando todo serían 735 |
| `arrendamiento_circuitos.py` | Arrendamiento de circuitos | AC | 45 | Tipo de moneda + Periodo de contrato (meses) + Tecnología de transmisión (texto libre, máx. 40) + 2 velocidades |

Regla de SIRT en móvil: "Datos para tethering", "Traslada datos" y "Datos
Roaming incluidos en los datos del plan" deben ser <= "Datos tope descarga".

1.ª prueba de móvil postpago (13/10, 17:44): llegó hasta los minutos y falló
en "Minutos móviles off net": esa fila está armada distinto y la búsqueda por
"contenedor de la fila" tomaba la sección entera (9 casillas). Corregido:
`comun.campos_condicional` toma la casilla y los 2 campos que SIGUEN al título.
2.ª prueba (14/10, 09:00): pasaron los minutos; falló en la 1.ª tabla por
países ("Tarifa por minuto LDI" abrió la ventana de "Bolsa de datos"): mismo
problema de anidamiento. Corregido: `comun.boton_agregar` toma el primer
"Agregar" que SIGUE al título. Borradores sueltos de estas pruebas: 26225, 26230.
3.ª prueba (14/10): Móvil postpago OK (registro 26231) y Móvil control OK
(registro 26232), ambos confirmados.

Servicio público móvil postpago: formulario de ~45 campos. Nuevo:
`comun.agregar_bolsa_datos` (tablas "Bolsa de gigas") y `comun.boton_agregar`
(ubica el "Agregar" de cada tabla por el título de su sección; hay 15 tablas).
Regla de SIRT: "Datos para tethering" y "Traslada datos" deben ser <= "Datos
tope descarga" (la grabación falló con 24 > 12). Por ahora fijos: Componentes
del plan = ambos ("Voz y SMS" + "Datos"; falta ver si con uno solo el
formulario oculta campos) e Internet Ilimitado = "Si" (como Internet fijo).
Combinando todo serían 3 × 5 × 7 × 3 × 2 = 630 casos.

Internet fijo admite además Tecnología (IF) con 7 opciones e Internet Ilimitado
Si/No (630 casos en total), hoy fijados en HFC y "Si".

Servicios disponibles para el operador (15): Telefonía fija postpago/prepago,
Servicio público móvil postpago/control/paquete, Televisión de paga
Postpago/Prepago, Internet fijo postpago, Arrendamiento de circuitos,
Otros servicios, Servicio móvil otros, Paquetes adicionales y otras tarifas de
televisión de paga, Paquetes adicionales y otros servicios de internet fijo,
Paquetes adicionales de telefonía fija, Otras tarifas de telefonía fija.

## Bloques reutilizables probados (en comun.py)

- `agregar_velocidad`: ventana de velocidades (Tecnología "Única" + 4 velocidades).
- `completar_condicional_triple`: filas de minutos ("Es ilimitado" → no llenar;
  si no, llenar Cantidad de minutos y Tarifa por minuto adicional). Ubica la fila
  por su título.
- `agregar_tarifa_paises`: ventana LDI (Tipo "Bloque", Tarifa, Países; escribe el
  país para filtrar la lista).
- `cargar_archivo`: campos "Carga y descarga" (adjunta el archivo sin abrir la
  ventana del sistema y espera a que aparezca su nombre en la fila).
- `elegir_valor`, `llenar_si_editable`, `completar_textos_finales`.

## Lecciones aprendidas

- Las opciones se buscan con texto exacto. Varían por servicio:
  "Alta Nueva" (IF) vs "Alta nueva" (TF); "Si" sin tilde; "Tecnología (IF)" vs
  "Tecnologia" sin tilde. Tomar los textos del JSON del servicio.
- El grabador a veces registra `listbox` en vez de la opción elegida: completar
  las opciones con capturas o con el JSON.
- Las ventanas emergentes no están marcadas como `dialog`: se detectan por un
  campo propio, y mientras están abiertas el resto del formulario queda oculto.
- Selectores con números internos (`#mat-input-102`) cambian en cada carga: nunca usarlos.
- El JSON del GET del paso 2 describe todo el formulario del servicio (campos,
  tipo, obligatoriedad, listas, validaciones). Más adelante se puede automatizar.

## Cómo se agrega un servicio

1. El usuario envía: grabación del paso 2, JSON del GET, capturas de bloques
   nuevos o listas no registradas, y reglas especiales.
2. Se crea `servicios/<nombre>.py` desde la plantilla; los bloques comunes nuevos
   van a `comun.py`.
3. `--solo-calcular`, luego `--limite 1`.

## Fase siguiente: después de crear las tarifas (mapeado, sin implementar)

Las tarifas confirmadas pasan a **Tarifas temporales** (`/extranet/tarifa-temporal`).
Ahí hay que, para cada una nuestra (estado "Ingreso confirmado"; nunca las
"En proceso - Paso 1", que son borradores):

1. **Disponibilidad comercial** (solo en algunos servicios; probablemente los
   que tienen `flagDisponibilidadComercial: true` en su `paso2.json`):
   marcar la fila → "Registro de Disponibilidad comercial" (o ⋮ → "Ver
   disponibilidad comercial") → árbol "Ubigeo para el listado de tarifas"
   ("A nivel nacional" o departamentos, desplegables) → Guardar → resumen
   "Tarifas seleccionadas / Guardados / Fallidos" → Aceptar.
   El botón de arriba parece trabajar en bloque (varias filas a la vez).
2. **Registro definitivo**: marcar las filas → "Registro definitivo".
   Probablemente irreversible: probar con 1.

**Implementado (14/10): `registro_definitivo.py`** (probado con página
simulada; pendiente de probar en SIRT). Flujo grabado, una tarifa a la vez:
recarga Tarifas temporales (sin filas marcadas) → "Filtro avanzado" →
"Ingresa nombre de tarifa" → Buscar → fila confirmada con ese nombre (descarta
borradores con el mismo nombre) → [si requiere: Acciones → "Ver Disponibilidad
Comercial" → ubicaciones (por defecto Amazonas) → Guardar → Aceptar ×2] →
marcar fila → "Registro definitivo" → "Registrar tarifa" → Aceptar → verifica
que ya no esté en temporales. Protecciones: exige 1 sola fila marcada y que la
confirmación no indique más de 1; revisa "Fallidos: N". Lee las tarifas de
`resultados/*/resultados_<lote>.csv` (EXITOSO) y guarda en
`resultados/registro_definitivo_<lote>.csv` (reanudable).
Uso: `python registro_definitivo.py --lote MUESTRA --solo-listar | --limite 1 --diagnostico`.
Qué servicio lleva disponibilidad: `REQUIERE_DISPONIBILIDAD` (por código) en el script.
Pendiente: modo de varias a la vez (botón "Registro de Disponibilidad
comercial" con varias filas marcadas; no está grabado) y lista de ubicaciones.
La API del filtro es POST `/api/tarifa/filtrar-temporal` (devuelve por tarifa
`estado`, `paso`, `disponibilidadComercial`, `modalidad`).

**Empaquetada (14/10): `empaquetar.py`** (probado con página simulada;
pendiente en SIRT). Paso 1: Condición "Empaquetada" → "Agregar Tarifa" →
buscar por nombre → marcar → Agregar (espera "Total N") → datos de la tarifa →
Guardar. Paso 2: por pestaña "Servicio N" (identificada por el nombre del
servicio que muestra) llena "Valor del servicio" y pulsa Guardar; luego
continuar y confirmar. Reglas: mismo alcance en todas; valor de la más cara ≤
suma de valores (= Valor del plan) < suma de las tarifas base (bases 30 y 50:
3+5 rechaza por ser < 50; 30+50 rechaza por no ser < 80, probado 14/10 con
borrador 26266; 29+49 acepta). Por defecto: valor de cada tarifa − 1. Base: tarifas MUESTRA con registro definitivo. Por definir: qué paquetes
se armarán y si se pueden mezclar servicios libremente.
Registro definitivo del paquete (confirmado 14/10): pide disponibilidad
comercial solo si alguna de sus tarifas la requiere (TFP + SMP sí; SMP + OS
no). `registro_definitivo.py` ya lo calcula desde la columna `clave` (código EMP).
Ubicaciones del paquete (regla de OSIPTEL, 14/10): deben estar en la
intersección de las ubicaciones de sus tarifas que requieren disponibilidad
(p. ej. Amazonas+Arequipa y Arequipa+Lima → Arequipa). `registro_definitivo.py`
las lee de `registro_definitivo_*.csv` (columna `disponibilidad`); sin
`--ubigeos` usa toda la intersección. Tarifas registradas a mano deben
anotarse ahí con sus ubicaciones.

**Optimización de velocidad (mapeada 15/10, sin implementar).** Tiempos
medidos (lote PROMO, 14/10): fijos/TV/paquetes 7–15 s por tarifa; móviles
~30 s (25 s en el paso 2); registro definitivo ~10 s sin disponibilidad y ~20 s
con ella. Ideas, de mayor a menor efecto:
1. Registro definitivo en bloque (marcar varias filas → disponibilidad →
   registro definitivo); falta grabarlo.
2. Paralelo por servicio: N procesos `ejecutar.py` a la vez, cada uno con su
   pestaña en el mismo Edge (no ThreadPool: la API síncrona de Playwright no
   admite hilos); CSV y correlativos ya van por servicio. Lanzador `--paralelo N`.
   En registro definitivo, un CSV por proceso o con bloqueo. Probar 2–3
   procesos (riesgos: pestañas en segundo plano, límites del servidor QA).
3. `--trace` de un caso móvil para ver qué frena su paso 2.
4. Sin `--diagnostico` en lotes (1–2 s por tarifa).
También: `--total N` en ejecutar.py (dejar N por servicio, hoy `--limite` =
crear N más) y `--por-servicio N` en registro_definitivo.py.

**Promocionales: `promocionar.py` (15/10; probado con página simulada,
pendiente en SIRT).** Catálogo de atributos por servicio = 2.º JSON de
`grabaciones/_promocional/<servicio>/paso2.json` (data = secciones → apartados;
con subApartados una fila por sub; tipos PROMOCIONAL, PROMOCIONAL_VELOCIDAD,
PROMOCIONAL_ADJUNTO (Excel + duración), PROMOCIONAL_APP (plataforma + duración)).
Cada fila: casilla (nombre exacto) y los controles que la siguen hasta la
próxima casilla (orden del documento): mat-select Tipo de variación,
Variación, Duración. Telefonía fija no tiene el 2.º JSON (solo atributos
probados); otros servicios no trae atributos (solo paso 1). Falta la grabación
de `_empaquetada`. Pendiente: registro definitivo de promociones.
Mapeo previo:
**Siguiente: tarifas promocionales (mapeado 14/10, sin implementar).**
Flujo visto en capturas: Nueva tarifa → Tipo de Tarifa "Tarifa promocional" →
"Agregar Tarifa" (misma ventana que en empaquetadas: filtro por nombre, Tipo
de tarifa fijo en "Tarifa establecida", marcar 1 o más → Agregar) → datos de
la tarifa como en el paso 1 (nombre, comercialización, enlace, fono, email,
condiciones) + bloque "Datos promocionales": casilla "Valor del plan" (Tipo de
variación Ilimitado/Porcentual/Cantidad, Variación, Duración en días) y
casilla "Tarifa instalación" (Tipo de variación, Variación) → "Guardar y
continuar con el paso 2". Paso 2: una pestaña "Servicio N" por servicio (en
un paquete, una por cada servicio), con el servicio y su código SIRT; cada
atributo se activa con su casilla y lleva Tipo de variación, Variación (con
unidad, p. ej. Mbps) y Duración → Guardar → paso 3 → confirmar.
Regla visible: la duración de una promoción no supera 180 días.
Plan: `grabaciones/_promocional/<servicio>/` con grabación + paso2.json de la
promoción de cada monoproducto; un `promocionar.py` que reutilice la ventana
"Agregar Tarifa" de empaquetar.py y, en el paso 2, identifique cada pestaña por
el nombre del servicio (como en empaquetados) y llame a los atributos
promocionales del módulo de ese servicio. Así sirve para monoproductos y paquetes.
Por confirmar: qué pasa al promocionar varias tarifas a la vez (¿una pestaña
por tarifa?), si "Ilimitado" deja vacía la variación, qué indica la flecha de
Variación, y si la promoción lleva registro definitivo y disponibilidad.

**Promocionales probadas en SIRT (15/10):** 1 y 2 tarifas por servicio,
todas EXITOSO (TV tras corregir la carga del Excel: se usa el botón de carga de
la fila, el input oculto está antes de su casilla). Registro definitivo de
promociones: pendiente; puede hacerse con o sin disponibilidad (por confirmar).

**Acciones (15/10, recogiendo grabaciones en `grabaciones/_acciones/`):**
renombre, cese y fe de erratas desde Tarifas registradas. Cese y fe de erratas
exigen tarifa "En comercialización" y vigente (pasa a ese estado el día
siguiente al registro definitivo): validar el estado en el filtro antes de actuar.
Base: `referencia_lyc/renombrar_tarifas.py` y `fe_erratas_tarifas.py`.
Implementado (15/10): `acciones.py buscar|renombre` (probado con página
simulada). Búsqueda: Filtro avanzado → "Ingresa nombre de tarifa" → Buscar; el
estado se lee del JSON del filtro (data.items: codigoSIRT, version,
comercializacion, vigencia, esUltimaVersion); la fila se ubica por código y
versión. El renombre funciona también con "Por iniciar comercialización".
`acciones.py cese` (15/10, simulada): ⋮ → "Cese de Comercialización" →
calendario "Fin de comercialización" (por defecto mañana) → Guardar → Guardar
tarifa → Aceptar. Valida En comercialización y vigencia antes (nombre exacto del
estado de vigencia: "Vigente"). Cese probado en SIRT el 15/10.
`acciones.py fe_erratas` (15/10, simulada): ⋮ → "Fe de Erratas" → paso 1
(textareas, p. ej. Otros datos generales) → "Continuar con el paso 2" → paso 2
(textareas del servicio; 3 formas vistas: fijos/IF 2 campos, OS/AC 4, móviles
~14) → "Continuar con el paso 3" → resumen → "Guardar cambios" (lo que sigue no
está grabado: se maneja confirmación genérica + Aceptar). Se rellenan todas las
textareas visibles con un sello con fecha y hora. Grabaciones en
`grabaciones/_acciones/fe_erratas/caso1..3`; 6 servicios por revisar el 16/10.
Fe de erratas probada en SIRT (15/10) en IF, PTF, OS y SMP (casos 1-3).
16/10: revisados los 6 restantes → OTF, TVP, TVPR, PIF = caso1; SMQ = caso3;
PTV = caso4 (caso1 + "Tipo de Contratación", mat-select múltiple Postpago/Prepago
en el paso 2; grabación con TE-PTV-PROMO-003 hasta el resumen). `cambiar_lista_multiple`
la cambia sin dejarla vacía (2 marcadas → desmarca la última; 1 → marca otra),
cierra con Escape y busca las opciones finales en el resumen. `--sin-listas` la omite.
Fallos del 16/10: SMQ-PROMO-001 no existe (usar SMQ-PROMO-002); PIF ignoró el clic en
"Continuar con el paso 2" → `avanzar()` reintenta hasta 3 veces. El paso 1 se leía
antes de cargar (0 textos, instalación "(vacío) → 1" en OS/SMC del 15/10) →
`esperar_datos_cargados()`. Número vacío: espera ~5 s; si sigue vacío pone 1 o
`--valor-si-vacio`.
16/10 (prueba de 15): las textarea del paso 2 están en la página desde el paso 1
(pestaña oculta) → el paso 1 contaba de más y SMQ falló al llenar una que se ocultó.
`corregir_textos` solo toma las de la pestaña activa (`MARCAR_TEXTOS`, mat-tab-body
activo) y salta las que se ocultan. SMO-PROMO-001 se renombró a mano a
`TE-SMO-PROMO-001_RENOMBRE`; OTF/PTF/TVPR/AC/SMO-PROMO-002 tienen sufijo `_REN`.
**Cambio de tarifa** (16/10; grabación OS TE-OS-MUESTRA-001 en
`grabaciones/_acciones/cambio/`, verificación de campos en `verificacion_cambio.xlsx`):
menú "Cambio de Tarifa" → paso 1 (Fecha de ejecución con calendario, Tarifa de
instalación, Otros datos generales; igual en los 15) → paso 2 (el del registro sin los
campos que definen el plan: periodicidad/duración, tipo de acceso/línea/contratación,
tecnología, internet ilimitado, componentes; OS y AC además sin Tipo de moneda, AC sin
Tecnología de transmisión) → resumen → Guardar cambios → 2 Aceptar. Regla SIRT: al subir
precios la fecha no puede ser anterior a hoy + 2 días (`DIAS_ANTICIPACION_ALZA`); al
bajar: hoy siempre está bloqueado en el calendario → mañana (`DIAS_ANTICIPACION_BAJA` = 1).
`terminar_guardando` ahora falla si el mensaje final es de validación (antes lo daba por
EXITOSO). Por confirmar: qué pasa con la tarifa tras el cambio (versión, estado).
`--todo` (16/10): `cambiar_todos_los_numeros` (MARCAR_NUMEROS, pestaña activa) cambia
todos los input numéricos del paso 2 salvo Valor del servicio (aparte); paso = 1 unidad
del último decimal; 0 no se toca; % ≤ 100; topes primero al subir / al final al bajar.
Tablas (filas registradas) sin tocar: para editarlas haría falta grabar la edición de una fila.
`acciones.py rectificar` (15/10, simulada; grabación OS completa): menú
"Rectificar" → paso 1 (Alcance, "Tarifa de instalación" textbox, Otros datos
generales) → paso 2 (= paso 2 del registro, ya lleno) → paso 3 → Guardar cambios
→ 2 Aceptar. Cambia: textos + "| Rectificación <fecha>" y +1 en Valor del
servicio y Tarifa de instalación (leídos de la pantalla). Por confirmar: si crea
versión 2 y si pasa a temporales / registro definitivo.

**Después: renombrar** (sí se usará). Base: `referencia_lyc/renombrar_tarifas.py`.
Rectificar y fe de erratas quedan para más adelante (sus campos dependen del servicio).

## Proyecto anterior (referencia)

Scripts en `referencia_lyc/` (ver su LEEME.md). Eran de "Licitación y
Contratos" (START-UP COMPANY, ambiente sirt.pe…, no QA): 511 tarifas que
combinaban 1 a 9 servicios en una misma tarifa. NO hacían disponibilidad ni
registro definitivo. Piezas reutilizables: búsqueda por código con "Filtro
avanzado", menú `get_by_label("Acciones")` → `menuitem`, paginador a 100
filas (`listar_tarifas.py`), versión más alta, doble "Aceptar".

**Acciones en EMPAQUETADOS** (16/10; grabaciones con TE-EMP-MUESTRA-006 en
`grabaciones/_acciones/empaquetado/`, antes llamada `promocional/`; las promociones TP van en `_acciones/promociones_tp/`):
fe de erratas / rectificar / cambio = mismo flujo de 3 pasos; paso 1 con Tarifa de
instalación (+ Fecha de ejecución en cambio, + Alcance en rectificar), paso 2 con pestañas
"Servicio N" y "Valor del servicio" (spinbutton) en cada una. `editar_tres_pasos` ya
recorre las pestañas. Reglas SIRT vistas: cada servicio <= valor de su tarifa base y
valor del plan (suma) < suma de las bases → al SUBIR se deja sin subir el último
servicio (el paquete se creó con base − 1). Alcance: debe coincidir con la intersección
de las tarifas base → no se toca. Cambio: alza → hoy + 2 días (igual que monoproducto).

## Pendientes

- **Licitaciones y Contratos (16/10: se implementará después, no ahora).** Registro
  de tarifas de tipo Licitación y Contratos (combinan hasta 9 servicios en una tarifa).
  Base: `referencia_lyc/` (proyecto anterior, ambiente sirt.pe; sin disponibilidad ni
  registro definitivo). Falta grabarlo en QA para FLOR TV.
- **Promociones: catálogo completo (16/10).** 145 atributos promocionables en 14
  servicios (OS: 0), de 4 tipos: PROMOCIONAL (tipo de variación + variación + duración),
  PROMOCIONAL_VELOCIDAD (filas de la tabla de velocidades), PROMOCIONAL_ADJUNTO (canales
  TV, Excel) y PROMOCIONAL_APP (plataforma). `promocionar_fila` maneja los 4; probados 45
  (`ATRIBUTOS_PROBADOS`). `--atributos todos` usa los 145 (sin probar en SIRT).

- **Tablas en cambio / rectificación (16/10, mapeado):** `--todo` no toca filas de tablas
  (Tecnologías y velocidades, LDI, roaming, bolsas de datos): se editan con el botón de
  cada fila (ventana). Falta grabar en un Cambio de Tarifa móvil la edición de una fila
  de "Tecnologías y velocidades" (y si se puede una de países) hasta el resumen. Si la
  ventana es la de "Agregar" del registro, se reutiliza ese código; vale para cambio y
  rectificación (mismo flujo `editar_tres_pasos`).

- **Para después:** opción tipo `--hasta-resumen` (llegar al último paso y
  detenerse sin guardar/confirmar) en todos los scripts: ejecutar, empaquetar,
  promocionar, registro definitivo y acciones. Hoy solo la tiene fe de erratas.

- ~~Agregar `--diagnostico` a `ejecutar.py`~~ **Hecho (13/10/2026):** capturas por
  paso, mapa (aria snapshot) + error.txt al fallar, `--trace` opcional (solo se
  guarda en casos fallidos); la etapa del fallo va al inicio de `detalle`.
  Probado en SIRT real con Telefonía fija (`--diagnostico` y `--trace`, 2 casos
  EXITOSO, registros 26194 y 26195). Falta ver un caso con fallo real.
- **Para después (diagnóstico):** las capturas `full_page=True` solo muestran lo
  visible, porque SIRT desplaza el contenido dentro de un panel interno y no en
  la página. Opción: tomar varias capturas bajando por ese panel (o ampliar su
  alto antes de capturar). Se pospuso para no alargar cada caso.
- Los 15 servicios tienen su archivo (14/10). Pendientes de probar con
  `--limite 1 --diagnostico`: movil_paquete, movil_otros, arrendamiento_circuitos.
- Antes de lanzar lotes completos: decidir qué listas se fijan y cuáles se
  combinan en cada servicio (ver notas de cada archivo: Internet fijo,
  móvil postpago/control/otros tienen listas fijadas para no pasar de ~100 casos).
- Al escribir archivos en la computadora de Bryan, verificar después que
  quedaron iguales (el 13/10 `comun.py` y `ejecutar.py` quedaron con una
  versión anterior y móvil no cargaba). Fase actual: que cada servicio funcione con
  `--limite 1` o `--limite 2`; los lotes completos se lanzarán más adelante.
- Automatizar el GET del JSON para obtener las listas de cada servicio.
- Estimación de tiempo por caso (ya diseñada en una versión previa; no aplicada).
- Probar Internet fijo con Internet Ilimitado = "No" (llena Cantidad de datos).
- Revisar borradores sueltos que quedaron de pruebas fallidas (ver `id_registro`).
- Seguridad: los archivos de grabación del proyecto anterior
  (`grabacion_renomb.py`, `grabacion_renombrar.py`) contienen una contraseña en
  texto plano; cambiarla y borrarla de esos archivos.

## Comandos que leen los Excel de automatizacion/ (19/10/2026, sin probar en SIRT)

La web quedó congelada; ahora cada Excel de `automatizacion/` tiene su comando. Todos aceptan
`--solo-listar` (valida y muestra, sin tocar SIRT ni el Excel), `--nombres`, `--limite`, `--si`,
`--diagnostico`; escriben Estado / Código SIRT / Fecha / Detalle en la fila y dejan una copia en
`resultados/<X>_EXCEL/`. Lo registrado va al lote **EXCEL** (`resultados/<código>/resultados_EXCEL.csv`);
una fila ya registrada no se repite («Ya registrada»; `--repetir` para forzar). Piezas comunes en
`excel_comun.py`. Si Edge no responde (9222) se detienen sin marcar errores.

| Excel | Comando | Usa |
|---|---|---|
| registro_tarifas_establecidas.xlsx | `registro_excel.py` (+ `--hojas IF`, `--sin-definitivo`) | código de `servicios/` con los valores de la fila; tablas de las hojas «Tablas – …» |
| empaquetados.xlsx | `empaquetados_excel.py` | `empaquetar.py` (nuevas opciones: --alcance, fechas, --instalacion, --otros-datos, contacto) |
| registro_licitaciones.xlsx | `licitaciones_excel.py` | `licitacion.py` (lee también el formato nuevo: incluir/detalle por servicio, inicio, instalación) |
| promociones.xlsx | `promociones_excel.py` | `promocionar.py --config <json>` (fechas, textos, valor del plan, instalación, atributos por servicio) |
| modificaciones.xlsx | `modificaciones_excel.py` (+ `--hasta-resumen`) | `acciones.py rectificar|cambio` (nueva opción `--sin-canales`) |
| acciones_basicas.xlsx | `acciones_excel.py` | `acciones.py fe_erratas|renovar|cese|renombre` |

Registro definitivo: si la fila dice Sí, después de registrar se llama a `registro_definitivo.py
--lote EXCEL --nombre N --ubigeos <Disponibilidad comercial>` (vacío = Lima; paquetes: la intersección).
`registro_excel.py` anota el valor del servicio en `resultados/valores_tarifas.csv`, y `empaquetar.py`
lo usa para las reglas de precio (antes tomaba el valor fijo del código).
En registro: Enlace/Fono/Email se reemplazan solo si la fila los trae (SIRT los precarga); listas
vacías → primera opción del código (aviso); países y bolsas = solo las filas del Excel; velocidades
móviles sin filas → las del código. `comun.elegir_fecha` ahora avanza mes a mes para fechas futuras.

## Pendientes anotados el 20/10/2026

- Fechas pasadas: hoy los Excel rechazan «Inicio de vigencia» / «Inicio comercialización» anteriores a hoy
  (registro_excel, empaquetados_excel, promociones_excel, licitaciones_excel). Revisar si SIRT las permite y quitar el bloqueo.
- Windows: Edge minimizado detiene el proceso (las opciones --disable-backgrounding... no bastan). Regla de uso: no minimizar; taparlo u otro escritorio.
- Empaquetados de prueba (automatizacion/empaquetados.xlsx, TE-EMP-CM-101..201): 96 de 4 (TF×TV×IF×SM), 1 de 6, 4 de una sola familia.
  IF + PIF queda en Ejecutar = No: PIF vale 0 y la regla de precios no se puede cumplir. 6 paquetes con OTF/PTF + PIF + SMO usan alcance
  Residencial + Comercial (SMO-005 vale 10).

- Registro definitivo EN BLOQUE (mapeado 20/10, sin implementar): tandas de 10-20. Filtro avanzado por la parte común
  del nombre (p. ej. TE-EMP-CM-1; confirmar si SIRT busca por «contiene»), paginador a 100 filas, leer el JSON de
  POST /api/tarifa/filtrar-temporal (nombre, estado, paso, disponibilidad) para saber qué hay y en qué página;
  marcar solo nombres de la lista con «Ingreso confirmado»; antes de confirmar exigir «Tarifas seleccionadas» = marcadas
  (si no, cancelar); disponibilidad en bloque (botón superior: FALTA GRABARLO); al final un filtro para ver cuáles
  siguen en temporales → esas se reintentan 1x1. Estimado ~1,5 s/tarifa (hoy 15-20 s). Medio día + pruebas.
- Registro definitivo 1x1: se puede quitar la recarga cuando SIRT ya volvió solo al listado (comprobando que no quede
  ninguna fila marcada): ~3 s menos por tarifa. Mantener la verificación final (que ya no esté en temporales).
- Si una tarifa no aparece en temporales, buscarla en «Tarifas registradas» y marcarla «ya tenía registro definitivo»
  (caso TE-EMP-CM-106, cortado a mitad del registro). Empaquetados: «Código SIRT» vacío (no se lee el n.º del borrador)
  y avisos del resumen («no muestra el bloque de …»): revisar con una captura del resumen de un paquete.
- Paralelo (mapeado): elegir puerto por variable (SIRT_PUERTO), un perfil de Edge por proceso, resultados por proceso
  y luego juntarlos en el Excel. Antes probar si SIRT permite dos sesiones del mismo usuario.

## Cantidades de prueba acordadas (20/10/2026) — base para la guía de uso

- Registro de establecidas (registro_tarifas_establecidas.xlsx): 149 = 9–10 por servicio (móviles 10).
- Empaquetados (empaquetados.xlsx, TE-EMP-CM-101..201): 96 de 4 (TF×TV×IF×SM), 1 de 6, 4 de una familia (IF+PIF en No).
- Licitaciones (registro_licitaciones.xlsx, TL-LYC-CM-101..148): 48 = 6 de 1 + 20 de 3 + 15 de 4 + 6 de 5 + 1 de 6.
- Promociones (promociones.xlsx, TP-…-CM-101..155): 55 = 15 servicios × 3 (preferido Porcentual / Cantidad / Ilimitado)
  + 2 de 15 monoproductos + 2 de 3 mono + 3 empaquetados + 6 de un empaquetado. Todas con «Todos» los atributos.
- Acciones básicas (acciones_basicas.xlsx, automatizacion/generar_acciones_prueba.py): 81.
  · Monoproductos (15 servicios): 1.ª sin acción, 2.ª renombre, 3.ª cese, 4.ª fe de erratas (renombre y cese saltan
    a otra tarifa si es base de un paquete o promoción).
  · Empaquetados: 1 renombre, 1 cese, 4 fe de erratas (4 serv., 6 serv., solo TF, solo móvil).
  · Promociones: fe de erratas 6, renovación 6, renombre 6 (2 de monoproducto, 2 de empaquetado, 2 mixtas cada una).
  · Licitaciones: fe de erratas 12 (las 6 de un servicio + 6 de varios).
  · Correr el día siguiente al registro definitivo (fe de erratas / cese / renovación exigen «En comercialización»).
- Pendiente después: guía de uso para quienes van a usar el proyecto (con estas cantidades y tiempos medidos).

## Trabajo en paralelo (mapeado, no se usa por ahora)
- Decisión 21/10/2026: se trabaja con UNA sola pestaña de SIRT; el paralelo queda disponible.
- Otro navegador/perfil (puerto 9223, perfil sirt2) pide iniciar sesión de nuevo: no sirve.
- Pestañas duplicadas del mismo Edge comparten la sesión: SÍ sirve.
- Cómo: duplicar la pestaña de SIRT en su propia ventana (ambas visibles, sin minimizar);
  en el segundo cmd `set SIRT_PESTANA=2` (comun.obtener_pagina_sirt). Sin la variable = pestaña 1.
- Pendiente: probarlo con --limite 1 y que cada proceso trabaje tarifas distintas.
- iniciar_paralelo.bat / SIRT_PUERTO=9223 quedan obsoletos para este fin.

## Registro definitivo en bloque (decidido 21/10, pendiente de grabación)
- Flag `--en-bloque` en registro_definitivo.py; por defecto sigue el modo 1x1 actual (respaldo).
- Disponibilidad: diálogo con «Agregar» (buscar → check → Agregar, ~3-4 s c/u) y un solo registro
  para todas; contar que la lista agregada = las del Excel. Un grupo por disponibilidad distinta.
- Definitivo: recorrer temporales (paginador 100, posiciones desde el JSON) y marcar solo nombres
  exactos del Excel; «Tarifas seleccionadas» = marcadas o cancelar; página por página.
- SIRT registra las demás aunque una falle: leer resumen Exitosos/Fallidos, re-filtrar temporales,
  las que sigan ahí → reintento 1x1. Códigos SIRT desde el JSON del listado.
- El filtro por nombre de SIRT busca por «contiene» (confirmado por el usuario).
- Grabación: grabaciones/_registro_definitivo_bloque/ (notas.txt con lo que hay que grabar).

## Entregables para los analistas (acordado 21/10; hacer AL TERMINAR la prueba funcional)
- Claude: requirements.txt, guía de instalación (Windows/cmd/Edge), manual de usuario, hoja de referencia
  rápida, preguntas frecuentes y errores comunes, limitaciones conocidas, nota de mantenimiento,
  documentación técnica, CHANGELOG + tag v1.0, plantillas vacías + Excel de ejemplo, diapositivas,
  matriz de cobertura (de verificar_funcional.py), guion de cada video (3-5 min por módulo).
- Usuario: grabar los videos y la sesión de entrega en vivo.
- Usar tiempos reales medidos (p. ej. registro definitivo en bloque vs 1x1).

## Entregables v1.0 (hechos 22/10/2026)
- README.md (índice), CHANGELOG.md, requirements.txt fijado (playwright 1.64.0, openpyxl 3.1.5).
- docs/01..09: instalación, manual, referencia rápida, FAQ, limitaciones, mantenimiento, técnica, pruebas, guiones de video.
- entrega/: plantillas vacías (6), ejemplos (serie EJ, 6), matriz_de_cobertura.xlsx (P4 + P5: 268/268).
- Diapositivas: artifact «Automatización de tarifas SIRT v1.0» (12 diapositivas).
- Pendiente del usuario: videos, sesión de entrega, tag v1.0 en git, reintento de 3 filas de Prueba3.
