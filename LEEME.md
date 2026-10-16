# Combinaciones de tarifas SIRT

Tarifa establecida / Monoproducto. Crea una tarifa confirmada por cada
combinación de valores definida para un servicio.

## Estructura

```
sirt_tarifas/
├── comun.py            Lo común a todos: conexión, listas, paso 1, tabla de
│                       velocidades, confirmación, combinaciones y resultados.
├── ejecutar.py         Ejecutor único (un solo punto de entrada).
├── servicios/
│   ├── _plantilla.py   Punto de partida para un servicio nuevo.
│   └── internet_fijo.py
├── grabaciones/        Grabaciones crudas del Inspector, como referencia.
└── resultados/<CODIGO>/
    ├── plan_<lote>.csv
    ├── resultados_<lote>.csv
    ├── capturas_<lote>/
    └── diagnostico_<lote>/<correlativo>/   (solo con --diagnostico)
```

## Antes de ejecutar

1. Abrir Edge en modo controlable:
   `microsoft-edge --remote-debugging-port=9222 --user-data-dir="$HOME/perfiles/sirt" &`

   Para seguir usando el equipo mientras corre, abrirlo además con estas
   opciones (evitan que Edge frene la página cuando no está a la vista):
   ```
   microsoft-edge --remote-debugging-port=9222 --user-data-dir="$HOME/perfiles/sirt" \
     --disable-backgrounding-occluded-windows --disable-renderer-backgrounding \
     --disable-background-timer-throttling &
   ```
   y llevar la ventana a otro escritorio (Super + Shift + Re Pág). No
   minimizarla ni hacer clic dentro mientras corre. `ejecutar.py` ya no trae
   la ventana al frente (salvo con `--al-frente`).
2. Iniciar sesión en SIRT y quedarse en **Tarifas registradas**.
3. Activar el entorno virtual.

## Uso

```bash
python ejecutar.py                                 # lista los servicios
python ejecutar.py internet_fijo --solo-calcular   # solo cuenta y guarda el plan
python ejecutar.py internet_fijo --limite 2        # prueba con 2 casos
python ejecutar.py internet_fijo                   # todos los pendientes (pide confirmación)
python ejecutar.py internet_fijo --si              # todos, sin pedir confirmación
```

Sin `--limite`, el programa muestra cuántas tarifas va a crear y pide escribir
`si` antes de empezar: así un lote completo nunca se lanza por accidente.

Con el mismo `--lote` (por defecto, la fecha del día), una nueva ejecución
omite los casos que ya quedaron como EXITOSO.

En `resultados_<lote>.csv`, la columna `detalle` de un ERROR empieza con la
etapa en que falló, por ejemplo `[paso_2_completo] ...`.

## Registro definitivo

Después de crear las tarifas, pasan a "Tarifas temporales". Para hacer su
registro definitivo (con disponibilidad comercial cuando el servicio la pide):

```bash
python registro_definitivo.py --lote MUESTRA --solo-listar            # qué haría, sin tocar SIRT
python registro_definitivo.py --lote MUESTRA --limite 1 --diagnostico # prueba con 1
python registro_definitivo.py --lote MUESTRA --servicio SMP --limite 1
python registro_definitivo.py --lote MUESTRA                          # todas (pide confirmación)
python registro_definitivo.py --lote MUESTRA --ubigeos "Amazonas,Ancash"
```

Empaquetadas (`--servicio EMP`): piden disponibilidad solo si alguna de sus
tarifas la requiere, y sus ubicaciones deben estar en la intersección de las de
esas tarifas (las que no requieren disponibilidad no cuentan). Sin `--ubigeos`
se usa toda la intersección; con `--ubigeos`, cada ubicación debe estar en ella.
Si no hay intersección, el paquete se lista como "No se procesarán".

Procesa todos los servicios del lote en una sola ejecución. El registro
definitivo probablemente no se puede deshacer: probar primero con `--limite 1`.
Resultados en `resultados/registro_definitivo_<lote>.csv` (reanudable).

## Tarifa empaquetada

Arma una tarifa "Empaquetada" con tarifas que ya tienen registro definitivo:

```bash
python empaquetar.py --tarifas TE-TFP-MUESTRA-001 TE-SMP-MUESTRA-001 --lote MUESTRA --solo-listar
python empaquetar.py --tarifas TE-TFP-MUESTRA-001 TE-SMP-MUESTRA-001 --lote MUESTRA --diagnostico
python empaquetar.py --tarifas ... --valores 29 49          # valor por servicio, en el mismo orden
```

Antes de tocar SIRT revisa: al menos 2 tarifas, que existan, que tengan
registro definitivo, mismo alcance, sin servicios repetidos y que la suma de
valores (= "Valor del plan") cumpla: valor de la tarifa más cara ≤ suma <
suma de las tarifas base. Por defecto cada servicio lleva el valor de su
tarifa menos 1 (30 y 50 → 29 y 49). En el paso 2 cada pestaña
"Servicio N" se reconoce por el nombre del servicio (el orden no coincide con
el de las tarifas) y se pulsa Guardar en cada una. Nombre:
`TE-EMP-<lote>-NNN`; resultados en `resultados/EMP/resultados_<lote>.csv`.

## Tarifa promocional

Promociona una o más tarifas con registro definitivo (monoproductos o paquetes):

```bash
python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --solo-listar
python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --diagnostico
python promocionar.py --tarifas TE-EMP-MUESTRA-006 --lote PROMO --diagnostico
python promocionar.py --tarifas TE-IF-PROMO-001 --lote PROMO --atributos todos --solo-listar
```

Paso 1: Valor del plan (y Tarifa instalación si el servicio la tiene). Paso 2:
una pestaña por servicio; cada atributo se marca y lleva Tipo de variación
(Porcentual > Cantidad > Ilimitado), Variación (10 % o 1) y Duración (30 días).
`--atributos probados` (por defecto: los de la grabación), `todos` (los del
JSON del servicio en `grabaciones/_promocional/<servicio>/paso2.json`) o una
lista `"A;B"`. Nombre `TP-<código>-<lote>-NNN`; resultados en
`resultados/PRM/resultados_<lote>.csv`.

## Acciones sobre tarifas registradas

```bash
python acciones.py buscar --tarifas TE-IF-MUESTRA-001 TE-SMO-PROMO-001   # estados, sin cambiar nada
python acciones.py renombre --tarifas TE-SMO-PROMO-002 --solo-listar
python acciones.py renombre --tarifas TE-SMO-PROMO-002 --diagnostico      # nombre + "_REN"
python acciones.py renombre --tarifas TE-AC-PROMO-001 --nuevo-nombre TE-AC-PROMO-001_NUEVO
python acciones.py cese --tarifas TE-TF-MUESTRA-001 --solo-listar           # fin de comercialización mañana
python acciones.py cese --tarifas TE-TF-MUESTRA-001 --fecha 2026-10-20 --diagnostico
```

Cese de comercialización: valida antes (del JSON del filtro) que la tarifa esté
"En comercialización" y vigente; si no, no la toca (`--sin-validar` para forzar).
Pide confirmación salvo `--si`.

Fe de erratas: agrega "| Fe de erratas <fecha hora>" a todos los campos de texto
visibles de los pasos 1 y 2 (sirve para cualquier servicio; solo cambia cuántos
campos hay), comprueba el resumen del paso 3 y pulsa "Guardar cambios".
Si el paso 2 tiene "Tipo de Contratación" (PTV), también la cambia sin dejarla
vacía: con Postpago + Prepago desmarca la última; con una sola, marca la otra
(`--sin-listas` para no tocarla; aplica también a rectificar).
`--hasta-resumen` se detiene en el paso 3 sin guardar.
Casos por servicio (`grabaciones/_acciones/fe_erratas/pendientes.txt`): caso1 fijos
(TF, IF, PTF, OTF, TVP, TVPR, PIF), caso2 OS/AC, caso3 móviles (SMP, SMC, SMO, SMQ),
caso4 PTV (con Tipo de Contratación).
```bash
python acciones.py fe_erratas --tarifas TE-IF-MUESTRA-001 --hasta-resumen --diagnostico
python acciones.py fe_erratas --tarifas TE-IF-MUESTRA-001 --diagnostico
```

Rectificar: mismo flujo que fe de erratas (menú "Rectificar", pasos 1-2-3,
"Guardar cambios", dos "Aceptar"); además suma `--delta` (1) al Valor del
servicio y a la Tarifa de instalación si está habilitada. Si el campo sigue
vacío tras esperar a que cargue, pone 1 (o `--valor-si-vacio N`). No toca
tablas ni el alcance.
Las tarifas con la acción ya EXITOSA en el CSV se saltan; `--repetir` lo evita
(p. ej. una segunda rectificación).
```bash
python acciones.py rectificar --tarifas TE-OS-MUESTRA-001 --hasta-resumen --diagnostico
python acciones.py rectificar --tarifas TE-OS-MUESTRA-001 --diagnostico
```

Cambio de tarifa (menú "Cambio de Tarifa"): mismo flujo de 3 pasos + "Fecha de
ejecución" en el paso 1. Por defecto BAJA 1 el valor del servicio y la instalación
(`--delta -1`, beneficioso) con fecha hoy + `DIAS_ANTICIPACION_BAJA` (1: el calendario siempre bloquea hoy;
si el día está bloqueado prueba los siguientes, `DIAS_EXTRA_SI_BLOQUEADO`). Si sube
(`--delta 1`), la fecha es hoy + `DIAS_ANTICIPACION_ALZA` (2; regla de SIRT). Ambas
constantes están al inicio de `acciones.py`. `--fecha AAAA-MM-DD` la fija a mano.
```bash
python acciones.py cambio --tarifas TE-OS-MUESTRA-001 --hasta-resumen --diagnostico
python acciones.py cambio --tarifas TE-OS-MUESTRA-001 --diagnostico
python acciones.py cambio --tarifas TE-OS-MUESTRA-001 --delta 1 --diagnostico
```
`--todo` (cambio / rectificar): además cambia TODOS los demás campos numéricos del
paso 2 (datos, velocidades, minutos/SMS, tarifas por minuto, roaming...) en 1 unidad de
su último decimal (12 → 13, 0.10 → 0.11), sin tocar los que valen 0 ni pasar de 100 en
%, y alterna "Apps sin consumo de datos". Con "Datos tope descarga" primero al subir y
último al bajar (tethering / traslada / roaming ≤ tope). Las filas de tablas no se tocan.

Busca por nombre en Tarifas registradas (Filtro avanzado) y lee código SIRT,
versión, comercialización y vigencia de la respuesta del filtro. Renombre:
⋮ Acciones → Renombre → nombre nuevo → Guardar → Guardar tarifa → Aceptar, y
verifica buscando el nombre nuevo. Resultados en
`resultados/ACCIONES/acciones.csv` (salta las ya hechas). se agregarán aquí; exigirán "En comercialización" y vigente.

## Diagnóstico

```bash
python ejecutar.py telefonia_fija --limite 1 --diagnostico   # capturas por paso + mapa al fallar
python ejecutar.py telefonia_fija --limite 1 --trace         # lo anterior + trace de Playwright
```

En `resultados/<CODIGO>/diagnostico_<lote>/<correlativo>/` queda:

| Archivo                          | Cuándo                | Para qué                                  |
|----------------------------------|-----------------------|-------------------------------------------|
| `01_abrir_nueva_tarifa.png` … `06_confirmada.png` | cada paso terminado | ver hasta dónde llegó y cómo quedó |
| `NN_FALLO_<etapa>.png`           | si falla              | pantalla en el momento del error          |
| `NN_FALLO_<etapa>_mapa.yaml`     | si falla              | roles y nombres que ve Playwright (aria snapshot): si un botón o lista no aparece con el nombre esperado, `get_by_role` no lo encontrará |
| `NN_FALLO_<etapa>_error.txt`     | si falla              | etapa, URL, error y traceback completo    |
| `trace.zip`                      | si falla y `--trace`  | `playwright show-trace trace.zip`: paso a paso con DOM, red y consola |

Etapas: `abrir_nueva_tarifa`, `paso_1_completo`, `paso_1_guardado`,
`paso_2_completo`, `paso_2_guardado`, `confirmada`.

Las capturas por paso hacen cada caso algo más lento: usar el diagnóstico con
`--limite` al probar un servicio nuevo o al investigar un error, no en lotes largos.
Los traces de casos exitosos se descartan.

## Agregar un servicio

1. Grabar el paso 2 con `grabar_sirt.py` y guardar la grabación en `grabaciones/`.
2. Obtener las opciones de sus listas (capturas o el JSON del GET).
3. Copiar `servicios/_plantilla.py` como `servicios/<nombre>.py` y completarlo.
4. `python ejecutar.py <nombre> --solo-calcular`, luego `--limite 1`.

Si algo sirve para varios servicios (por ejemplo, otra tabla emergente),
se agrega como función en `comun.py` y cada servicio la llama.

## Lecciones aprendidas

- Las opciones se buscan con texto exacto: "Si" va sin tilde.
- La ventana de velocidades no es un `dialog`: se detecta por su primer campo.
- "Confirmar Tarifa" (resumen) y "Confirmar tarifa" (ventana) son botones distintos.
- Guardar el paso 1 ya crea un borrador; su número queda en `id_registro`.
- Los textos de las listas cambian entre servicios: "Alta Nueva" (Internet fijo)
  vs "Alta nueva" (Telefonía fija), "Tecnología (IF)" vs "Tecnologia" sin tilde.
  Tomarlos siempre del JSON del servicio.
- Las tecnologías disponibles cambian por servicio.
- Bloques "condicional triple" (minutos): si "Es ilimitado" está marcado, los
  dos campos se dejan vacíos; si no, se llenan los dos.

## Bloques reutilizables en comun.py

| Función                         | Para qué                                          |
|---------------------------------|---------------------------------------------------|
| `agregar_velocidad`             | Tabla de velocidades (Tecnología + 4 velocidades) |
| `completar_condicional_triple`  | Fila "Es ilimitado" + minutos + tarifa            |
| `agregar_tarifa_paises`         | Tabla de tarifas por países (Tipo, Tarifa, Países)|
| `elegir_valor`                  | Lista única o múltiple según el valor del caso    |
| `llenar_si_editable`            | Campo que a veces está bloqueado                  |
| `completar_textos_finales`      | Condiciones generales + Restricciones técnicas    |
| `cargar_archivo`                | Campo "Carga y descarga" (adjunta un Excel y espera que SIRT muestre su nombre) |
| `agregar_bolsa_datos`           | Tabla "Bolsa de gigas" (Tipo, Gigas, Plataformas múltiple) |
| `boton_agregar`                 | Botón "Agregar" por posición (int) o por título de la sección (str) |

Cuando una página tiene muchas tablas (Móvil), las funciones de tablas reciben
el título de la sección en vez del número de botón:
`agregar_tarifa_paises(page, "Tarifa SMS roaming", fila)`.

Los archivos a cargar se guardan en `grabaciones/<servicio>/` y el servicio
los referencia con su ruta (ver `servicios/tv_paga_postpago.py`).
