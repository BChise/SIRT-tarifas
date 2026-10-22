# Guía de uso — Automatización de tarifas SIRT

Versión 1.0 · octubre de 2026

Esta guía es para quienes **registran y modifican tarifas en SIRT** con la herramienta. No hace falta saber programar
ni conocer la consola de Windows: se llena un Excel, se escribe **un comando** y la herramienta hace el trabajo en
SIRT. Al terminar, el mismo Excel muestra el resultado de cada fila.

> En esta guía, **`<carpeta del proyecto>`** es la carpeta donde se instaló la herramienta (la que contiene
> `iniciar.bat`). Puede estar en cualquier lugar, por ejemplo `D:\sirt_tarifas` o `C:\Herramientas\sirt_tarifas`.

---

## 1. Qué hace la herramienta

| Tarea | Excel que se llena | Comando |
|---|---|---|
| Registrar tarifas establecidas (monoproductos) | `registro_tarifas_establecidas.xlsx` | `python registro_excel.py --si --definitivo-en-bloque` |
| Registrar empaquetados | `empaquetados.xlsx` | `python empaquetados_excel.py --si --definitivo-en-bloque` |
| Registrar licitaciones y contratos | `registro_licitaciones.xlsx` | `python licitaciones_excel.py --si --definitivo-en-bloque` |
| Registrar promociones | `promociones.xlsx` | `python promociones_excel.py --si --definitivo-en-bloque` |
| Fe de erratas, renovación, cese y renombre | `acciones_basicas.xlsx` | `python acciones_excel.py --si` |
| Rectificación y cambio de tarifa | `modificaciones.xlsx` | `python modificaciones_excel.py --si` |

Los seis Excel están en `<carpeta del proyecto>\automatizacion\`. Cada fila del Excel es una tarifa (o una acción
sobre una tarifa).

Al registrar, la herramienta también hace la **disponibilidad comercial** (cuando el servicio la pide) y el
**registro definitivo**.

**Qué significan las partes del comando**

- **`--si`:** empieza a trabajar sin preguntar. Sin esta opción, la herramienta muestra cuántas filas va a procesar y
  pide escribir `si` para confirmar. Es una protección para no lanzar un Excel completo por accidente.
- **`--definitivo-en-bloque`** (solo en los cuatro registros): hace la disponibilidad comercial y el registro
  definitivo de **varias tarifas a la vez** (de 10 en 10) en lugar de una por una. Toma 1 a 2 segundos por tarifa en
  vez de 15 a 22. Solo marca las tarifas del Excel y comprueba que SIRT cuente las mismas antes de confirmar. Si alguna
  falla, se reintenta sola de una en una. Sin esta opción, el registro definitivo se hace de una en una, más lento
  pero igual de válido.

Las demás opciones están en la sección 8.

---

## 2. Cómo se trabaja: en dos días

SIRT solo permite hacer acciones sobre una tarifa cuando ya está **«En comercialización»** y **«Vigente»**, y eso
ocurre **el día siguiente** a su registro definitivo. Por eso el trabajo se reparte en dos días:

| | Qué se hace | Excel, en este orden |
|---|---|---|
| **Día 1 — Registrar** | Crear las tarifas con su registro definitivo | 1. `registro_tarifas_establecidas.xlsx` → 2. `empaquetados.xlsx` y `registro_licitaciones.xlsx` → 3. `promociones.xlsx` |
| **Día 2 (o después) — Modificar** | Acciones y modificaciones sobre tarifas ya vigentes | 4. `acciones_basicas.xlsx` → 5. `modificaciones.xlsx` |

- **El orden importa:** un empaquetado necesita que sus tarifas ya estén registradas; una promoción necesita que sus
  tarifas base (monoproductos o empaquetados) también lo estén.
- **El día 2 no tiene que ser exactamente el día siguiente:** puede ser cualquier día posterior.
- Si una acción se intenta antes de tiempo, la fila queda **«No aplica»** y no se toca SIRT; basta con correrla otro
  día.
- Las licitaciones solo necesitan estar vigentes.
- Solo se usan los Excel que hagan falta: si hoy solo hay tarifas establecidas, se usa solo ese.

### 2.1 Cuánto demora cada Excel

Tiempos medidos en las pruebas funcionales en SIRT QA, con `--definitivo-en-bloque`. Incluyen todo lo que hace la
herramienta en cada fila: registro, disponibilidad comercial y registro definitivo.

| Día | Excel | Por fila | 10 filas | 50 filas |
|---|---|---|---|---|
| 1 | Tarifas establecidas, servicios fijos | ~15 s | ~3 min | ~13 min |
| 1 | Tarifas establecidas, servicios móviles | ~18 s | ~3 min | ~15 min |
| 1 | Empaquetados | ~30 s | ~5 min | ~25 min |
| 1 | Licitaciones y contratos | 25–40 s | ~5 min | ~30 min |
| 1 | Promociones | 25–45 s | ~6 min | ~30 min |
| 2 | Acciones básicas (fe de erratas, renovación, cese, renombre) | 10–15 s | ~2 min | ~10 min |
| 2 | Rectificación y cambio de tarifa | 15–20 s | ~3 min | ~15 min |

- **Arranque:** se suma alrededor de 1 minuto por comando (conectar con Edge, revisar el Excel, guardar al final).
- **Sin `--definitivo-en-bloque`:** el registro definitivo de una en una suma 15 a 22 s por tarifa. Por ejemplo, 129
  tarifas establecidas tardaron **1 h 15 min** de una en una; en bloque serían unos **35 min**.
- **Ejemplo real del día 1:** 80 filas (52 establecidas, 10 empaquetados, 6 licitaciones y 12 promociones) en
  **26 min**.
- **Ejemplo real del día 2:** 15 acciones básicas en **~3 min** y 23 rectificaciones y cambios en **~8 min**.
- Son tiempos de referencia: si SIRT está lento, todo puede tardar el doble. Si la herramienta se detiene, lo ya hecho
  no se repite al volver a correr el comando.

**Ejemplo de una carga grande (como la serie de pruebas CM)**

Supuesto: 13 tarifas por cada uno de los 15 servicios (194 establecidas: 142 de servicios fijos y 52 móviles), 100
empaquetados, 48 licitaciones y 55 promociones. El día 2, cada tarifa establecida y cada empaquetado recibe una
acción (294 acciones), mitad acciones básicas y mitad rectificaciones o cambios.

| Día | Excel | Filas | Con `--definitivo-en-bloque` | De una en una |
|---|---|---|---|---|
| 1 | Tarifas establecidas | 194 | ~50 min | ~1 h 55 min |
| 1 | Empaquetados | 100 | ~50 min | ~1 h 20 min |
| 1 | Licitaciones y contratos | 48 | 20–30 min | 35–45 min |
| 1 | Promociones | 55 | 25–40 min | 40–55 min |
| | **Total día 1** (con arranques) | **397** | **~2 h 30 min a 3 h** | **~4 h 30 min a 5 h** |
| 2 | Acciones básicas | 147 | 25–40 min | — |
| 2 | Rectificación y cambio de tarifa | 147 | 40–50 min | — |
| | **Total día 2** (con arranques) | **294** | **~1 h 05 min a 1 h 30 min** | — |

- El día 2 no usa registro definitivo, así que la opción en bloque no cambia su tiempo.
- Si todas las acciones del día 2 fueran básicas, serían unos 50 min a 1 h 15 min. Si todas fueran rectificaciones o
  cambios, serían unos 1 h 15 min a 1 h 40 min.
- Si también se modifican las licitaciones, se suman 15 a 20 s por cada una. Las promociones no se pueden rectificar
  ni cambiar en SIRT.

**Ejemplo para un solo servicio (una tarifa por cada acción)**

Es el mismo formato de la serie CM, pero para un servicio (por ejemplo, Internet fijo):
- 13 tarifas establecidas;
- su parte de empaquetados (7), licitaciones (3) y promociones (4);
- el día 2, cada tarifa establecida y cada empaquetado recibe una acción.

| Día | Tipo | Acción | Cantidad | Por tarifa | Total |
|---|---|---|---|---|---|
| 1 | Monoproducto | Registro + definitivo en bloque | 13 | ~15 s (móvil ~18 s) | ~3–4 min |
| 1 | Empaquetado | Registro + definitivo en bloque | 7 | ~30 s | ~3–4 min |
| 1 | Licitación | Registro + definitivo en bloque | 3 | 25–40 s | ~1–2 min |
| 1 | Promoción | Registro + definitivo en bloque | 4 | 25–45 s | ~2–3 min |
| 1 | | Arranque de los 4 comandos | | ~1 min c/u | ~4 min |
| | | **Total día 1** | **27** | | **~14–17 min** |
| 2 | Monoproducto / Empaquetado | Fe de erratas | 4 | ~14 s | ~1 min |
| 2 | Monoproducto / Empaquetado | Renombre | 3 | ~9 s | ~30 s |
| 2 | Monoproducto / Empaquetado | Cese | 3 | ~11 s | ~30 s |
| 2 | Monoproducto / Empaquetado | Rectificación | 4 | 15–20 s | ~1–1,5 min |
| 2 | Monoproducto / Empaquetado | Cambio beneficioso | 3 | 15–20 s | ~1 min |
| 2 | Monoproducto / Empaquetado | Cambio perjudicial | 3 | 15–20 s | ~1 min |
| 2 | | Arranque de los 2 comandos | | ~1 min c/u | ~2 min |
| | | **Total día 2** | **20** | | **~7–8 min** |

- **Total por servicio:** unos **21–25 minutos** entre los dos días. Sin `--definitivo-en-bloque`, el día 1 sube a
  unos 25 minutos.
- Si las promociones también reciben una acción (fe de erratas, renovación o renombre), se suman unos 10–15 s por
  cada una.
- Si se cargan varios servicios en el mismo Excel, el arranque se cuenta una sola vez por comando, no por servicio.

---

## 3. Reglas de oro

1. **Nunca minimizar Edge** mientras la herramienta trabaja. Minimizado, Windows lo congela y el proceso se detiene.
   Sí se puede tapar con otras ventanas o bloquear la computadora (tecla Windows + L).
2. **No hacer clic ni escribir en Edge** mientras trabaja.
3. **El Excel debe estar cerrado** al lanzar el comando, y **no se debe abrir** mientras corre.
4. **Primero se prueba con una fila** la primera vez que se usa un Excel (sección 5).
5. **Si algo se corta** (sesión vencida, Edge cerrado, corte de luz): iniciar sesión de nuevo y **correr el mismo
   comando con el mismo Excel**. Lo que ya se hizo no se repite.
6. **Acciones y modificaciones, desde el día siguiente** al registro definitivo.

---

## 4. Paso a paso de una sesión de trabajo

### 4.1 Abrir la herramienta

1. **Cerrar Edge por completo**, si está abierto.
2. Abrir `<carpeta del proyecto>` en el Explorador de archivos y hacer **doble clic en `iniciar.bat`**. Se abren dos
   ventanas:
   - **Edge**, preparado para que la herramienta lo maneje.
   - Una **ventana negra** (la consola, «cmd»), ya ubicada en la carpeta del proyecto. Ahí se escriben los comandos.
3. En ese Edge, **iniciar sesión en SIRT** (usuario, contraseña, reCAPTCHA y código por correo) y dejar la pantalla en
   **«Tarifas registradas»**.

### 4.2 Preparar el Excel

1. Llenar el Excel en `<carpeta del proyecto>\automatizacion\` (ver la sección 6).
2. **Guardarlo y cerrarlo.**

### 4.3 Escribir el comando

En la ventana negra, escribir el comando del Excel (tabla de la sección 1) y pulsar **Enter**. Por ejemplo:

```
python registro_excel.py --si --definitivo-en-bloque
```

La ventana va mostrando cada fila. Al final aparece:

```
Resumen: Registrada: 12 · Registro definitivo: 12
Tiempo total: 6 min 40 s · promedio por fila: 33 s
Excel actualizado: <carpeta del proyecto>\automatizacion\registro_tarifas_establecidas.xlsx
```

### 4.4 Ver el resultado

Abrir el Excel: cada fila tiene su **Estado**, el **Código SIRT** y, si algo falló, el motivo en **Detalle**
(sección 7).

> **Consejo:** en la ventana negra, la flecha ↑ del teclado recupera el último comando escrito. Así no hace falta
> volver a escribirlo.

---

## 5. Forma de trabajo recomendada (cualquier Excel)

La primera vez que se usa un Excel, o con un Excel muy grande, conviene ir en tres pasos:

| Paso | Comando (ejemplo con el registro) | Qué hace |
|---|---|---|
| 1. Revisar | `python registro_excel.py --solo-listar` | Muestra qué haría y los **errores de datos** (marcados con ✖). **No toca SIRT ni el Excel.** |
| 2. Probar con una | `python registro_excel.py --limite 1 --si` | Procesa solo la primera fila pendiente. Revisar en SIRT que quedó bien. |
| 3. Todo | `python registro_excel.py --si --definitivo-en-bloque` | Procesa todas las filas pendientes. |

Para los demás Excel es igual: se cambia el nombre del comando (tabla de la sección 1). En `modificaciones.xlsx`, el
paso 2 puede hacerse sin guardar nada en SIRT:

```
python modificaciones_excel.py --hasta-resumen --limite 1
```

Con eso la herramienta llega al resumen de SIRT y se detiene sin guardar.

---

## 6. Los Excel

### 6.1 Formato común

- **La fila 5 tiene los títulos** y los datos empiezan en la **fila 6**. No se deben cambiar los títulos ni el orden
  de las columnas.
- **Colores de los títulos:** verde oscuro = identificación · azul oscuro = obligatorio · azul claro = opcional ·
  gris = lo llena la herramienta (no escribir ahí).
- **Listas desplegables:** muestran los valores válidos; conviene elegir de la lista en vez de escribir.
- **Ejecutar:** `Sí` (o vacío) = se procesa; `No` = se salta.
- **Nombre de la tarifa:** debe ser único y coincidir exactamente con el de SIRT cuando la tarifa ya existe (acciones y
  modificaciones).
- Cada Excel trae una **hoja de instrucciones** con el detalle de cada columna.
- **Columnas de resultado** (grises, al final): **Estado**, **Código SIRT** (p. ej. `TEMTF2026000758`), **ID de
  registro** (número interno de SIRT), **Fecha** y **Detalle**.

En la carpeta `entrega\ejemplos\` hay un Excel de ejemplo por tarea. Cada fila muestra una opción distinta y explica
qué hace en la columna «Qué se prueba». Sus nombres de tarifas son de ejemplo: hay que cambiarlos antes de usarlos
en SIRT. En `entrega\plantillas\` están los seis Excel vacíos.

### 6.2 Tarifas establecidas — `registro_tarifas_establecidas.xlsx` (día 1)

- **Una hoja por servicio**, con el código del servicio como nombre: TF, TFP, OTF, PTF, IF, PIF, TVP, TVPR, PTV, SMP,
  SMC, SMO, SMQ, OS y AC. Una fila = una tarifa.
- **Hojas de tablas** («Tablas – Países», «Tablas – Velocidades», «Tablas – Bolsas»): una fila por fila de la tabla
  de SIRT, unida a su tarifa por el **nombre**.
- Columnas principales: Nombre, **Registro definitivo** (Sí/No), **Disponibilidad comercial** (solo en servicios que
  la piden), **Alcance** (Residencial / Comercial / Residencial + Comercial), Inicio de vigencia, Inicio de
  comercialización, datos de contacto, **Valor del servicio** y las listas y cantidades propias del servicio.
- Para varias ubicaciones de disponibilidad se escriben separadas por coma: `Lima, Arequipa`.
- Las fechas solo pueden ser hoy o futuras.

Ejemplo (hoja TF):

| Nombre de la tarifa | Registro definitivo | Disponibilidad comercial | Alcance | Valor del servicio | Periodicidad | Tipo de acceso |
|---|---|---|---|---|---|---|
| TE-TF-EJ-011 | Sí | Lima | Residencial | 50 | Mensual | Alta nueva |

y su fila en «Tablas – Países»:

| Servicio | Nombre de la tarifa | Tabla | Tipo | Tarifa | Países |
|---|---|---|---|---|---|
| TF | TE-TF-EJ-011 | Tarifa por minuto a fijo LDI | Bloque | 12.45 | ALBANIA + BRUNEI |

### 6.3 Empaquetados — `empaquetados.xlsx` (día 1, después del registro)

- Una fila = un paquete de **2 a 6 tarifas** (columnas Tarifa 1 a Tarifa 6), ya registradas y con registro
  definitivo.
- El paquete y sus tarifas deben tener el **mismo alcance**.
- **Valor N en el paquete:** cuánto vale cada tarifa dentro del paquete. Reglas de SIRT: cada valor ≤ el de su tarifa,
  y la suma del paquete menor que la suma de sus tarifas. Vacío = valor de la tarifa − 1, que siempre cumple.
- **Tarifa instalación:** obligatoria si el paquete tiene servicios fijos; se deja vacía solo si el paquete no la usa
  (p. ej. solo móviles).

Ejemplo:

| Nombre del paquete | Alcance | Tarifa instalación | Tarifa 1 | Valor 1 | Tarifa 2 | Valor 2 |
|---|---|---|---|---|---|---|
| TE-EMP-EJ-101 | Residencial | 10 | TE-TF-EJ-011 | *(vacío)* | TE-SMP-EJ-011 | *(vacío)* |

### 6.4 Licitaciones y contratos — `registro_licitaciones.xlsx` (día 1)

- Una fila = una licitación con 1 a 6 servicios. Por cada servicio: **«<servicio>: incluir»** (Sí/No) y
  **«<servicio>: detalle del servicio»**. Debe haber al menos uno incluido.
- Datos del contrato: Inicio de vigencia, Duración (días), Identificador, **Ámbito** (Público / Privado), **Valor de
  contrato**, **Moneda** (Soles / Dólares / Euro) y datos de la empresa.
- «Tarifa instalación» solo se usa con «Otros servicios».

Ejemplo:

| Nombre | Duración (días) | Identificador | Ámbito | Valor de contrato | Moneda | TV de paga: incluir | Internet Fijo: incluir |
|---|---|---|---|---|---|---|---|
| TL-LYC-EJ-103 | 180 | CTR-EJ-103 | Público | 1250.00 | Euro | Sí | Sí |

### 6.5 Promociones — `promociones.xlsx` (día 1, al final)

- Una fila = una promoción sobre **1 a 15 tarifas base** (columnas Tarifa base 1 a 15), con registro definitivo.
- **Valor del plan** (y **Tarifa instalación**): promocionar Sí/No, tipo de variación (Porcentual / Cantidad) y
  variación.
- **Atributos a promocionar:**
  - **Vacío o «Los de siempre»:** los atributos habituales del servicio.
  - **«Todos»:** todos los que SIRT permite.
  - **«Lista»:** solo los escritos en «Lista de atributos», separados por `;`.
  - **«Detalle en hoja Atributos»:** en la hoja «Atributos», una fila por atributo con su tipo, variación y duración.
- Los nombres exactos de los atributos están en la hoja **«Catálogo de atributos»**.
- **Tipo de variación preferido:** Porcentual / Cantidad / Ilimitado. Si un atributo no admite el preferido, se usa
  otro que sí admita.
- Si las tarifas base son de servicios distintos, SIRT solo deja promocionar lo que tienen **en común**.

Ejemplo:

| Nombre de la promoción | Tarifa base 1 | Días de comercialización | Valor del plan: promocionar | Tipo | Variación | Atributos a promocionar |
|---|---|---|---|---|---|---|
| TP-TVP-EJ-104 | TE-TVP-EJ-011 | 60 | Sí | Porcentual | 10 | *(vacío)* |

### 6.6 Acciones básicas — `acciones_basicas.xlsx` (día 2)

| Acción | Sobre qué tarifas | Columna que usa |
|---|---|---|
| Fe de erratas | todas | — (agrega «\| Fe de erratas fecha hora» a los textos) |
| Renovación | solo promociones (nombres TP-…) | Días de comercialización, Duración (días) |
| Cese | monoproductos y empaquetados | Fecha de fin de comercialización (posterior a hoy) |
| Renombre | todas | Sufijo (se agrega al final del nombre) |

- Se ejecutan en este orden: fe de erratas → renovación → cese → renombre.
- Si una tarifa **ya fue renombrada**, las demás acciones encuentran su nombre nuevo solas.
- Para renombrar otra vez una tarifa ya renombrada, se escribe en la fila su **nombre actual**
  (p. ej. `TARIFA_REN`).

Ejemplo:

| Nombre de la tarifa | Acción | Renombre: sufijo | Cese: fecha de fin |
|---|---|---|---|
| TE-IF-EJ-012 | Renombre | *(vacío → `_REN`)* | |
| TE-TFP-EJ-012 | Cese | | *(vacío → mañana)* |

### 6.7 Rectificación y cambio de tarifa — `modificaciones.xlsx` (día 2)

- **Acción:**
  - **Rectificación:** corrige la tarifa vigente y se guarda al momento.
  - **Cambio:** crea una versión nueva que empieza en la **fecha de ejecución**. Solo puede haber un cambio pendiente
    por tarifa.
- **Delta:** cuánto suben o bajan los números (acepta decimales y negativos).
- **Cambio: sentido:**
  - **Beneficioso:** bajan los precios y suben los beneficios (minutos, datos, velocidades…).
  - **Perjudicial:** al revés.
  - **Vacío:** el delta se aplica igual a todo.
- **Campos a modificar:**
  - **Por defecto:** textos (con un sello), valor del servicio, instalación y algunas listas.
  - **Todos los campos:** además, todos los números del servicio y las filas de sus tablas.
  - **Solo la lista:** solo los campos escritos en «Lista de campos», con su nombre tal como en SIRT y separados por
    `;`.
  - **Detalle en hoja Campos:** en la hoja «Campos», una fila por campo, con su **Valor nuevo** o su propio
    **Delta**. Se puede apuntar a una fila de una tabla con `›`; por ejemplo,
    `Tecnologías y velocidades › 4G › Descarga máx. contratada`.
- **Opcionales:** agregar una fila a cada tabla, pasar «ilimitado» a limitado, cambiar canales (TV), texto del sello.
- **Empaquetados:** el valor de los servicios del paquete no sube (regla de SIRT); lo demás sí se modifica.
- **Licitaciones:** solo rectificación.
- **Promociones:** SIRT no permite rectificarlas ni cambiarlas; las filas de promociones (TP-…) quedan como «Error de datos».
- Antes de guardar, la herramienta comprueba que **todos los cambios aparezcan en el resumen** de SIRT; si falta
  alguno, no guarda y lo indica en Detalle.

Ejemplo:

| Nombre de la tarifa | Acción | Delta | Sentido | Campos a modificar | Lista de campos |
|---|---|---|---|---|---|
| TE-TF-EJ-013 | Rectificación | *(vacío → +1)* | | Por defecto | |
| TE-SMP-EJ-014 | Cambio | 1 | Beneficioso | Todos los campos | |
| TE-IF-EJ-013 | Rectificación | 0.5 | | Solo la lista | Valor del servicio; Otros datos generales |

### 6.8 Qué pasa cuando una celda se deja vacía

| Celda vacía | Se usa |
|---|---|
| Ejecutar | Sí |
| Inicio de vigencia / Inicio de comercialización | Hoy |
| Disponibilidad comercial (registro) | Lima |
| Disponibilidad comercial (empaquetado) | La que tienen en común sus tarifas |
| Valor N en el paquete | Valor de la tarifa − 1 |
| Atributos a promocionar | Los de siempre |
| Delta | +1 en rectificación · −1 en cambio |
| Cambio: fecha de ejecución | Hoy + 3 días |
| Cese: fecha de fin de comercialización | Mañana |
| Renombre: sufijo | `_REN` |
| Renovación: días de comercialización | 19 días |
| Renovación: duración | 25 días (31 si ya eran 25) |
| Cambiar canales (TV) | Sí |
| Texto del sello | «Rectificación» o «Cambio» |

---

## 7. Resultados

### 7.1 Qué significa cada estado

| Estado | Significado | Qué hacer |
|---|---|---|
| **Registro definitivo** | Registrada con registro definitivo. | Nada. |
| **Aplicada** | La acción o modificación se guardó en SIRT. | Nada. |
| **Ya registrada / Ya aplicada** | Ya estaba hecha de una vez anterior; no se repitió. | Nada. |
| **Registrada** | Creada en SIRT, todavía sin registro definitivo (la fila decía «No», o falta hacerlo). | Correr de nuevo el comando para completarlo, si se quería. |
| **Hasta resumen** | Prueba sin guardar (`--hasta-resumen`). | Correr sin esa opción para guardar. |
| **No aplica** | La tarifa aún no está vigente o en comercialización. No se tocó SIRT. | Correr otro día. |
| **Error de datos** | Un dato del Excel no es válido. No se tocó SIRT. | Corregir la fila según el Detalle y correr de nuevo. |
| **Error** | SIRT rechazó la operación o algo falló. | Leer el Detalle (sección 9). |
| **Omitida** | «Ejecutar» = No. | Nada. |

### 7.2 Dónde quedan los resultados

- **En el mismo Excel:** Estado, Código SIRT, ID de registro, Fecha y Detalle de cada fila.
- **Copia de cada corrida:** `<carpeta del proyecto>\resultados\<TAREA>_EXCEL\`, con fecha y hora en el nombre (por
  ejemplo `REGISTRO_EXCEL\registro_excel_20261022_1451.xlsx`).
- **Capturas de pantalla de los errores:** `<carpeta del proyecto>\resultados\<SERVICIO>\capturas_EXCEL\`.
- **Historial completo** (archivos .csv que se abren con Excel): `<carpeta del proyecto>\resultados\`.
- **Códigos SIRT de todas las tarifas:** `<carpeta del proyecto>\resultados\codigos_sirt.csv`.

Si el Excel quedó abierto al terminar, la ventana negra lo avisa: se cierra el Excel y se escribe `si`. Si no se
puede, se escribe `copia` y los resultados quedan en la copia de `resultados\`.

---

## 8. Opciones de los comandos (para quien las necesite)

El comando de la sección 1 basta para el trabajo diario. Estas opciones se agregan al final del comando, separadas
por espacios:

| Opción | Para qué |
|---|---|
| `--solo-listar` | Ver qué haría y los errores de datos. **No toca SIRT ni el Excel.** |
| `--limite N` | Procesar solo las primeras N filas pendientes (p. ej. `--limite 1`). |
| `--nombres A B C` | Procesar solo esas tarifas (por su nombre). |
| `--si` | No preguntar «¿continuar?» antes de empezar. Sin esta opción, hay que escribir `si`. |
| `--excel RUTA` | Usar otro Excel con el mismo formato (p. ej. una copia en otra carpeta). |
| `--definitivo-en-bloque` | (registros) Disponibilidad y registro definitivo de varias tarifas a la vez. Es mucho más rápido. |
| `--sin-definitivo` | (registros) Solo registrar; el registro definitivo queda para después. |
| `--hasta-resumen` | (modificaciones) Llegar al resumen de SIRT sin guardar. |
| `--repetir` | (registros y modificaciones) Volver a hacer filas que ya figuran como hechas. Con cuidado. |
| `--diagnostico` | Guardar capturas de cada paso, para investigar un error. |

Ejemplo: registrar solo dos tarifas, sin preguntar:

```
python registro_excel.py --nombres TE-TF-EJ-011 TE-TF-EJ-012 --si --definitivo-en-bloque
```

---

## 9. Preguntas frecuentes y errores comunes

### Al empezar

**«El Excel está abierto: ciérralo y vuelve a correr el comando.»**
Cerrar el Excel y volver a escribir el comando (flecha ↑ y Enter).

**«No se pudo conectar con Edge (puerto 9222).»**
Edge no se abrió con `iniciar.bat`. Cerrar Edge del todo y volver a abrir `iniciar.bat`.

**«No se encontró una pestaña abierta de SIRT.»**
En el Edge de la herramienta no está abierto SIRT. Abrirlo e iniciar sesión.

**«No se reconoce python» o «no existe el archivo .py».**
La ventana negra no es la que abrió `iniciar.bat`. Cerrarla y usar la que abre `iniciar.bat`.

**`iniciar.bat` dice que Edge ya está abierto, pero no pasa nada.**
Había un Edge abierto de forma normal. Cerrarlo del todo (también desde el Administrador de tareas) y volver a abrir
`iniciar.bat`.

### Mientras trabaja

**Parece que se quedó quieto.**
Casi siempre es porque Edge está **minimizado**: restaurarlo. Si no, puede que SIRT esté lento (sección 10).

**¿Puedo usar la computadora mientras tanto?**
Sí, con otras ventanas encima de Edge, sin minimizarlo y sin hacer clic dentro. También se puede bloquear la
computadora (Windows + L).

**¿Puedo correr dos comandos a la vez?**
No. Se corre uno, y al terminar el siguiente.

**Fallan varias filas seguidas o aparece «Error al obtener las acciones de la bandeja».**
La **sesión de SIRT venció**. Iniciar sesión de nuevo y correr el mismo comando: continúa donde quedó.

### En los resultados

**Una fila dice «No aplica».**
La tarifa todavía no está en comercialización y vigente. Correr el comando otro día.

**Una fila dice «Error de datos».**
El Detalle indica qué corregir. Algunos ejemplos:
- «Alcance de la tarifa: … no es una opción válida»
- «la fecha … debe ser posterior a hoy»
- «un paquete necesita al menos 2 tarifas»

Se corrige la fila y se vuelve a correr.

**«SIRT no guardó el paso 1: Tarifa Instalación: Este apartado es obligatorio.»**
El paquete tiene servicios fijos: hay que poner una tarifa de instalación.

**«El valor del servicio … no debe de ser mayor al valor del plan de la tarifa … que está siendo empaquetada.»**
En el paquete, cada valor debe ser ≤ el de su tarifa y la suma menor que la suma de las tarifas. Lo más simple es
dejar los valores vacíos.

**«El alcance pedido (…) no coincide con el de las tarifas (…)».**
El paquete y sus tarifas deben tener el mismo alcance.

**«no aparece en Tarifas registradas».**
Revisar que el nombre sea exactamente el de SIRT. Si la tarifa se renombró a mano en SIRT, escribir su nombre actual.

**«El resumen no muestra N de los cambios».**
Medida de seguridad: SIRT no mostró un cambio en el resumen y por eso no se guardó nada. Avisar a soporte, de
preferencia corriendo esa fila con `--diagnostico` para que queden capturas.

**El Detalle de una promoción dice «solo muestra los servicios en común» o «no tienen apartados en común».**
Es normal cuando se mezclan servicios distintos: SIRT solo deja promocionar lo compartido.

**Una renovación dice «renovación ya registrada».**
El intento anterior sí se guardó justo antes de un corte. Está bien.

**«Ya registrada» o «Ya aplicada».**
Ya se había hecho antes; no se repitió. El código y la fecha son los originales.

**El «Código SIRT» está vacío.**
Se completa al hacer el registro definitivo. Si la fila quedó «Registrada», falta ese paso.

**Al terminar dice que el Excel está ABIERTO.**
Cerrarlo y escribir `si`. Si no se puede, escribir `copia`: los resultados quedan en la copia de `resultados\`.

---

## 10. Posibles problemas (para soporte y desarrolladores)

| Problema | Cómo se nota | Qué hacer |
|---|---|---|
| **SIRT lento (tiempos de carga altos)** | Filas con «Timeout … exceeded» en el Detalle, o pasos que tardan mucho más de lo normal (una tarifa suele tardar entre 10 y 30 s). | Volver a correr el comando cuando SIRT responda mejor: lo hecho no se repite. Si es constante, avisar a desarrollo para ajustar los tiempos de espera. |
| **Sesión de SIRT vencida** | Varias filas seguidas en error al buscar o abrir pantallas. | Iniciar sesión y correr el mismo comando. |
| **Edge minimizado o cerrado** | El proceso se queda quieto, o aparece «No se pudo conectar con Edge». | Restaurar Edge, o abrirlo con `iniciar.bat` e iniciar sesión, y correr el mismo comando. |
| **Corte a mitad de una operación** (luz, red, cierre de la ventana) | Una fila en «Error» cuya operación sí quedó hecha en SIRT. | Correr el mismo comando: la herramienta lo detecta en la mayoría de los casos. Si no, revisar esa tarifa en SIRT y avisar a desarrollo. |
| **SIRT cambió una pantalla o una regla** | Todas las filas de un mismo tipo fallan siempre en el mismo paso (el Detalle empieza con `[nombre del paso]`). | Avisar a desarrollo con el Detalle y la captura de `resultados\…\capturas_EXCEL\`. |
| **Red de la empresa o proxy** | No carga SIRT en el Edge de la herramienta. | Revisar la conexión con TI; la herramienta usa la misma red que Edge. |

---

## 11. Lo que la herramienta no hace (versión 1.0)

- Fechas de inicio anteriores a hoy.
- Pasar un beneficio de limitado a ilimitado.
- Trabajar con dos procesos a la vez.
- Iniciar sesión en SIRT: siempre lo hace la persona. La herramienta no guarda contraseñas.

El registro definitivo no se puede deshacer en SIRT. Por eso, ante la duda, conviene usar primero `--solo-listar` y
`--limite 1`.
