# 2. Manual de usuario

La herramienta registra y modifica tarifas en **SIRT** (OSIPTEL) a partir de **Excel**. Cada tarea tiene su Excel en
la carpeta `automatizacion\` y un comando que lo lee. Se llena el Excel, se corre el comando y el mismo Excel queda con
el resultado de cada fila.

## 2.1 Antes de cada sesión de trabajo

1. Doble clic en `iniciar.bat` (abre Edge controlable y deja el cmd listo).
2. Iniciar sesión en SIRT en ese Edge y dejar **«Tarifas registradas»**.
3. **Cerrar el Excel** que se va a usar.

**Reglas de oro**

- **No minimizar Edge.** Minimizado, Windows lo congela y el proceso se detiene. Sí se puede tapar con otras ventanas
  o bloquear la sesión de Windows (Win + L).
- **No tocar Edge** (clics, teclas) mientras un comando trabaja.
- **No editar el Excel** mientras su comando corre: al terminar, el script lo sobrescribe con su versión. Para mirar el
  avance, ver la consola o abrir la copia de `resultados\`.
- **Primero `--solo-listar`, después `--limite 1`, recién después todo.**

## 2.2 Qué hace cada Excel y en qué orden

| Orden | Tarea | Excel (`automatizacion\`) | Comando |
|---|---|---|---|
| 1 | Registrar tarifas establecidas (monoproductos) | `registro_tarifas_establecidas.xlsx` | `python registro_excel.py` |
| 2 | Registrar empaquetados | `empaquetados.xlsx` | `python empaquetados_excel.py` |
| 2 | Registrar licitaciones y contratos (LyC) | `registro_licitaciones.xlsx` | `python licitaciones_excel.py` |
| 3 | Registrar promociones | `promociones.xlsx` | `python promociones_excel.py` |
| 4 | Fe de erratas, renovación, cese, renombre | `acciones_basicas.xlsx` | `python acciones_excel.py` |
| 4 | Rectificación y cambio de tarifa | `modificaciones.xlsx` | `python modificaciones_excel.py` |

- Un empaquetado o una promoción necesita que sus tarifas base ya tengan **registro definitivo**.
- Las acciones (orden 4) exigen que la tarifa esté **«En comercialización» y «Vigente»**, y SIRT las pasa a ese estado
  **el día siguiente** al registro definitivo. Por eso las acciones y modificaciones se corren **al día siguiente**
  (las licitaciones solo necesitan estar vigentes). Si se corren antes, la fila queda «No aplica» y no se toca SIRT.

## 2.3 Partes comunes de todos los Excel

- **Fila 5 = títulos.** Los datos empiezan en la fila 6. No cambiar los títulos de las columnas.
- **Colores de los títulos:** verde oscuro = identificación, azul oscuro = obligatorio, azul claro = opcional,
  gris = lo llena el script.
- **«Ejecutar»:** `Sí` o vacío = se procesa; `No` = se salta (queda «Omitida»).
- Las listas desplegables muestran los valores válidos. Cada Excel tiene una hoja de instrucciones («Instrucciones»,
  «Qué hace cada opción»…).
- **Columnas de resultado (las llena el script):**
  - **Estado** (ver tabla abajo).
  - **Código SIRT**: el código oficial de la tarifa (TEMTF…, TPRDS…, TLCDS…).
  - **ID de registro**: el número interno de SIRT (se agrega solo al final si no existe).
  - **Fecha** de registro o de ejecución.
  - **Detalle**: qué se hizo, avisos y, si falló, el motivo.

| Estado | Significado | Qué hacer |
|---|---|---|
| Registrada | Creada y confirmada en SIRT, sin registro definitivo (porque la fila lo pidió o porque falta). | Nada, o correr de nuevo para el registro definitivo. |
| Registro definitivo | Registrada con registro definitivo. | Nada. |
| Ya registrada / Ya aplicada | Ya estaba hecha en una corrida anterior; **no se repitió**. | Nada. |
| Aplicada | La acción o modificación se guardó en SIRT. | Nada. |
| Hasta resumen | Se llegó al resumen sin guardar (opción `--hasta-resumen`). | Correr sin esa opción para guardar. |
| No aplica | La tarifa no está en el estado que pide SIRT (p. ej. aún no vigente). **No se tocó SIRT.** | Correr otro día. |
| Error de datos | El Excel tiene un dato inválido. **No se tocó SIRT.** | Corregir la fila según el Detalle y correr de nuevo. |
| Error | SIRT rechazó o algo falló al trabajar. | Leer el Detalle (hay captura en `resultados\`); ver [preguntas frecuentes](04_preguntas_frecuentes.md). |
| Omitida | «Ejecutar» = No. | Nada. |

## 2.4 Opciones de los comandos

| Opción | Para qué |
|---|---|
| `--solo-listar` | Muestra qué haría y los errores de datos. **No toca SIRT ni el Excel.** |
| `--limite N` | Procesa solo las primeras N filas pendientes (para probar: `--limite 1`). |
| `--nombres A B C` | Solo esas tarifas (por su nombre en el Excel). |
| `--si` | No pregunta «¿continuar?» antes de empezar. |
| `--excel RUTA` | Usar otro Excel con el mismo formato (p. ej. una copia). |
| `--diagnostico` | Guarda capturas de cada paso (para investigar un error). |
| `--repetir` | (registros y modificaciones) Vuelve a hacer filas que ya figuran como hechas (con cuidado). |
| `--sin-definitivo` | (registros) Solo registra; el registro definitivo queda para después. |
| `--definitivo-en-bloque` | (registros) Disponibilidad y registro definitivo de varias tarifas a la vez (mucho más rápido). |
| `--tamano-bloque N` | (con la anterior) Tarifas por bloque; por defecto 10, máximo 50. |
| `--hasta-resumen` | (modificaciones) Llega al resumen de SIRT y **no guarda**: prueba segura. |

**Volver a correr es seguro.** Lo que ya salió bien no se repite; lo que falló o quedó pendiente se intenta de nuevo.
Si se cortó (sesión vencida, Edge cerrado, corte de luz), se vuelve a iniciar sesión y se corre **el mismo comando**
con **el mismo Excel** (no reemplazarlo: guarda el avance).

## 2.5 Forma de trabajo recomendada (para cualquier Excel)

```
python registro_excel.py --solo-listar
python registro_excel.py --limite 1 --si
python registro_excel.py --si --definitivo-en-bloque
```

1. `--solo-listar`: revisar el número de filas y corregir los «✖» (errores de datos).
2. `--limite 1`: confirmar en SIRT que la primera quedó bien.
3. Todo. Al final la consola muestra el **Resumen** y el **Tiempo total**, y guarda una copia del Excel en
   `resultados\<MÓDULO>_EXCEL\`.

Si el Excel quedó abierto al final, el script avisa y pide cerrarlo y escribir `si`; si no se puede, los resultados
quedan en la copia de `resultados\`.

---

## 2.6 Registro de tarifas establecidas — `registro_tarifas_establecidas.xlsx`

- **Una hoja por servicio**: TF, TFP, OTF, PTF, IF, PIF, TVP, TVPR, PTV, SMP, SMC, SMO, SMQ, OS, AC. Una fila = una
  tarifa.
- **Hojas de tablas**: «Tablas – Países», «Tablas – Velocidades», «Tablas – Bolsas». Cada fila se enlaza con su
  tarifa por el **nombre**. Sin filas: Países y Bolsas quedan vacías; Velocidades usa los valores del código.
- Columnas clave: Nombre, **Registro definitivo** (Sí/No), **Disponibilidad comercial** (solo servicios que la piden:
  fija, Internet fijo y TV; vacío = Lima; varias separadas por coma: `Lima, Arequipa`), **Alcance** (Residencial /
  Comercial / Residencial + Comercial), Inicio de vigencia e Inicio de comercialización (vacío = hoy; solo hoy o
  fechas futuras), datos de contacto, Valor del servicio y las listas y números del servicio.
- TV de paga: los archivos de canales están en `automatizacion\canales\`.

```
python registro_excel.py --solo-listar
python registro_excel.py --hojas SMP SMC --limite 1 --si     (solo esas hojas)
python registro_excel.py --si --definitivo-en-bloque
```

## 2.7 Empaquetados — `empaquetados.xlsx`

- Una fila = un paquete de **2 a 6 tarifas** (Tarifa 1…6). Las tarifas deben estar registradas, con registro
  definitivo y el **mismo alcance** que el paquete.
- **Valor N en el paquete**: vacío = valor de la tarifa base − 1. Reglas de SIRT, revisadas antes de tocarlo:
  cada servicio ≤ el valor de su tarifa base, el más caro ≤ la suma y la suma < la suma de las bases.
- **Tarifa instalación**: obligatoria si el paquete tiene servicios fijos; vacía solo en paquetes que no la
  tienen (p. ej. solo móviles).
- **Disponibilidad comercial**: vacía = la que tienen en común sus tarifas.

```
python empaquetados_excel.py --si --definitivo-en-bloque
```

Si 3 paquetes seguidos fallan en la misma etapa, el comando se detiene solo (suele ser la sesión vencida).

## 2.8 Licitaciones y contratos — `registro_licitaciones.xlsx`

- Una fila = una licitación. Por servicio: «<servicio>: incluir» (Sí/No) y «<servicio>: detalle del servicio».
  Al menos uno incluido.
- Inicio de vigencia (vacío = hoy), Duración del contrato (días), Identificador, Ámbito (Público/Privado), Valor de
  contrato, Moneda (Soles/Dólares/Euro), datos de la empresa. «Tarifa instalación» solo con «Otros servicios».

```
python licitaciones_excel.py --si --definitivo-en-bloque
```

## 2.9 Promociones — `promociones.xlsx`

- Una fila = una promoción sobre **1 a 15 tarifas base** (monoproductos y/o empaquetados con registro definitivo).
- **Atributos a promocionar**: `Los de siempre` (vacío), `Todos`, `Lista` (escribirlos en «Lista de atributos»,
  separados por `;`) o `Detalle en hoja Atributos` (una fila por atributo con su tipo, variación y duración). Los
  nombres exactos están en la hoja **«Catálogo de atributos»**.
- **Tipo de variación preferido**: Porcentual / Cantidad / Ilimitado (si el atributo no lo admite, se usa otro).
- Valor del plan e instalación: promocionar Sí/No, tipo y variación. Días de comercialización.
- Con tarifas de servicios distintos, SIRT solo deja promocionar los **servicios en común**; si no hay ninguno, solo
  el valor del plan / instalación. El Detalle lo indica.

```
python promociones_excel.py --si --definitivo-en-bloque
```

## 2.10 Acciones básicas — `acciones_basicas.xlsx` (al día siguiente)

| Acción | Para | Columnas | Vacío = |
|---|---|---|---|
| Fe de erratas | todas | — | agrega «\| Fe de erratas fecha hora» a los textos |
| Renovación | solo promociones (TP-…) | Días de comercialización, Duración (días) | 19 días; duraciones 25 (31 si ya eran 25) |
| Cese | monoproductos y empaquetados | Fecha de fin de comercialización (posterior a hoy) | mañana |
| Renombre | todas | Sufijo | `_REN` |

- Se ejecutan en el orden fe de erratas → renovación → cese → renombre.
- Si la tarifa **ya fue renombrada**, las demás acciones usan su nombre actual automáticamente (el Detalle lo indica).
  Para renombrarla otra vez, escribir en la fila el nombre actual (p. ej. `TARIFA_REN`).

```
python acciones_excel.py --solo-listar
python acciones_excel.py --limite 1 --si
python acciones_excel.py --si
```

## 2.11 Rectificación y cambio de tarifa — `modificaciones.xlsx` (al día siguiente)

- **Acción**: `Rectificación` (corrige la tarifa vigente, se guarda al momento) o `Cambio` (nueva versión desde una
  **fecha de ejecución**; vacía = hoy + 3 días). Solo puede haber **un cambio pendiente** por tarifa.
- **Delta**: cuánto suben o bajan los números (vacío: +1 en rectificación, −1 en cambio; admite decimales y negativos).
- **Cambio: sentido**: `Beneficioso` (bajan los precios, suben los beneficios) o `Perjudicial` (al revés). Usa el
  delta sin signo. Vacío = el delta igual para todo.
- **Campos a modificar**:
  - `Por defecto`: textos (con el sello), valor del servicio, instalación y listas como Tipo de contratación.
  - `Todos los campos`: además todos los números del paso 2 y las filas de las tablas.
  - `Solo la lista`: solo los campos escritos en «Lista de campos» (nombre tal como en SIRT, separados por `;`).
  - `Detalle en hoja Campos`: en la hoja **«Campos»**, una fila por campo con **Valor nuevo** o su propio **Delta**.
    Rutas con `›`: `campo`, `bloque › campo`, `tabla`, `tabla › fila`, `tabla › fila › campo`
    (p. ej. `Tecnologías y velocidades › 4G › Descarga máx. contratada`). No distingue mayúsculas ni tildes.
- Opcionales: Agregar una fila a cada tabla, Pasar «ilimitado» a limitado, Cambiar canales (TV; vacío = Sí),
  Texto del sello.
- **Empaquetados**: el valor de cada servicio nunca sube (regla de SIRT: cada servicio ≤ su tarifa base); el resto sí
  se modifica. **Licitaciones**: solo rectificación. **Promociones**: aún no disponibles.
- Antes de guardar, el script comprueba que **todos los cambios aparezcan en el resumen** de SIRT; si falta alguno, no
  guarda.

```
python modificaciones_excel.py --solo-listar
python modificaciones_excel.py --hasta-resumen --limite 1      (prueba sin guardar)
python modificaciones_excel.py --si
```

---

## 2.12 Herramientas adicionales

| Comando | Para qué |
|---|---|
| `python registro_definitivo.py --lote EXCEL --en-bloque` | Registro definitivo de lo que quedó solo «Registrada». Opciones: `--nombre X` (repetible), `--ubigeos "Lima, Arequipa"`, `--tamano-bloque N`. |
| `python completar_codigos_sirt.py` | Completa el «Código SIRT» de tarifas registradas antes del 21/10/2026 (una sola vez). |
| `prueba_dia1.bat SERIE CARPETA` / `prueba_dia2.bat CARPETA` | Prueba funcional completa (ver [pruebas y cobertura](08_pruebas_y_cobertura.md)). |
| `python verificar_funcional.py --carpeta CARPETA` | Matriz de cobertura de una prueba funcional. |
| `python grabar_sirt.py` | Abre el grabador de Playwright sobre el Edge del proyecto (para mantenimiento). |

## 2.13 Dónde quedan los resultados

- En el **mismo Excel** (Estado, Código SIRT, ID, Fecha, Detalle).
- `resultados\<MÓDULO>_EXCEL\`: copia del Excel de cada corrida (con fecha y hora).
- `resultados\<SERVICIO>\resultados_EXCEL.csv`, `resultados\registro_definitivo_EXCEL.csv`,
  `resultados\ACCIONES\acciones.csv`: historial completo de cada operación.
- `resultados\codigos_sirt.csv`: nombre → Código SIRT e ID.
- `resultados\<SERVICIO>\capturas_EXCEL\`: captura de pantalla de cada error.
