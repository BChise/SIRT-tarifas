# Tutorial: app web de la herramienta SIRT

La app reúne en una página del navegador todo lo que hoy se hace con comandos:
registrar tarifas, el registro definitivo, empaquetar, promocionar y las acciones
(buscar, renombre, cese, fe de erratas, rectificar y cambio). Cada botón lanza el
mismo script de siempre, muestra su salida en vivo y deja los resultados en los CSV
de `resultados/`.

---

## 1. Arrancar

Desde la carpeta del proyecto:

```bash
bash iniciar_app.sh
```

El script:
1. activa el entorno `register_web_env`;
2. instala `fastapi`, `uvicorn` y `openpyxl` si faltan (solo la primera vez);
3. abre Edge en el puerto 9222 con el perfil de SIRT si no estaba abierto;
4. arranca la app y abre **http://127.0.0.1:8000** en el navegador.

Si Edge se abrió recién, **inicia sesión en SIRT** en esa ventana antes de lanzar tareas.
Para detener la app: `Ctrl+C` en la terminal.

Opciones: `bash iniciar_app.sh --sin-edge` (no toca Edge) · `PUERTO=8080 bash iniciar_app.sh`.

> Arranque manual equivalente: `source ../../register_web_env/bin/activate`,
> `pip install fastapi uvicorn openpyxl` y `python -m app.servidor`.

---

## 1b. Versión simple (la que se abre por defecto)

La página principal (**http://127.0.0.1:8000**) es la versión simple, pensada para la
presentación. Trabaja con las tarifas de **`datos/pruebas.xlsx`** (hoja «Tarifas»):

- **Resumen:** cuántas tarifas hay en el Excel, pendientes, registradas y en definitivo.
- **Registrar tarifas:** abre una ventana con las filas del Excel; marca las que quieras
  (por defecto, las pendientes) → se registran en SIRT con su nombre exacto y alcance.
- **Registro definitivo:** igual, con las que ya están registradas.
- **Tabla:** estado de cada fila (Pendiente → Registrada → Definitivo, o Error) y la
  acción prevista. Se actualiza sola mientras corre la tarea.
- **Progreso:** una línea con lo que está haciendo; «Ver detalle» muestra la consola completa.
- Empaquetados, Promociones y Acciones aparecen como «próximamente».

El Excel trae 7 tarifas de ejemplo (lote DEMO): 1 testigo sin acciones y 1 por cada acción
(renombre, cese, fe de erratas, rectificación, cambio con reducción y con incremento). Para
otras pruebas, edita las celdas amarillas o agrega filas (la hoja «Leyenda» explica cada
columna). Las acciones exigen que la tarifa esté vigente, lo que ocurre **el día siguiente**
a su registro definitivo.

### Licitación y Contratos

- Botón **«Registrar licitaciones y contratos»** (debajo de los otros cuatro): elige tarifas de
  **`datos/licitaciones.xlsx`** (por defecto ninguna marcada) → `licitacion.py` las registra
  (servicios, fechas, duración, datos del contrato y el detalle de cada servicio) y luego
  `registro_definitivo.py` hace su registro definitivo (sin disponibilidad).
- Icono naranja **«Excel de licitaciones»** arriba, junto al otro Excel: lo descarga.
- El Excel trae 27 tarifas: 6 de 1 servicio, 20 de 3 y 1 de 6. Para agregar grupos de 2, 4 y 5:
  `python datos/generar_licitaciones.py --grupos 1 2 3 4 5 6` (63 tarifas).
- La duración del contrato se calcula sola: fin de vigencia = hoy + duración − 1 (SIRT exige
  que cuadre con las fechas). Si SIRT la rechaza, el script prueba un día menos y uno más.
- En terminal: `python licitacion.py --solo-listar` · `python licitacion.py --nombres TL-LYC-DEMO-001 --diagnostico`.

La versión completa sigue disponible en **http://127.0.0.1:8000/avanzado**.

---

## 2. La pantalla (versión avanzada)

- **Arriba a la derecha:** «Edge conectado» en verde (si sale en rojo, Edge no está en el
  puerto 9222) y la tarea en curso.
- **Pestañas:** una por tipo de trabajo.
- **Consola (derecha):** la salida en vivo de la tarea, igual que en la terminal.
  «Cancelar» detiene la tarea (lo ya guardado en SIRT no se deshace).
- **Historial:** las tareas de esta sesión; clic en una para volver a ver su salida.

Solo corre **una tarea a la vez** (hay un solo Edge). Las tareas que guardan en SIRT
piden confirmación antes de empezar; no hace falta escribir «si» en la terminal.

---

## 3. Prueba guiada (sin riesgo primero)

### Paso 1 — Buscar (no cambia nada)
1. Pestaña **Acciones** → Acción: **Buscar**.
2. En **Tarifas**, escribe `OS` en el buscador y marca `TE-OS-PROMO-001`.
3. **Ejecutar acción**. En la consola debe salir su código SIRT, versión y estados.

### Paso 2 — Fe de erratas hasta el resumen (no guarda)
1. Acción: **Fe de erratas**, deja marcado **«Solo hasta el resumen»**.
2. Elige una tarifa en comercialización (p. ej. `TE-OS-PROMO-002`).
3. Campos: **Elegir campos** → aparece la lista de campos del servicio OS; marca solo
   «Condiciones generales».
4. **Ejecutar acción**. En la consola verás `Campos elegidos (1): Condiciones generales`
   y `1 texto(s)` en el paso 2; termina con «detenido en el resumen (sin guardar)».

### Paso 3 — Rectificar con campos elegidos (hasta el resumen)
1. Acción: **Rectificar**, «Solo hasta el resumen» marcado.
2. Tarifa móvil (escribe `SMP`, marca `TE-SMP-PROMO-001`).
3. **Elegir campos** → en el desplegable de SMP marca «Datos para tethering» y «Valor del servicio».
4. Ejecuta. Deben cambiar solo esos dos (`Datos para tethering GB: 5 → 6`, `Valor del servicio: …`).

### Paso 4 — Ver resultados
Pestaña **Resultados** → archivo `ACCIONES/acciones.csv`. Las filas más recientes salen
primero; EXITOSO en verde, ERROR en rojo, EN_RESUMEN / NO_APLICA en ámbar. Usa
«Filtrar» para buscar una tarifa.

Cuando esto funcione, repite cualquiera de los pasos **desmarcando «Solo hasta el
resumen»** para guardar de verdad (la app pide confirmación).

---

## 4. Cada pestaña

| Pestaña | Qué hace | Script |
|---|---|---|
| **Registrar establecidas** | Marca servicios, lote y casos por servicio → registra uno tras otro. | `ejecutar.py` (uno por servicio) |
| **Registro definitivo** | Pasa a definitivo las temporales de un lote (con disponibilidad cuando aplica). Filtros opcionales por servicio o nombre. | `registro_definitivo.py` |
| **Empaquetar** | Elige 2 o más tarifas con registro definitivo → crea un empaquetado. | `empaquetar.py` |
| **Promocionar** | Tarifas (mono o empaquetadas) → promociones. «Una promoción por tarifa» o todas juntas; atributos probados, todos o una lista. | `promocionar.py` |
| **Acciones** | Buscar, renombrar, cese, fe de erratas, rectificar y cambio. | `acciones.py` |
| **Resultados** | Muestra cualquier CSV de `resultados/`. | — |

### El selector de tarifas
- La lista sale de los CSV de resultados (tarifas EXITOSAS; los renombres ya aplicados).
- Busca por nombre o por código (`SMP`, `EMP`…). Lo elegido **se conserva** al cambiar el filtro.
- «solo con registro definitivo» oculta las temporales.
- **Otras tarifas:** para nombres que no aparezcan (p. ej. registradas a mano), una por línea.

### Campos a modificar (fe de erratas, rectificar, cambio)
- **Por defecto:** textos + valor del servicio + tarifa de instalación (lo de siempre).
- **Todos los campos:** además todos los números del paso 2 (equivale a `--todo`).
- **Elegir campos:** un desplegable por servicio de las tarifas elegidas, con cada campo
  del paso 2 (sacado del JSON de SIRT de ese servicio) y del paso 1. La casilla del
  título marca todos los del servicio.
  - Los campos en gris **todavía no se pueden cambiar** y dicen por qué: tablas (falta
    grabar cómo se edita una fila), listas de una sola opción, archivos Excel.
  - En **Cambio** no aparecen los campos que SIRT no muestra en esa acción (según
    `verificacion_cambio.xlsx`).
  - Fe de erratas solo permite textos.
  - Empaquetados (`EMP`): campos comunes, aplicados en cada pestaña de servicio.

---

## 5. Si algo falla

- **«Edge NO conectado»:** abre Edge con
  `microsoft-edge --remote-debugging-port=9222 --user-data-dir="$HOME/perfiles/sirt" &`.
- **«Ya hay una tarea en curso»:** espera a que termine o cancélala.
- **Una tarea con errores:** la consola muestra el error; con «Diagnóstico» marcado, las
  capturas quedan en `resultados/<tipo>/diagnostico…` como siempre.
- **El puerto 8000 está ocupado:** `PUERTO=8080 bash iniciar_app.sh`.

Los comandos de terminal siguen funcionando igual; la app no los reemplaza.

---

## 6. Archivos nuevos

- `app/servidor.py` — servidor (FastAPI): lanza los scripts y entrega salida y resultados.
- `app/catalogo.py` — servicios, campos por acción y tarifas conocidas.
- `app/static/` — la página (HTML, CSS y JavaScript, sin dependencias externas).
- `iniciar_app.sh` — arranque.
- `acciones.py` — nueva opción `--campos "A;B;C"` (la usa «Elegir campos»; también sirve en terminal).
