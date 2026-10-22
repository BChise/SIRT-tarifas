# 9. Guiones de los videos

Videos cortos (3–5 min), uno por tema. Grabar con la pantalla completa (cmd a un lado y Edge al otro). Antes de grabar:
ocultar correos y datos personales, y **no mostrar las herramientas de desarrollador (F12)** ni el inicio de sesión
con la contraseña. Usar Excel de prueba (una serie nueva) para no tocar tarifas reales.

---

## Video 1 — Instalación y primer arranque (5 min)

| Mostrar | Decir |
|---|---|
| Carpeta `D:\proyecto_sirt\sirt_tarifas` | «Este es el proyecto. Se instala una sola vez.» |
| cmd: `python -m venv .venv`, `.venv\Scripts\activate.bat`, `pip install -r requirements.txt` | «Creamos el entorno de Python e instalamos dos librerías: Playwright, que maneja Edge, y openpyxl, que lee los Excel. Usamos cmd, no PowerShell.» |
| Cerrar Edge; doble clic en `iniciar.bat` | «Antes, Edge cerrado del todo. `iniciar.bat` activa el entorno y abre Edge en modo controlable.» |
| Iniciar sesión en SIRT, ir a «Tarifas registradas» | «El inicio de sesión siempre lo hacemos nosotros. Dejamos la pantalla en Tarifas registradas.» |
| `python registro_excel.py --solo-listar` | «Comprobamos que todo funciona: `--solo-listar` muestra qué haría, sin tocar SIRT.» |
| Minimizar Edge → señalar que no | «Regla de oro: Edge nunca minimizado. Taparlo sí; minimizarlo, no.» |

## Video 2 — Cómo son los Excel (4 min)

| Mostrar | Decir |
|---|---|
| Carpeta `automatizacion\`, los 6 Excel | «Cada tarea tiene su Excel y su comando.» |
| Fila 5 y colores de los títulos | «Fila 5, títulos; los datos desde la 6. Verde oscuro: identificación; azul oscuro: obligatorio; azul claro: opcional; gris: lo llena el script.» |
| Lista desplegable; hoja de instrucciones | «Las listas muestran los valores válidos; cada Excel trae su hoja de instrucciones.» |
| Columna Ejecutar | «Sí o vacío: se procesa. No: se salta.» |
| Columnas Estado, Código SIRT, ID, Detalle | «Al terminar, aquí queda el resultado de cada fila y el código oficial de SIRT.» |
| Tabla de estados (manual 2.3) | «Registro definitivo y Aplicada: hecho. Ya registrada: no se repite. No aplica: otro día. Error de datos: corregir el Excel.» |

## Video 3 — Registro de tarifas establecidas (5 min)

| Mostrar | Decir |
|---|---|
| Hoja SMP del Excel, una fila; hojas de Tablas | «Una hoja por servicio, una fila por tarifa. Las tablas se enlazan por el nombre.» |
| Columnas Registro definitivo, Disponibilidad, Alcance, fechas | «Vacío en las fechas es hoy; disponibilidad vacía es Lima.» |
| Cerrar el Excel; `--solo-listar` | «Siempre primero el listado: aquí aparecen los errores de datos.» |
| `--limite 1 --si`; Edge trabajando | «Una sola para comprobar.» |
| `--si --definitivo-en-bloque`; resumen y tiempo total | «Después todo. En bloque, el registro definitivo pasa de 20 segundos a 1 o 2 por tarifa.» |
| Excel con resultados | «Cada fila con su estado y su código SIRT.» |

## Video 4 — Empaquetados, licitaciones y promociones (5 min)

| Mostrar | Decir |
|---|---|
| empaquetados.xlsx | «De 2 a 6 tarifas con registro definitivo y el mismo alcance. Valor vacío es el de la tarifa menos 1. Con fijos, la instalación es obligatoria.» |
| registro_licitaciones.xlsx | «Por servicio: incluir y detalle. Moneda, ámbito, vigencia y duración.» |
| promociones.xlsx + hoja Catálogo | «De 1 a 15 tarifas base. Atributos: Todos, Lista o Detalle, con los nombres del catálogo.» |
| Comandos con `--definitivo-en-bloque` | «El mismo patrón: listar, 1, todo.» |

## Video 5 — Acciones básicas (4 min)

| Mostrar | Decir |
|---|---|
| acciones_basicas.xlsx | «Fe de erratas, renovación (solo promociones), cese (no promociones) y renombre.» |
| Hoja «Qué hace cada acción» | «Vacíos: cese mañana, sufijo _REN, renovación 19 días.» |
| Mensaje «No aplica» | «Se corre al día siguiente del registro definitivo: antes SIRT no lo permite.» |
| Fila con tarifa ya renombrada | «Si ya se renombró, la herramienta usa su nombre actual sola.» |
| `python acciones_excel.py --si` | |

## Video 6 — Rectificación y cambio de tarifa (5 min)

| Mostrar | Decir |
|---|---|
| modificaciones.xlsx | «Rectificación corrige la tarifa vigente; Cambio programa una versión nueva con fecha de ejecución.» |
| Delta y Sentido | «Delta: cuánto cambia. Sentido beneficioso: bajan los precios y suben los beneficios.» |
| Campos a modificar; hoja Campos con rutas `›` | «Por defecto, Todos, Solo la lista o Detalle, donde se elige cada campo, incluso una velocidad de una fila de tabla.» |
| `--hasta-resumen --limite 1` | «Prueba segura: llega al resumen y no guarda.» |
| `--si` | «Antes de guardar, la herramienta comprueba que cada cambio aparezca en el resumen.» |

## Video 7 — Qué hacer si algo falla (4 min)

| Mostrar | Decir |
|---|---|
| Fila con Error y su Detalle | «El Detalle dice en qué paso falló y por qué; en `resultados` hay una captura.» |
| Sesión vencida → iniciar sesión → mismo comando | «Si se corta, iniciamos sesión y corremos el mismo comando con el mismo Excel: lo hecho no se repite.» |
| Mensaje «El Excel está abierto» | «Cerrarlo y escribir si.» |
| `docs\04_preguntas_frecuentes.md` | «Aquí están los casos más comunes con su solución.» |

## Video 8 (opcional) — Prueba funcional (3 min)

| Mostrar | Decir |
|---|---|
| `prueba_dia1.bat P9 Prueba9` | «Prueba todas las opciones con una serie nueva de tarifas. Se usa tras cualquier cambio en SIRT o en el código.» |
| `prueba_dia2.bat Prueba9` al día siguiente | |
| `verificar_funcional.py --carpeta Prueba9` | «OK, falla y pendiente por cada fila: la matriz de cobertura.» |
