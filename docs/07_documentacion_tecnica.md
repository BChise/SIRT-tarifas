# 7. Documentación técnica

Para quien mantenga o amplíe el proyecto. Python 3.13 + Playwright 1.64 (sync API) + openpyxl 3.1.

## 7.1 Cómo funciona

```
Excel (automatizacion\*.xlsx)
   │  lo lee y valida
   ▼
<módulo>_excel.py  ──── por cada fila ────►  script de SIRT (subproceso)  ──►  Edge (CDP :9222) ──► SIRT
   ▲                                              │
   │  Estado / Código / Detalle                   ▼
   └──────────────────────────────  resultados\*.csv (historial: fuente de verdad)
```

- Playwright **no abre** un navegador: se conecta por CDP (`connect_over_cdp`) al Edge que abrió `iniciar.bat`
  (`comun.CDP_URL`, puerto en `SIRT_PUERTO`, por defecto 9222) y usa la pestaña cuyo URL contiene «sirt»
  (`comun.obtener_pagina_sirt`; con `SIRT_PESTANA=N` usa la N-ésima).
- Cada `*_excel.py` valida las filas (sin tocar SIRT), arma los argumentos y llama al **script de SIRT** del módulo
  como subproceso (`excel_comun.correr`); `registro_excel.py` registra en su propio proceso. El script de SIRT anota cada intento en un CSV de `resultados\`; el
  `*_excel.py` lee ese CSV y escribe el Estado en el Excel. Así la reanudación se basa en el CSV (lo EXITOSO no se
  repite) y el Excel se guarda después de cada fila.
- Lote: todo lo que viene de los Excel usa el lote **`EXCEL`** (`excel_comun.LOTE_EXCEL`).

## 7.2 Archivos

### Comandos de Excel (los que usan los analistas)
| Archivo | Lee | Llama a |
|---|---|---|
| `registro_excel.py` | `registro_tarifas_establecidas.xlsx` | Registra en el mismo proceso con el código de `servicios\*.py` y `comun.py` (como `ejecutar.py`); `registro_definitivo.py` como subproceso. |
| `empaquetados_excel.py` | `empaquetados.xlsx` | `empaquetar.py`, `registro_definitivo.py` |
| `licitaciones_excel.py` | `registro_licitaciones.xlsx` | `licitacion.py`, `registro_definitivo.py` |
| `promociones_excel.py` | `promociones.xlsx` | `promocionar.py`, `registro_definitivo.py` |
| `acciones_excel.py` | `acciones_basicas.xlsx` | `acciones.py` (fe_erratas, renovar, cese, renombre) |
| `modificaciones_excel.py` | `modificaciones.xlsx` | `acciones.py` (rectificar, cambio) |
| `excel_comun.py` | — | Clase `Hoja` (fila 5 = títulos, columnas de resultado), validaciones, `correr`, `guardar_final`, `precargar_definitivos`, `restaurar_validaciones`. |

### Scripts de SIRT
| Archivo | Qué hace |
|---|---|
| `comun.py` | Conexión, listas, fechas (calendario), paso 1, bloques reutilizables (tablas, condicionales), confirmación, CSV de valores y de códigos SIRT, `campos_con_error`, `cerrar_calendarios`. |
| `servicios\*.py` | Un archivo por servicio (15): `SERVICIO` (texto exacto en SIRT), `CODIGO`, `DIMENSIONES`, `VALORES`, `completar_paso_2()`. Plantilla: `_plantilla.py`. |
| `ejecutar.py` | Ejecutor original por combinaciones (lotes de prueba de la fase 1). |
| `registro_definitivo.py` | Registro definitivo 1×1 o `--en-bloque` (disponibilidad «Selector geográfico» + varias filas marcadas; controla «Tarifas seleccionadas»; lo que falla pasa a 1×1). Lee el Código SIRT del JSON de temporales. |
| `empaquetar.py` | Empaquetado: valida reglas de precio/alcance antes de tocar SIRT, agrega tarifas, valores por pestaña. |
| `licitacion.py` | Licitación y Contratos (servicios, detalle por servicio, vigencia/duración). |
| `promocionar.py` | Promociones: catálogo de atributos, tipo preferido, servicios en común, `--config` JSON. |
| `acciones.py` | Todas las acciones sobre tarifas registradas: buscar (JSON del filtro), fe de erratas, rectificar, cambio (con `--sentido`, `--detalle`, `--todo`, `--campos`, tablas, ilimitado, canales), renovar, cese, renombre. Comprueba el resumen antes de guardar. |
| `completar_codigos_sirt.py` | Relleno único de Códigos SIRT. |
| `grabar_sirt.py`, `inspeccionar_json.py` | Herramientas de mantenimiento (grabar flujos, leer el JSON de un formulario). |

### Pruebas
| Archivo | Qué hace |
|---|---|
| `automatizacion\pruebas_funcionales\generar_pruebas_funcionales.py` | Genera los 6 Excel de una serie (`--serie`, `--carpeta`). |
| `prueba_funcional.py` + `prueba_dia1.bat` / `prueba_dia2.bat` | Corre un día completo módulo por módulo; se detiene si algo no da lo esperado. |
| `verificar_funcional.py` | Compara «Estado» vs «Resultado esperado»; escribe la matriz de cobertura. |
| `automatizacion\generar_*.py` | Generadores de plantillas y de los Excel de escenarios. |

### Otros
- `grabaciones\`: material de cada flujo (grabación del Inspector, JSON, capturas, `notas.txt`). Es la referencia
  para mantenimiento.
- `app\`, `TUTORIAL_APP.md`, `iniciar_app.sh`: app web local de la fase anterior (congelada; no forma parte de la
  entrega).
- `CONTEXTO.md`: bitácora de decisiones y hallazgos.

## 7.3 Archivos de resultados (`resultados\`, no se suben a git)

| Archivo | Columnas principales |
|---|---|
| `<SERVICIO>\resultados_EXCEL.csv` | correlativo, nombre, dimensiones, id_registro, resultado (EXITOSO/ERROR), detalle, captura, fecha_hora |
| `registro_definitivo_EXCEL.csv` | nombre, codigo_servicio, id_registro, disponibilidad, resultado, detalle, fecha_hora |
| `ACCIONES\acciones.csv` | fecha_hora, accion, nombre, codigo_sirt, version, nombre_nuevo, comercializacion, vigencia, resultado (EXITOSO/ERROR/NO_APLICA/EN_RESUMEN), detalle |
| `codigos_sirt.csv` | nombre, codigo_sirt, id_tarifa, fecha_hora |
| `valores_tarifas.csv` | nombre, valor_servicio (para las reglas de empaquetados) |
| `<MÓDULO>_EXCEL\*.xlsx` | Copia del Excel de cada corrida |
| `verificacion_funcional_*.csv` | Matriz de cobertura |

## 7.4 Decisiones de diseño

- **Selectores por rol y texto visible** (`get_by_role(..., name=...)`), nunca por ids internos (`#mat-input-102`
  cambia en cada carga).
- **Datos desde el JSON de SIRT** cuando la tabla no basta: buscar tarifa (`/api/tarifa/filtrar…`), renovación
  (atributos promocionados), códigos SIRT.
- **Verificación antes de guardar**: en modificaciones, todo valor escrito debe aparecer en el resumen; si no, no se
  guarda. Tras escribir un número se relee lo que quedó en el campo.
- **Validación previa** de reglas de SIRT conocidas (alcance, precios de paquetes, fechas, acciones por tipo) para no
  dejar borradores.
- **Reanudación por CSV**: idempotencia por nombre + acción.
- **Seguridad en bloque**: solo nombres exactos del Excel en «Ingreso confirmado»; marcadas = «Tarifas seleccionadas»
  o se cancela; bloques de 10.

## 7.5 Variables de entorno

| Variable | Uso |
|---|---|
| `SIRT_PUERTO` | Puerto CDP de Edge (por defecto 9222). |
| `SIRT_PESTANA` | Qué pestaña de SIRT usar si hay varias (por defecto 1). |
| `PYTHONUTF8=1` | Lo pone `iniciar.bat`; tildes en la consola de Windows. |

## 7.6 Seguridad

- Ningún archivo guarda usuario ni contraseña: el inicio de sesión es manual.
- Las capturas de las herramientas de desarrollador (F12) muestran el token de sesión: no subirlas
  (`.gitignore` excluye `grabaciones\_registro_definitivo\Pasted image*.png`).
- El repositorio debe ser **privado**.
