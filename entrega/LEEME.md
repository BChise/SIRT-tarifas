# Carpeta de entrega

| Carpeta / archivo | Qué es |
|---|---|
| `plantillas\` | Los 6 Excel de trabajo **vacíos** (mismo formato, listas e instrucciones). Copiar a `automatizacion\` con el mismo nombre, o usar con `--excel RUTA`. |
| `ejemplos\` | Los 6 Excel de la prueba funcional (serie `EJ`): cada fila muestra una opción distinta, con las columnas «Qué se prueba» y «Resultado esperado». Sirven de modelo para llenar los de trabajo. **Son ejemplos:** antes de usarlos en SIRT, cambiar los nombres de las tarifas. |
| `matriz_de_cobertura.xlsx` | Resultado de las dos pruebas funcionales completas (P4 y P5): resumen por módulo y detalle fila por fila. |

| Plantilla | Comando |
|---|---|
| `registro_tarifas_establecidas.xlsx` | `python registro_excel.py` |
| `empaquetados.xlsx` | `python empaquetados_excel.py` |
| `registro_licitaciones.xlsx` | `python licitaciones_excel.py` |
| `promociones.xlsx` | `python promociones_excel.py` |
| `acciones_basicas.xlsx` | `python acciones_excel.py` |
| `modificaciones.xlsx` | `python modificaciones_excel.py` |

Correspondencia de los ejemplos: `registro_funcional` → registro de establecidas, `empaquetados_funcional` →
empaquetados, `licitaciones_funcional` → licitaciones, `promociones_funcional` → promociones,
`acciones_funcional` → acciones básicas, `modificaciones_funcional` → modificaciones.
