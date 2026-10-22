# 3. Referencia rápida (para imprimir)

## Arrancar
1. Cerrar Edge del todo → doble clic en **`iniciar.bat`**.
2. Iniciar sesión en SIRT → dejar **«Tarifas registradas»**.
3. **Cerrar el Excel** que se va a usar.

## Reglas de oro
- **No minimizar Edge** (sí se puede tapar o bloquear Windows con Win + L). No tocar Edge mientras corre.
- **No editar el Excel** mientras su comando corre.
- Siempre: **`--solo-listar` → `--limite 1 --si` → todo**.
- Acciones y modificaciones: **al día siguiente** del registro definitivo.
- Si algo se corta: iniciar sesión otra vez y correr **el mismo comando** con **el mismo Excel**.

## Comandos (desde `D:\proyecto_sirt\sirt_tarifas>`)

| Tarea | Excel en `automatizacion\` | Comando completo |
|---|---|---|
| Establecidas | `registro_tarifas_establecidas.xlsx` | `python registro_excel.py --si --definitivo-en-bloque` |
| Empaquetados | `empaquetados.xlsx` | `python empaquetados_excel.py --si --definitivo-en-bloque` |
| Licitaciones | `registro_licitaciones.xlsx` | `python licitaciones_excel.py --si --definitivo-en-bloque` |
| Promociones | `promociones.xlsx` | `python promociones_excel.py --si --definitivo-en-bloque` |
| Acciones | `acciones_basicas.xlsx` | `python acciones_excel.py --si` |
| Modificaciones | `modificaciones.xlsx` | `python modificaciones_excel.py --si` |

## Opciones
| | |
|---|---|
| `--solo-listar` | ver qué haría (no toca nada) |
| `--limite 1` | solo la primera fila pendiente |
| `--nombres A B` | solo esas tarifas |
| `--si` | sin preguntar |
| `--excel RUTA` | otro Excel con el mismo formato |
| `--hasta-resumen` | (modificaciones) llega al resumen sin guardar |
| `--diagnostico` | capturas de cada paso |
| `--sin-definitivo` | (registros) sin registro definitivo |

## Estados
**Registro definitivo / Aplicada** = hecho · **Ya registrada / Ya aplicada** = ya estaba, no se repitió ·
**Registrada** = falta el registro definitivo · **No aplica** = aún no vigente, otro día · **Error de datos** = corregir
el Excel · **Error** = leer el Detalle · **Omitida** = Ejecutar = No.

## Valores por defecto (celda vacía)
| | |
|---|---|
| Ejecutar | Sí |
| Inicio de vigencia / comercialización | hoy |
| Disponibilidad (registro) | Lima |
| Valor en el paquete | valor base − 1 |
| Delta | +1 rectificación · −1 cambio |
| Fecha de ejecución (cambio) | hoy + 3 días |
| Fecha de cese | mañana |
| Sufijo de renombre | `_REN` |
| Renovación | 19 días · duraciones 25 (31 si ya eran 25) |
