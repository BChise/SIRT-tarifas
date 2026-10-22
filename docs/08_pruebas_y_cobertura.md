# 8. Pruebas y cobertura

## 8.1 Qué se probó

**Prueba funcional**: un juego de Excel con el mismo formato que los de trabajo, donde cada fila prueba una opción
concreta y declara su **resultado esperado** (incluidos errores a propósito, que deben quedar en «Error de datos» sin
tocar SIRT). `verificar_funcional.py` compara el Estado real con el esperado.

| Día | Módulo | Filas | Cubre |
|---|---|---|---|
| 1 | Registro de establecidas | 52 | 15 servicios; alcances; tablas con y sin filas; fechas futuras; disponibilidad Lima / varias / vacía; sin registro definitivo; omitida; error de datos (alcance inválido) |
| 1 | Empaquetados | 10 | 2 a 6 tarifas; valores vacíos, escritos y mixtos; con y sin instalación; fechas; datos de contacto; error de alcance; error de datos; omitida |
| 1 | Licitaciones | 6 | 1, 3 y 6 servicios; instalación; Soles / Dólares / Euro; Público / Privado; inicio futuro; error de datos; omitida |
| 1 | Promociones | 12 | Todos / Los de siempre / Lista / Detalle; preferido Porcentual / Cantidad / Ilimitado; sin promocionar valor del plan; mezcla de servicios; empaquetado; empaquetado + monoproducto; error; omitida |
| 2 | Acciones | 20 | Fe de erratas (mono, paquete, LyC, promoción, tarifa ya renombrada); renovación (escrita, vacía, de paquete); cese (fecha vacía y escrita); renombre (por defecto, sufijo propio, doble, promoción); 4 errores; omitida |
| 2 | Modificaciones | 34 | Rectificación y cambio; delta vacío / entero / decimal / negativo; sentido; Por defecto / Todos / Solo la lista / Detalle (todas las rutas); agregar filas; ilimitado → limitado; canales sí/no; fecha de ejecución; sello; paquetes; LyC; 10 errores; omitida |

Total por serie: **134 filas** verificadas.

## 8.2 Resultados

| Serie | Día 1 | Día 2 | Resultado |
|---|---|---|---|
| FN (Prueba 1) | 60 filas · 1 error de SIRT (calendario abierto) → corregido | — | Encontró 1 fallo |
| FA (Prueba 2) | 60/60 tras corregir 3 fallos (instalación en paquetes de fijos, lista vacía de SIRT, promociones con servicios en común parciales) | — | Encontró 3 fallos |
| 1AP (Prueba 3) | 60/60 sin reintentos | 111/114 → 3 fallos corregidos (precio de servicios en paquetes ×2, «Todos los campos» en LyC) | Encontró 3 fallos |
| P4 (Prueba 4) | 80/80 en 26 min, sin intervención | 54/54 (con un corte de sesión, retomado) | **134/134** |
| P5 (Prueba 5) | 80/80 | 54/54 (con un corte de sesión, retomado) | **134/134** |

**Fallos encontrados y corregidos por las pruebas (11):** alcance sin validar · calendario que tapa el botón ·
instalación obligatoria con fijos · lista de SIRT que se abre vacía · promociones con solo algunos servicios en común ·
mensajes de error poco claros · valor de servicio > tarifa base al rectificar/cambiar paquetes · «Todos los campos»
tocaba un campo del paso anterior (contrato +0.01 por servicio) · prueba de 1 fila que se detenía por errores viejos ·
reintento de una renovación guardada justo antes de un corte · reanudación del día 1 sin regenerar los Excel.

La matriz de cobertura detallada (fila por fila, con «Qué se prueba», esperado y resultado de P4 y P5) está en
`entrega\matriz_de_cobertura.xlsx`; también se genera con `python verificar_funcional.py --carpeta Prueba4`.

## 8.3 Escenarios de volumen (prueba de carga)

Además se registraron los escenarios grandes de trabajo:

| Escenario | Cantidad | Estado |
|---|---|---|
| Tarifas establecidas | 149 (9–10 por servicio) + 45 de reserva | Ejecutado |
| Empaquetados | 101 (96 de 4 servicios, 1 de 6, 4 de una familia) | Ejecutado (100 + 1 omitido a propósito) |
| Licitaciones | 48 (6 de 1 servicio, 20 de 3, 15 de 4, 6 de 5, 1 de 6) | Ejecutado |
| Promociones | 55 | Ejecutado |
| Acciones básicas | 81 | Ejecutado |
| Modificaciones | 55 (`modificaciones.xlsx`) | Preparado; pendiente de ejecutar |

## 8.4 Tiempos medidos

| Operación | Tiempo |
|---|---|
| Registro de una tarifa fija | 9–13 s |
| Registro de una tarifa móvil | 20–28 s |
| Registro de un empaquetado | 20–28 s |
| Registro definitivo **de una en una** | 15–22 s por tarifa |
| Registro definitivo **en bloque** | 1–2 s por tarifa |
| Prueba funcional día 1 (80 filas, con registro definitivo en bloque) | 26 min |
| 129 tarifas establecidas con registro definitivo 1×1 (20/10) | ~1 h 15 min |
