# 5. Limitaciones conocidas

## No disponible en esta versión (v1.0)

| Qué | Situación |
|---|---|
| Fechas pasadas | Los Excel rechazan inicios de vigencia / comercialización anteriores a hoy. Falta confirmar si SIRT los permite. |
| Rectificación y cambio de **promociones** | No están grabados; esas filas quedan «Error de datos». |
| «Pasar ilimitado a limitado» al revés (limitado → ilimitado) | No implementado. |
| Carga masiva de disponibilidad por archivo | Se usa el «Selector geográfico»; la opción «Carga masiva» de SIRT no se usa. |
| Trabajo en paralelo | Una sola sesión (ver mantenimiento: opción experimental con dos pestañas). |
| Ambiente de producción | Todo se probó en **SIRT QA**. Antes de usarlo en producción, hacer una prueba funcional mínima allí. |

## Comportamientos de SIRT que la herramienta respeta

- **Registro definitivo**: no se puede deshacer.
- **Acciones y modificaciones**: requieren tarifa «En comercialización» y «Vigente» (desde el día siguiente al registro
  definitivo); licitaciones, solo vigente. Si no, la fila queda «No aplica».
- **Un solo cambio de tarifa pendiente** por tarifa.
- **Empaquetados**: mismo alcance que sus tarifas, cada servicio ≤ su base, suma < suma de bases, instalación
  obligatoria con servicios fijos. En modificaciones de paquetes el valor de los servicios no sube.
- **Promociones mezcladas**: solo se promocionan los servicios en común.
- **Hoy siempre está bloqueado** en los calendarios de acciones: se usa el primer día permitido.
- **Fecha de ejecución de un cambio**: SIRT pide al menos 2 días si el cambio perjudica al usuario (por defecto se usan
  3).

## Uso del equipo

- Edge **no** debe minimizarse mientras corre (Windows lo congela).
- La sesión de SIRT vence: en corridas largas puede que haya que iniciar sesión y relanzar (es seguro).
- Los Excel deben estar cerrados al empezar y conviene no abrirlos mientras corren.

## Qué se probó y qué no

- Se probó cada opción de cada Excel al menos una vez, en todos los tipos de tarifa donde aplica, con resultado
  verificado (ver [pruebas y cobertura](08_pruebas_y_cobertura.md)).
- **No** se probaron todas las combinaciones posibles (opción × servicio × tipo), y no se probaron nombres con
  caracteres especiales ni de 100 caracteres, ni más de 20 tarifas pendientes sin parte común en el nombre durante el
  registro definitivo en bloque (lo que no se ve en el filtro pasa al método de una en una).
- Si SIRT cambia una pantalla o una regla, el módulo afectado puede fallar hasta ajustarlo
  (ver [mantenimiento](06_mantenimiento.md)).
