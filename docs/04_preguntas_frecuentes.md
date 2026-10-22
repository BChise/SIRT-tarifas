# 4. Preguntas frecuentes y errores comunes

Todos estos casos se vieron en las pruebas reales (octubre de 2026).

## Antes de empezar

**«El Excel está abierto: ciérralo y vuelve a correr el comando.»**
El comando no empezó. Cerrar el Excel (también cualquier otra ventana de Excel con ese archivo) y repetir.

**«No se pudo conectar con Edge (puerto 9222).»**
Edge no está abierto en modo controlable. Cerrar Edge del todo y abrir `iniciar.bat`.

**«No se encontró una pestaña abierta de SIRT.»**
Edge está abierto pero sin SIRT. Abrir SIRT en ese Edge e iniciar sesión.

**Escribí el comando y dice que no existe el archivo `.py`.**
El cmd está en otra carpeta. Volver con `cd /d D:\proyecto_sirt\sirt_tarifas`.

## Mientras corre

**El proceso se quedó quieto.**
Casi siempre Edge está **minimizado**: restaurarlo (no minimizar; sí se puede tapar con otra ventana).

**Varias filas fallan seguidas en «abrir_tarifas_registradas» o «buscar», o aparece «Error al obtener las acciones
de la bandeja».**
La **sesión de SIRT venció**. Iniciar sesión de nuevo y correr el mismo comando: lo hecho no se repite. En
empaquetados el comando se detiene solo tras 3 errores seguidos en la misma etapa.

**¿Puedo usar la computadora mientras corre?**
Sí, con otras ventanas encima de Edge, pero sin minimizarlo ni hacer clic dentro. Con Win + L sigue corriendo.

**¿Puedo correr dos comandos a la vez?**
No se recomienda. Un segundo navegador o perfil obliga a iniciar otra sesión de SIRT y cierra la primera.
(Hay una opción experimental con dos pestañas del mismo Edge: ver [mantenimiento](06_mantenimiento.md).)

## Resultados

**La fila dice «No aplica».**
La tarifa no está «En comercialización» y «Vigente». SIRT la pasa a ese estado el día siguiente al registro
definitivo. Correr de nuevo mañana; no se tocó SIRT.

**La fila dice «Error de datos».**
El Detalle dice qué dato corregir (p. ej. «Alcance … no es una opción válida», «fecha … debe ser posterior a hoy»,
«un paquete necesita al menos 2 tarifas»). No se tocó SIRT.

**«SIRT no guardó el paso 1: Tarifa Instalación: Este apartado es obligatorio.»**
Con servicios fijos, SIRT exige la tarifa de instalación en el paquete: ponerle un valor.

**«El valor del servicio … no debe de ser mayor al valor del plan de la tarifa … que está siendo empaquetada.»**
En un paquete, cada servicio debe valer ≤ su tarifa base y la suma debe ser menor que la suma de las bases. Revisar los
«Valor N en el paquete» (vacío = base − 1). En modificaciones de paquetes, el script ya no sube esos valores.

**«El alcance pedido (…) no coincide con el de las tarifas (…): SIRT lo rechaza.»**
El paquete y todas sus tarifas deben tener el mismo alcance.

**«no aparece en Tarifas temporales».**
Ya tenía registro definitivo (se comprueba en «Tarifas registradas» y queda como «ya tenía registro definitivo»), o el
nombre no coincide.

**«no aparece en Tarifas registradas» en una acción.**
Revisar el nombre. Si la tarifa fue renombrada **fuera** de la herramienta, escribir su nombre actual.

**«El resumen no muestra N de los cambios».**
Medida de seguridad de modificaciones: no se guardó nada porque SIRT no mostró un cambio. Correr con `--diagnostico`
para ver las capturas del resumen y avisar a quien mantiene el proyecto.

**Una promoción mezcla servicios y el Detalle dice «solo muestra los servicios en común» o «no tienen apartados en
común».**
Comportamiento de SIRT: solo se promocionan los servicios compartidos (o solo el valor del plan / instalación).

**Una renovación quedó en error y al reintentar dice «renovación ya registrada».**
El intento anterior sí se guardó antes del corte; SIRT ya tiene la versión nueva «Por iniciar comercialización».
Está bien.

**Salió «Ya aplicada» / «Ya registrada».**
Ya estaba hecho en una corrida anterior; no se repitió. Conserva el código y la fecha originales.

**El «Código SIRT» está vacío.**
Se completa al hacer el registro definitivo. Para tarifas antiguas: `python completar_codigos_sirt.py`.

**Al final dice que el Excel está ABIERTO.**
Cerrarlo y escribir `si`. Si no se puede, escribir `copia`: los resultados quedan en `resultados\<MÓDULO>_EXCEL\`.

## Errores pasajeros de SIRT (ya cubiertos por el código)

| Visto | Qué hace ahora la herramienta |
|---|---|
| Un calendario se queda abierto y tapa «Guardar y continuar» | Lo cierra antes de guardar. |
| Una lista (p. ej. «Tipo de Tarifa») se abre vacía | La cierra y la vuelve a abrir (hasta 3 veces). |
| Pestaña de servicio que tarda en cargar | Espera hasta 20 s. |
| SIRT recuerda tarifas de un empaquetado fallido | Las quita antes de empezar. |
| Filtros de una búsqueda anterior | Pulsa «Limpiar filtro» antes de buscar. |
