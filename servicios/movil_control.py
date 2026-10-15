"""
Servicio: Servicio público móvil control.

Fuente: grabaciones/movil_control/ (paso2.json, grabacion_paso2.py, 12 capturas).

El paso 2 es IDÉNTICO al de Servicio público móvil postpago (mismos campos,
listas y tipos según paso2.json). Por eso este archivo reutiliza todo de
movil_postpago.py y solo cambia el nombre y el código del servicio:
una corrección en movil_postpago.py vale para los dos.

Si algún día control necesita valores distintos, se copian aquí y se cambian,
por ejemplo:
    VALORES = {**movil_postpago.VALORES, "valor_servicio": "30"}

Este servicio no tiene "Tarifa Instalación" (registraTarifaInstalacion = false).
"""

from servicios import movil_postpago


SERVICIO = "Servicio público móvil control"   # texto exacto
CODIGO = "SMC"

DIMENSIONES = movil_postpago.DIMENSIONES
VALORES = movil_postpago.VALORES
completar_paso_2 = movil_postpago.completar_paso_2
