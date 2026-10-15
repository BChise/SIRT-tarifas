# Referencia: proyecto anterior (Licitación y Contratos)

Scripts del proyecto anterior (operador START-UP COMPANY, tipo "Licitación y
Contratos", ambiente sirt.pe.be1business.com). Se guardan solo como
referencia: NO se ejecutan en este proyecto (otro ambiente, otro operador,
otro formulario).

Ese proyecto NO hacía disponibilidad comercial ni registro definitivo:
creaba la tarifa (hasta "Confirmar tarifa") y luego operaba sobre
"Tarifas registradas".

| Archivo | Qué hace | Qué nos sirve |
|---|---|---|
| `renombrar_tarifas.py` | Renombre por código SIRT + versión (agrega sufijo `_REN`) | **Base para nuestro renombre (próximo paso).** Acciones → "Renombre" → "Ingresa el nombre de la tarifa" → Guardar → Guardar tarifa → Aceptar; omite las ya renombradas |
| `listar_tarifas.py` | Recorre "Tarifas registradas" página por página y guarda código, versión, nombre, estado | Paginador a 100 filas, lectura de celdas, espera robusta a que cambie la página. Útil para obtener el **código SIRT** de nuestras tarifas |
| `rectificar_tarifas.py` | Rectificar (cambia ámbito y agrega marca en textos) | Flujo de menú + doble "Aceptar". Los campos dependen del servicio: pendiente |
| `fe_erratas_tarifas.py` | Fe de erratas sobre la versión más alta | Elegir la versión más alta de un código. Campos dependen del servicio: pendiente |
| `obtener_modalidades.py` | Abre "Ficha informativa" y lee su texto | Leer la ficha de una tarifa para verificar datos |
| `validar_csv.py` | Separa las tarifas de un lote por patrón de nombre y detecta faltantes | Verificar que un lote esté completo |

Piezas comunes (repetidas en varios scripts):
- Buscar por código: "Filtro avanzado" → textbox "Ingresa código tarifa" → "Buscar"
  y esperar a que el código aparezca en `tbody tr`.
- Menú de una fila: `fila.get_by_label("Acciones").click()` y luego
  `page.get_by_role("menuitem", name="...")`.
- Columnas de "Tarifas registradas" (índice de celda): 1 fecha, 2 código SIRT,
  3 versión, 4 tipo, 5 condición, 6 nombre, 7 alcance, 8 usuario, 9 estado, 10 prioridad.

No se copiaron `grabacion_renomb.py` ni `grabacion_renombrar.py`: contienen una
contraseña real en texto plano (cambiarla y borrar esos archivos).
