# Registro de cambios

## v1.0 — 22/10/2026 (entrega a analistas)

Primera versión entregada. Probada en SIRT QA con dos pruebas funcionales completas (134/134 filas cada una) y los
escenarios de volumen.

### Funciones
- Registro de **tarifas establecidas** de los 15 servicios desde Excel (tablas de países, velocidades y bolsas;
  disponibilidad comercial; fechas futuras).
- **Empaquetados** de 2 a 6 tarifas, con validación previa de alcance y reglas de precio.
- **Licitaciones y contratos** de 1 a 6 servicios.
- **Promociones** de 1 a 15 tarifas (monoproductos y empaquetados): atributos Todos / Lista / Detalle, tipo de
  variación preferido, servicios en común.
- **Registro definitivo** de una en una o **en bloque** (disponibilidad y registro de 10 tarifas a la vez; 1–2 s por
  tarifa en vez de 15–22 s).
- **Acciones**: fe de erratas, renovación, cese y renombre (con resolución de tarifas ya renombradas).
- **Modificaciones**: rectificación y cambio de tarifa con delta, sentido beneficioso/perjudicial, campos por
  defecto / todos / lista / detalle por campo (rutas `tabla › fila › campo`), filas nuevas, ilimitado → limitado,
  canales.
- **Código SIRT** e **ID de registro** en todos los Excel.
- Reanudación segura: volver a correr un comando continúa donde quedó.
- **Prueba funcional** automatizada (`prueba_dia1.bat`, `prueba_dia2.bat`) y matriz de cobertura
  (`verificar_funcional.py`).

### Correcciones durante la validación (21–22/10/2026)
- Validación del alcance en el registro.
- Cierre de calendarios que tapaban «Guardar y continuar».
- Mensaje claro cuando SIRT no guarda el paso 1 (campo obligatorio).
- Reintento de listas de SIRT que se abren vacías.
- Promociones con solo algunos servicios en común.
- Modificaciones de empaquetados: el valor de los servicios no sube (cada servicio ≤ su tarifa base).
- «Todos los campos» ya no toca campos de pasos ocultos (valor de contrato en licitaciones).
- Reconocimiento de renovaciones guardadas justo antes de un corte de sesión.
- Aviso y reintento al guardar el Excel si está abierto.

### Dependencias
- Python 3.13.2 · playwright 1.64.0 · openpyxl 3.1.5.

## Antes de v1.0 (13–20/10/2026)
Desarrollo por servicios con `ejecutar.py` (combinaciones), primeras versiones de empaquetados, promociones,
licitaciones, acciones y la app web local (congelada). Ver `CONTEXTO.md`.
