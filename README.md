# Automatización de tarifas SIRT (OSIPTEL)

Registra y modifica tarifas en **SIRT** a partir de **Excel**, manejando Microsoft Edge con Playwright (Python).
Versión **1.0** — ver [CHANGELOG](CHANGELOG.md).

## Empezar

1. Instalar: [docs/01_instalacion.md](docs/01_instalacion.md).
2. Cada vez: doble clic en `iniciar.bat`, iniciar sesión en SIRT, cerrar el Excel.
3. Correr el comando del Excel que se va a usar:

| Tarea | Excel (`automatizacion\`) | Comando |
|---|---|---|
| Tarifas establecidas | `registro_tarifas_establecidas.xlsx` | `python registro_excel.py --si --definitivo-en-bloque` |
| Empaquetados | `empaquetados.xlsx` | `python empaquetados_excel.py --si --definitivo-en-bloque` |
| Licitaciones y contratos | `registro_licitaciones.xlsx` | `python licitaciones_excel.py --si --definitivo-en-bloque` |
| Promociones | `promociones.xlsx` | `python promociones_excel.py --si --definitivo-en-bloque` |
| Fe de erratas, renovación, cese, renombre | `acciones_basicas.xlsx` | `python acciones_excel.py --si` |
| Rectificación y cambio de tarifa | `modificaciones.xlsx` | `python modificaciones_excel.py --si` |

Siempre primero `--solo-listar` (no toca nada) y `--limite 1 --si`.

## Documentación

Todo junto en PDF: **`Manual_SIRT_v1.0.pdf`**.


| | |
|---|---|
| [1. Instalación](docs/01_instalacion.md) | Requisitos, entorno, primer arranque |
| [2. Manual de usuario](docs/02_manual_usuario.md) | Cada Excel, sus columnas, comandos y resultados |
| [3. Referencia rápida](docs/03_referencia_rapida.md) | Una página para imprimir |
| [4. Preguntas frecuentes y errores](docs/04_preguntas_frecuentes.md) | Qué significa cada mensaje y qué hacer |
| [5. Limitaciones conocidas](docs/05_limitaciones.md) | Qué no hace esta versión y reglas de SIRT |
| [6. Mantenimiento](docs/06_mantenimiento.md) | Si SIRT cambia, agregar servicios, versiones |
| [7. Documentación técnica](docs/07_documentacion_tecnica.md) | Arquitectura, archivos, resultados, diseño |
| [8. Pruebas y cobertura](docs/08_pruebas_y_cobertura.md) | Qué se probó, resultados y tiempos |
| [9. Guiones de video](docs/09_guiones_video.md) | Qué mostrar y decir en cada video |

Plantillas vacías, Excel de ejemplo y la matriz de cobertura: carpeta [`entrega\`](entrega/).

## Seguridad
Ningún archivo guarda usuario ni contraseña (el inicio de sesión es manual). El repositorio debe ser privado.
`LEEME.md`, `TUTORIAL_APP.md` y `CONTEXTO.md` (solo en el repositorio) son documentos de la etapa de desarrollo.
