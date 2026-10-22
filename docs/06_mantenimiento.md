# 6. Mantenimiento

## 6.1 Si SIRT cambia una pantalla

Síntoma típico: un módulo que funcionaba empieza a fallar siempre en la misma etapa (el Detalle empieza con
`[etapa]`, p. ej. `[paso_1_guardado]`, `[renovar]`) y la captura en `resultados\…\capturas_EXCEL\` muestra la
pantalla distinta.

1. Correr 1 fila con `--diagnostico` (capturas de cada paso en `resultados\…\diagnostico_EXCEL\`).
2. **Grabar el flujo nuevo**: con Edge abierto por `iniciar.bat` y la sesión iniciada, `python grabar_sirt.py`, en el
   Inspector pulsar **Record**, hacer el flujo a mano y copiar el código a
   `grabaciones\<flujo>\grabacion.py`. Agregar capturas y un `notas.txt` con lo observado.
   No capturar las herramientas de desarrollador (F12): muestran el token de sesión.
3. Si el formulario tiene listas o campos nuevos: guardar el JSON del GET del paso 2 en
   `grabaciones\<servicio>\paso2.json` y resumirlo con `python inspeccionar_json.py <servicio>`.
4. Ajustar el script del módulo (ver [documentación técnica](07_documentacion_tecnica.md) para saber cuál) usando
   selectores por rol y texto visible.
5. Volver a probar: `--solo-listar` → `--limite 1` → una **prueba funcional** completa.

## 6.2 Prueba funcional (recomendada tras cualquier cambio)

```
prueba_dia1.bat SERIE CARPETA        (p. ej. prueba_dia1.bat P8 Prueba8)
prueba_dia2.bat CARPETA              (al día siguiente o después)
```

- Crea tarifas nuevas en SIRT (~52 + 10 + 6 + 12 por serie): usar **QA** y una serie que no se haya usado.
- Si se detiene, revisar el motivo, corregir y lanzar **el mismo comando** (continúa donde quedó).
- Resultado: `python verificar_funcional.py --carpeta CARPETA` → «OK n · FALLA 0 · PENDIENTE 0».
- **Prueba rápida periódica** (semanal o cuando OSIPTEL anuncie cambios): el día 1 de una serie nueva (~26 min).

## 6.3 Agregar un servicio nuevo

1. Copiar `servicios\_plantilla.py` como `servicios\<nombre>.py`; completar `SERVICIO` (texto exacto de «Seleccionar
   servicio»), `CODIGO`, `DIMENSIONES` (listas y opciones exactas del JSON) y `completar_paso_2`.
2. Bloques comunes nuevos (tablas, condicionales) van a `comun.py`.
3. Agregar la hoja del servicio al Excel de registro con `automatizacion\generar_plantilla_registro.py`.
4. Si pide disponibilidad comercial: `REQUIERE_DISPONIBILIDAD` en `registro_definitivo.py`.
5. Probar con `--limite 1`.

## 6.4 Ajustes frecuentes en el código

| Qué | Dónde |
|---|---|
| Días de anticipación de un cambio de tarifa | `acciones.py`: `DIAS_ANTICIPACION_ALZA` / `_BAJA` (hoy 3) |
| Disponibilidad por defecto del registro | `registro_excel.py`: `DISPONIBILIDAD_POR_DEFECTO` («Lima») |
| Tamaño máximo del bloque de registro definitivo | `registro_definitivo.py`: `TAMANO_BLOQUE_MAX` (50) |
| Servicios que piden disponibilidad | `registro_definitivo.py`: `REQUIERE_DISPONIBILIDAD` |
| Campos que «Todos los campos» nunca toca | `acciones.py`: `NUNCA_EN_TODO` |
| Qué cuenta como precio (sentido beneficioso/perjudicial) | `acciones.py`: expresión `PRECIO` |
| Errores seguidos antes de detenerse (empaquetados) | `empaquetados_excel.py`: `ERRORES_SEGUIDOS_MAX` (3) |

## 6.5 Versiones y git

- Repositorio **privado**. `resultados\` y los bloqueos de Excel están en `.gitignore`.
- Cada versión entregada se etiqueta (`v1.0`, `v1.1`…) y se anota en `CHANGELOG.md`:
  ```
  git add -A
  git commit -m "v1.1: <resumen>"
  git tag v1.1
  git push && git push --tags
  ```
- Actualizar las dependencias solo después de una prueba funcional con la versión nueva; anotar las versiones en
  `requirements.txt`.

## 6.6 Opciones experimentales (no usadas en producción)

- **Dos procesos a la vez en el mismo Edge**: duplicar la pestaña de SIRT en otra ventana (ambas visibles, sin
  minimizar) y en el segundo cmd `set SIRT_PESTANA=2`. Cada proceso debe trabajar tarifas distintas. Otro perfil o
  navegador no sirve: pide iniciar sesión de nuevo y cierra la primera.
- `iniciar_paralelo.bat` (puerto 9223, perfil `sirt2`): queda obsoleto por lo anterior.
