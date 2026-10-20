// SIRT · Herramienta de pruebas — interfaz (sin dependencias).
"use strict";

const $ = (sel, raiz = document) => raiz.querySelector(sel);
const $$ = (sel, raiz = document) => Array.from(raiz.querySelectorAll(sel));
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

let CATALOGO = { servicios: [], campos: {}, tarifas: [], lotes: [] };
let fuente = null;          // EventSource de la consola
let tareaVista = null;      // id de la tarea que muestra la consola

// ---------------------------------------------------------------- utilidades
function aviso(texto, ms = 4000) {
  const caja = $("#aviso");
  caja.textContent = texto;
  caja.classList.remove("oculto");
  clearTimeout(aviso.t);
  aviso.t = setTimeout(() => caja.classList.add("oculto"), ms);
}

async function api(ruta, opciones = {}) {
  const r = await fetch(ruta, { headers: { "Content-Type": "application/json" }, ...opciones });
  const datos = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(datos.detail ? (typeof datos.detail === "string" ? datos.detail : JSON.stringify(datos.detail)) : r.statusText);
  return datos;
}

function codigoDe(nombre) {
  const m = /^T[EP]-([A-Z]+)-/.exec(nombre || "");
  return m ? m[1] : "";
}

// ---------------------------------------------------------------- pestañas
function mostrarVista(nombre) {
  $$("#pestanas button").forEach((b) => b.classList.toggle("activa", b.dataset.vista === nombre));
  $$(".vista").forEach((v) => v.classList.toggle("activa", v.id === `vista-${nombre}`));
  try { localStorage.setItem("vista", nombre); } catch (e) {}
  if (nombre === "resultados") cargarListaResultados();
}
$("#pestanas").addEventListener("click", (e) => {
  if (e.target.dataset.vista) mostrarVista(e.target.dataset.vista);
});

// ---------------------------------------------------------------- campos condicionales (data-si)
function actualizarCondicionales(form) {
  const datos = new FormData(form);
  $$("[data-si]", form).forEach((el) => {
    const [campo, valores] = el.dataset.si.split("=");
    const actual = datos.get(campo) ?? "";
    el.classList.toggle("oculto", !valores.split(",").includes(actual));
  });
}

// ---------------------------------------------------------------- servicios
function pintarServicios() {
  $$("[data-servicios]").forEach((caja) => {
    const marcar = !caja.hasAttribute("data-sin-marcar");
    caja.innerHTML = CATALOGO.servicios.map((s) => `
      <label class="check"><input type="checkbox" name="servicios" value="${esc(s.codigo)}" ${marcar ? "checked" : ""}>
        <b>${esc(s.codigo)}</b>&nbsp;<span>${esc(s.nombre)}</span></label>`).join("");
  });
}
document.addEventListener("click", (e) => {
  const marca = e.target.dataset.marcar, desmarca = e.target.dataset.desmarcar;
  if (!marca && !desmarca) return;
  const form = e.target.closest("form");
  $$(`input[name="${marca || desmarca}"]`, form).forEach((c) => (c.checked = !!marca));
});

// ---------------------------------------------------------------- selector de tarifas
function pintarSelectores() {
  $$("[data-selector]").forEach((caja) => {
    const origenes = (caja.dataset.origen || "").split(",").filter(Boolean);
    const previos = caja._sel || new Set();
    const manualPrevio = caja.querySelector(".manuales")?.value || "";
    const filtroPrevio = caja.querySelector("input[type=search]")?.value || "";
    const soloDef = caja.querySelector(".solo-def")?.checked ?? caja.hasAttribute("data-definitivas");
    const origenesMarcados = new Set($$(".origen:checked", caja).map((c) => c.value));
    const usar = origenesMarcados.size ? origenesMarcados : new Set(origenes);
    caja.innerHTML = `
      <div class="cabeza">
        <b>Tarifas</b>
        <input type="search" placeholder="Buscar por nombre o código (SMP, EMP…)" value="${esc(filtroPrevio)}">
        ${origenes.map((o) => `<label class="check"><input type="checkbox" class="origen" value="${o}" ${usar.has(o) ? "checked" : ""}> ${o}</label>`).join("")}
        <label class="check"><input type="checkbox" class="solo-def" ${soloDef ? "checked" : ""}> solo con registro definitivo</label>
        <button type="button" class="enlace marcar-visibles">marcar visibles</button>
        <button type="button" class="enlace limpiar">limpiar</button>
        <span class="contador"></span>
      </div>
      <div class="lista"></div>
      <label class="bloque" style="margin:8px 0 0">Otras tarifas (una por línea, si no aparecen arriba)
        <textarea class="manuales" rows="2" placeholder="tar_TM_postpago"></textarea></label>`;
    caja._sel = previos;                       // selección que se conserva al filtrar
    caja.querySelector(".manuales").value = manualPrevio;
    filtrarSelector(caja);
    if (caja._listo) return;
    caja._listo = true;
    caja.addEventListener("input", (e) => {
      if (e.target.matches("input[type=search], .origen, .solo-def")) filtrarSelector(caja);
      if (e.target.matches(".manuales")) avisarCambio(caja);
    });
    caja.addEventListener("change", (e) => {
      if (e.target.matches(".lista input")) {
        e.target.checked ? caja._sel.add(e.target.value) : caja._sel.delete(e.target.value);
        contarSeleccion(caja);
      }
      avisarCambio(caja);
    });
    caja.addEventListener("click", (e) => {
      if (e.target.classList.contains("marcar-visibles")) {
        $$(".lista input", caja).forEach((c) => { c.checked = true; caja._sel.add(c.value); });
      } else if (e.target.classList.contains("limpiar")) {
        caja._sel.clear(); $$(".lista input", caja).forEach((c) => (c.checked = false));
        caja.querySelector(".manuales").value = "";
      } else return;
      contarSeleccion(caja); avisarCambio(caja);
    });
  });
}

function filtrarSelector(caja) {
  const texto = caja.querySelector("input[type=search]").value.trim().toLowerCase();
  const origenes = new Set($$(".origen:checked", caja).map((c) => c.value));
  const soloDef = caja.querySelector(".solo-def").checked;
  const marcadas = caja._sel || (caja._sel = new Set());
  const visibles = CATALOGO.tarifas.filter((t) =>
    origenes.has(t.origen) && (!soloDef || t.definitivo) &&
    (!texto || t.nombre.toLowerCase().includes(texto) || (t.codigo || "").toLowerCase() === texto));
  caja.querySelector(".lista").innerHTML = visibles.map((t) => `
    <label class="check"><input type="checkbox" value="${esc(t.nombre)}" ${marcadas.has(t.nombre) ? "checked" : ""}>
      ${esc(t.nombre)} ${t.definitivo ? "" : '<span class="etiqueta">temporal</span>'}</label>`).join("")
    || '<span class="contador">Sin coincidencias.</span>';
  contarSeleccion(caja, visibles.length);
}

function contarSeleccion(caja, visibles) {
  if (visibles !== undefined) caja._visibles = visibles;
  caja.querySelector(".contador").textContent =
    `${caja._visibles ?? 0} de ${CATALOGO.tarifas.length} · ${caja._sel?.size || 0} elegida(s)`;
}

function tarifasDe(caja) {
  const marcadas = [...(caja._sel || [])];
  const manuales = (caja.querySelector(".manuales")?.value || "").split(/[\n,]+/).map((t) => t.trim()).filter(Boolean);
  return [...new Set([...marcadas, ...manuales])];
}

function avisarCambio(caja) {
  caja.dispatchEvent(new CustomEvent("tarifas", { bubbles: true }));
}

// ---------------------------------------------------------------- árbol de campos (acciones)
const GENERICOS = [
  { nombre: "Otros datos generales", paso: 1, tipo: "texto", ok: true },
  { nombre: "Tarifa de instalación", paso: 1, tipo: "numero", ok: true },
  { nombre: "Valor del servicio", paso: 2, tipo: "numero", ok: true },
  { nombre: "Condiciones generales", paso: 2, tipo: "texto", ok: true },
  { nombre: "Restricciones técnicas", paso: 2, tipo: "texto", ok: true },
];

function pintarArbol() {
  const form = $("#vista-acciones");
  const accion = form.accion.value;
  const modo = form.modo.value;
  const arbol = $("#arbol-campos");
  const mostrar = ["fe_erratas", "rectificar", "cambio"].includes(accion) && modo === "elegir";
  arbol.classList.toggle("oculto", !mostrar);
  if (!mostrar) return;

  const previos = new Set($$("input.campo:checked", arbol).map((c) => c.value));
  const tarifas = tarifasDe($("[data-selector]", form));
  const codigos = [...new Set(tarifas.map(codigoDe).map((c) => c || "OTRAS"))];
  if (!codigos.length) { arbol.innerHTML = '<p class="ayuda">Elige primero las tarifas: aquí aparecerán los campos de sus servicios.</p>'; return; }

  const porAccion = CATALOGO.campos[accion] || {};
  const nombreServicio = Object.fromEntries(CATALOGO.servicios.map((s) => [s.codigo, s.nombre]));
  arbol.innerHTML = codigos.map((cod) => {
    let campos = porAccion[cod];
    let titulo = `${cod} · ${nombreServicio[cod] || ""}`;
    if (!campos) {
      campos = GENERICOS.map((c) => ({ ...c, ok: accion === "fe_erratas" ? c.tipo === "texto" : true,
        motivo: accion === "fe_erratas" && c.tipo !== "texto" ? "fe de erratas solo corrige textos" : "" }));
      titulo = cod === "EMP" ? "EMP · Empaquetados (se aplica en cada pestaña de servicio)"
        : cod === "OTRAS" ? "Tarifas sin código reconocible (campos comunes)" : `${cod} · campos comunes`;
    }
    const usables = campos.filter((c) => c.ok).length;
    const n = tarifas.filter((t) => (codigoDe(t) || "OTRAS") === cod).length;
    return `<details ${codigos.length <= 2 ? "open" : ""}>
      <summary><input type="checkbox" class="todos-servicio" title="marcar todos"> ${esc(titulo)}
        <span class="contador">${n} tarifa(s) · ${usables}/${campos.length} campos modificables</span></summary>
      <div class="campos">${campos.map((c) => `
        <label class="check ${c.ok ? "" : "deshab"}">
          <input type="checkbox" class="campo" value="${esc(c.nombre)}" ${c.ok ? "" : "disabled"} ${c.ok && previos.has(c.nombre) ? "checked" : ""}>
          <span class="paso">P${c.paso}</span> ${esc(c.nombre)}
          ${c.ok ? "" : `<span class="motivo">— ${esc(c.motivo)}</span>`}</label>`).join("")}
      </div></details>`;
  }).join("");
  $$(".todos-servicio", arbol).forEach((caja) => caja.addEventListener("change", (e) => {
    e.stopPropagation();
    $$("input.campo:not(:disabled)", caja.closest("details")).forEach((c) => (c.checked = caja.checked));
  }));
  $$(".todos-servicio", arbol).forEach((caja) => caja.addEventListener("click", (e) => e.stopPropagation()));
}

// ---------------------------------------------------------------- envío de tareas
function datosDe(form) {
  const fd = new FormData(form);
  const tipo = form.dataset.tarea;
  const bool = (n) => !!form.querySelector(`[name="${n}"]`)?.checked;
  const lista = (texto) => (texto || "").split(/[\n,]+/).map((t) => t.trim()).filter(Boolean);
  const selector = form.querySelector("[data-selector]");
  if (tipo === "establecidas")
    return { servicios: fd.getAll("servicios"), lote: fd.get("lote"), limite: +fd.get("limite") || 1, diagnostico: bool("diagnostico") };
  if (tipo === "definitivo")
    return { lote: fd.get("lote"), servicios: fd.getAll("servicios"), nombres: lista(fd.get("nombres")),
             limite: +fd.get("limite") || 0, ubigeos: fd.get("ubigeos") || "", diagnostico: bool("diagnostico") };
  if (tipo === "empaquetar")
    return { tarifas: tarifasDe(selector), lote: fd.get("lote"), nombre: fd.get("nombre") || "",
             valores: lista(fd.get("valores")), diagnostico: bool("diagnostico") };
  if (tipo === "promocionar") {
    const modo = fd.get("atributos_modo");
    return { tarifas: tarifasDe(selector), lote: fd.get("lote"), una_por_tarifa: bool("una_por_tarifa"),
             atributos: modo === "lista" ? (fd.get("atributos_lista") || "probados") : modo,
             variacion_porcentual: fd.get("variacion_porcentual"), variacion_cantidad: fd.get("variacion_cantidad"),
             duracion: fd.get("duracion"), dias_comercializacion: +fd.get("dias_comercializacion") || 60,
             sin_valor_plan: bool("sin_valor_plan"), sin_instalacion: bool("sin_instalacion"), diagnostico: bool("diagnostico") };
  }
  // acciones
  return { accion: fd.get("accion"), tarifas: tarifasDe(selector), hasta_resumen: bool("hasta_resumen"),
           diagnostico: bool("diagnostico"), repetir: bool("repetir"), sin_validar: bool("sin_validar"),
           sufijo: fd.get("sufijo") || "_REN", nuevo_nombre: fd.get("nuevo_nombre") || "", fecha: fd.get("fecha") || "",
           delta: fd.get("delta") || "", texto: fd.get("texto") || "", modo: fd.get("modo") || "defecto",
           campos: $$("#arbol-campos input.campo:checked").map((c) => c.value) };
}

function confirmacion(tipo, datos) {
  if (tipo === "accion") {
    if (datos.accion === "buscar") return true;
    if (["fe_erratas", "rectificar", "cambio"].includes(datos.accion) && datos.hasta_resumen) return true;
    return confirm(`Se aplicará «${datos.accion}» en SIRT a ${datos.tarifas.length} tarifa(s) y se GUARDARÁ.\n\n${datos.tarifas.join("\n")}\n\n¿Continuar?`);
  }
  const textos = {
    establecidas: `Se registrarán tarifas establecidas de ${datos.servicios?.length} servicio(s) en el lote ${datos.lote}.`,
    definitivo: `Se registrarán en DEFINITIVO las tarifas del lote ${datos.lote}.`,
    empaquetar: `Se creará un empaquetado con ${datos.tarifas?.length} tarifas.`,
    promocionar: `Se crearán promociones sobre ${datos.tarifas?.length} tarifa(s).`,
  };
  return confirm(`${textos[tipo]}\n\n¿Continuar?`);
}

$$("form.vista").forEach((form) => {
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const tipo = form.dataset.tarea;
    const datos = datosDe(form);
    if ("tarifas" in datos && !datos.tarifas.length) return aviso("Elige al menos una tarifa.");
    if (!confirmacion(tipo, datos)) return;
    const boton = form.querySelector("button[type=submit]");
    boton.disabled = true;
    try {
      const tarea = await api(`/api/tareas/${tipo}`, { method: "POST", body: JSON.stringify(datos) });
      aviso(`Tarea iniciada: ${tarea.titulo}`);
      verTarea(tarea.id, tarea.titulo);
      refrescarEstado();
    } catch (error) {
      aviso(error.message, 7000);
    } finally {
      boton.disabled = false;
    }
  });
  form.addEventListener("change", () => { actualizarCondicionales(form); if (form.id === "vista-acciones") pintarArbol(); });
  form.addEventListener("tarifas", () => { if (form.id === "vista-acciones") pintarArbol(); });
  actualizarCondicionales(form);
});

// ---------------------------------------------------------------- consola y estado
function verTarea(id, titulo) {
  if (fuente) fuente.close();
  tareaVista = id;
  const consola = $("#consola");
  consola.textContent = "";
  $("#titulo-consola").textContent = titulo ? `Consola · #${id} ${titulo}` : `Consola · #${id}`;
  fuente = new EventSource(`/api/tareas/${id}/salida`);
  fuente.onmessage = (e) => {
    const abajo = consola.scrollTop + consola.clientHeight >= consola.scrollHeight - 30;
    consola.textContent += JSON.parse(e.data) + "\n";
    if (abajo) consola.scrollTop = consola.scrollHeight;
  };
  fuente.addEventListener("fin", () => { fuente.close(); fuente = null; refrescarEstado(); cargarCatalogo(false); });
  fuente.onerror = () => { if (fuente) { fuente.close(); fuente = null; } };
}

$("#btn-cancelar").onclick = async () => {
  if (!tareaVista || !confirm("¿Cancelar la tarea en curso? Lo que ya se guardó en SIRT no se deshace.")) return;
  try { await api(`/api/tareas/${tareaVista}/cancelar`, { method: "POST" }); } catch (e) { aviso(e.message); }
};

async function refrescarEstado() {
  try {
    const estado = await api("/api/estado");
    const edge = $("#estado-edge");
    edge.textContent = estado.edge ? "Edge conectado" : "Edge NO conectado (puerto 9222)";
    edge.className = `pildora ${estado.edge ? "ok" : "mal"}`;
    const t = $("#estado-tarea");
    t.textContent = estado.tarea ? `En curso: ${estado.tarea.titulo}` : "Sin tareas en curso";
    t.className = `pildora ${estado.tarea ? "ocupado" : ""}`;
    $("#btn-cancelar").disabled = !estado.tarea || estado.tarea.id !== tareaVista;
    const tareas = await api("/api/tareas");
    $("#historial").innerHTML = tareas.map((x) => `
      <li data-id="${x.id}" data-titulo="${esc(x.titulo)}"><span>#${x.id} ${esc(x.titulo)}</span>
        <span class="${x.estado.replace(" ", " ")}">${esc(x.estado)} · ${x.inicio}</span></li>`).join("")
      || '<li class="contador">Todavía no hay tareas.</li>';
  } catch (e) { /* servidor reiniciándose */ }
}
$("#historial").addEventListener("click", (e) => {
  const li = e.target.closest("li[data-id]");
  if (li) verTarea(+li.dataset.id, li.dataset.titulo);
});

// ---------------------------------------------------------------- resultados
async function cargarListaResultados() {
  const sel = $("#sel-resultado");
  const actual = sel.value;
  const archivos = await api("/api/resultados");
  sel.innerHTML = archivos.map((a) => `<option value="${esc(a.ruta)}">${esc(a.ruta)} · ${a.modificado}</option>`).join("");
  if (actual && archivos.some((a) => a.ruta === actual)) sel.value = actual;
  if (sel.value) mostrarResultado();
}
let RESULTADO = null;
async function mostrarResultado() {
  RESULTADO = await api(`/api/resultados/ver?ruta=${encodeURIComponent($("#sel-resultado").value)}`);
  pintarResultado();
}
function pintarResultado() {
  if (!RESULTADO) return;
  const filtro = $("#filtro-resultado").value.trim().toLowerCase();
  const filas = RESULTADO.filas.filter((f) => !filtro || f.join(" ").toLowerCase().includes(filtro));
  const iRes = RESULTADO.columnas.indexOf("resultado");
  $("#info-resultado").textContent = `${RESULTADO.total} filas en total · se muestran ${filas.length} (las más recientes primero)`;
  $("#tabla-resultado").innerHTML =
    `<thead><tr>${RESULTADO.columnas.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead>` +
    `<tbody>${filas.map((f) => `<tr>${f.map((v, i) => `<td class="${i === iRes ? esc(v) : ""}">${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody>`;
}
$("#sel-resultado").addEventListener("change", mostrarResultado);
$("#filtro-resultado").addEventListener("input", pintarResultado);
$("#btn-recargar").addEventListener("click", cargarListaResultados);

// ---------------------------------------------------------------- arranque
async function cargarCatalogo(primeraVez = true) {
  try {
    CATALOGO = await api("/api/catalogo");
  } catch (e) { aviso("No se pudo leer el catálogo: " + e.message); return; }
  $("#lista-lotes").innerHTML = CATALOGO.lotes.map((l) => `<option value="${esc(l)}">`).join("");
  if (primeraVez) pintarServicios();
  pintarSelectores();
  pintarArbol();
}

(async function iniciar() {
  let vista = "establecidas";
  try { vista = localStorage.getItem("vista") || vista; } catch (e) {}
  mostrarVista(vista);
  await cargarCatalogo(true);
  refrescarEstado();
  setInterval(refrescarEstado, 3000);
})();
