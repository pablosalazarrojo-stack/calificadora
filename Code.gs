/**
 * La calificadora — el "cerebro" que vive en la hoja de Google.
 * Guarda las calificaciones de cada evaluador y los nombres de los equipos.
 * La contraseña del panel del profe vive aquí (Propiedades del script), no en la página.
 *
 * Instalación (una sola vez):
 *  1. Hoja de Google nueva → Extensiones → Apps Script → pegar este archivo.
 *  2. Poner el id de la hoja en SHEET_ID y ejecutar la función `setup` (pide permisos de tu cuenta).
 *  3. Implementar → Nueva implementación → Aplicación web
 *     · Ejecutar como: yo   · Quién tiene acceso: cualquier persona
 *  4. Copiar la URL que termina en /exec y pegarla en CONFIG.SHEET_URL de index.html.
 */

const SHEET_ID = "1BOXpitL5DfwTORE-zp125eKF00AdoiAcZKBihHOQVuk";
const HOJA = "Votos";
const COLS = ["id", "fecha", "nombre", "puntajes"];
const CRIT = ["cre", "usa", "con", "ia"];
const N_EQUIPOS = 4;
const DEFAULT = ["equipo 1", "equipo 2", "equipo 3", "equipo 4"];

function libro() { return SpreadsheetApp.openById(SHEET_ID); }
function hoja() { return libro().getSheetByName(HOJA); }

function setup() {
  const ss = libro();
  const h = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
  h.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  h.setFrozenRows(1);
  const s1 = ss.getSheetByName("Sheet1") || ss.getSheetByName("Hoja 1");
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("ADMIN_PASS")) p.setProperty("ADMIN_PASS", "CAMBIAR");
  if (!p.getProperty("EQUIPOS")) p.setProperty("EQUIPOS", JSON.stringify(DEFAULT));
}

function doGet() { return salida({ ok: true, msg: "La calificadora activa" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "load": out = { ok: true, equipos: equipos(), votos: leer().map((v) => ({ id: v.id, puntajes: v.puntajes })) }; break;
      case "save": out = guardar(d); break;
      case "admin": admin(d); out = { ok: true, equipos: equipos(), votos: leer() }; break;
      case "deleteOne": admin(d); borrar(function (v) { return v.id === String(d.id); }); out = { ok: true }; break;
      case "reset": admin(d); borrar(function () { return true; }); out = { ok: true }; break;
      case "saveTeams": admin(d); guardarEquipos(d.nombres); out = { ok: true, equipos: equipos() }; break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function admin(d) {
  if (String(d.pass || "") !== PropertiesService.getScriptProperties().getProperty("ADMIN_PASS")) {
    Utilities.sleep(1500); // frena a quien intente adivinar
    throw new Error("Contraseña incorrecta.");
  }
}

function limpio(s, n) { return String(s || "").trim().replace(/\s+/g, " ").slice(0, n || 40); }

function equipos() {
  try {
    const v = JSON.parse(PropertiesService.getScriptProperties().getProperty("EQUIPOS") || "null");
    if (v && v.length === N_EQUIPOS) return v;
  } catch (e) {}
  return DEFAULT.slice();
}

function guardarEquipos(nombres) {
  if (!nombres || nombres.length !== N_EQUIPOS) throw new Error("Faltan nombres.");
  const v = nombres.map(function (n, i) { return limpio(n, 28) || DEFAULT[i]; });
  PropertiesService.getScriptProperties().setProperty("EQUIPOS", JSON.stringify(v));
}

function leer() {
  const sh = hoja();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      let p = {};
      try { p = JSON.parse(r[3] || "{}"); } catch (e) {}
      return { id: String(r[0]), nombre: String(r[2]), puntajes: p };
    });
}

/* Solo deja pasar e0..e3 con criterios conocidos y números enteros del 1 al 10. */
function validar(p) {
  const out = {};
  for (let i = 0; i < N_EQUIPOS; i++) {
    const src = p && p["e" + i];
    if (!src) continue;
    const t = {};
    CRIT.forEach(function (k) {
      const v = Number(src[k]);
      if (Number.isInteger(v) && v >= 1 && v <= 10) t[k] = v;
    });
    if (Object.keys(t).length) out["e" + i] = t;
  }
  return out;
}

function guardar(d) {
  const id = limpio(d.id, 40);
  if (!/^[a-z0-9-]{8,40}$/i.test(id)) throw new Error("Identificador inválido.");
  const fila = [id, new Date().toISOString(), limpio(d.nombre) || "anónimo", JSON.stringify(validar(d.puntajes))];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    const i = ids.indexOf(id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function borrar(cond) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja();
    if (sh.getLastRow() < 2) return;
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let r = ids.length - 1; r >= 0; r--) {
      if (ids[r][0] && cond({ id: String(ids[r][0]) })) sh.deleteRow(r + 2);
    }
  } finally {
    lock.releaseLock();
  }
}
