/**
 * Varadero App - conexión con Google Sheets
 * La hoja necesita 3 pestañas con estos nombres exactos: Jugadores, Partidos y Ajustes.
 */

// Cambiá esta clave por una propia (la misma va en config.json)
const TOKEN = "CAMBIAR-POR-UNA-CLAVE";

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return salida({ ok: false, error: "Clave incorrecta" });
  try {
    if (p.action === "base") return salida(leerBase());
    if (p.action === "add") return salida(agregarPartido(JSON.parse(p.p)));
    return salida({ ok: false, error: "Acción desconocida" });
  } catch (err) {
    return salida({ ok: false, error: String((err && err.message) || err) });
  }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function hoja(nombre) {
  const h = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!h) throw new Error('Falta la pestaña "' + nombre + '"');
  return h;
}

function leerBase() {
  const v = hoja("Jugadores").getDataRange().getValues();
  if (v.length < 2) throw new Error("La pestaña Jugadores está vacía");
  const cols = v[0].slice(1).map(function (c) { return String(c).trim(); }).filter(Boolean);
  const players = [];
  for (let i = 1; i < v.length; i++) {
    const nombre = String(v[i][0]).trim();
    if (!nombre) continue;
    const vals = cols.map(function (_, j) { return v[i][j + 1]; });
    if (vals.some(function (x) { return x === "" || isNaN(Number(x)); })) {
      throw new Error("Fila " + (i + 1) + " (" + nombre + "): falta un valor o no es un número");
    }
    players.push({ name: nombre, v: vals.map(Number) });
  }
  const aj = {};
  const hAj = SpreadsheetApp.getActive().getSheetByName("Ajustes");
  if (hAj) {
    hAj.getDataRange().getValues().forEach(function (r) {
      if (r[0]) aj[String(r[0]).trim().toLowerCase()] = r[1];
    });
  }
  const bonus = aj["ventaja"] === undefined || aj["ventaja"] === "" ? 0.5 : Number(aj["ventaja"]);
  return {
    ok: true,
    cols: cols,
    players: players,
    fis: aj["columna_fisico"] ? String(aj["columna_fisico"]).trim() : null,
    bonus: isNaN(bonus) ? 0.5 : bonus,
    partidos: leerPartidos()
  };
}

// Historial para las estadísticas. Las columnas se buscan por su encabezado:
// fecha | equipo_a | equipo_b | goles_a | goles_b | arqueros | registrado | mvp | puntajes_a | puntajes_b
function leerPartidos() {
  const h = SpreadsheetApp.getActive().getSheetByName("Partidos");
  if (!h) return [];
  const v = h.getDataRange().getValues();
  if (v.length < 2) return [];
  const enc = v[0].map(function (c) { return String(c).trim().toLowerCase(); });
  const col = function (n, def) { const i = enc.indexOf(n); return i >= 0 ? i : def; };
  const C = { f: col("fecha", 0), a: col("equipo_a", 1), b: col("equipo_b", 2), ga: col("goles_a", 3), gb: col("goles_b", 4),
    g: col("arqueros", 5), mvp: col("mvp", -1), pa: col("puntajes_a", -1), pb: col("puntajes_b", -1) };
  const cel = function (r, i) { return i >= 0 && i < r.length ? r[i] : ""; };
  const num = function (x) {
    if (x === "" || x === null) return null;
    const n = Number(String(x).replace(",", "."));
    return isNaN(n) ? null : n;
  };
  const notas = function (x, n) {
    const l = String(x || "").split(/\s*\/\s*/).map(function (t) { return num(t.trim()); });
    if (!String(x || "").trim() || l.length !== n) return null;
    return l;
  };
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[C.f]) continue;
    const a = lista(cel(r, C.a)), b = lista(cel(r, C.b));
    if (!a.length || !b.length) continue;
    let ga = num(cel(r, C.ga)), gb = num(cel(r, C.gb));
    if (ga === null || gb === null) ga = gb = null;
    out.push({
      f: fechaISO(r[C.f]), a: a, b: b, ga: ga, gb: gb,
      g: lista(cel(r, C.g) || ""), mvp: lista(cel(r, C.mvp) || ""),
      pa: notas(cel(r, C.pa), a.length), pb: notas(cel(r, C.pb), b.length)
    });
  }
  return out;
}

function fechaISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  return s;
}
function lista(s) {
  return String(s).split(/\s*[\/|]\s*/).map(function (x) { return x.trim(); }).filter(Boolean);
}
function clave(f, a, b) {
  return f + "|" + a.slice().sort().join(",") + "|" + b.slice().sort().join(",");
}

// Columnas de la pestaña Partidos: fecha | equipo_a | equipo_b | goles_a | goles_b | arqueros | registrado
function agregarPartido(o) {
  if (!o || !/^\d{4}-\d{2}-\d{2}$/.test(o.f) || !o.a || !o.b || !o.a.length || !o.b.length) {
    throw new Error("Datos del partido inválidos");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const h = hoja("Partidos");
    const v = h.getDataRange().getValues();
    const k = clave(o.f, o.a, o.b);
    for (let i = 1; i < v.length; i++) {
      if (!v[i][0]) continue;
      if (clave(fechaISO(v[i][0]), lista(v[i][1]), lista(v[i][2])) === k) return { ok: true, dup: true };
    }
    const fila = Math.max(h.getLastRow(), 1) + 1;
    h.getRange(fila, 1).setNumberFormat("@").setValue(o.f);
    h.getRange(fila, 2, 1, 6).setValues([[o.a.join(" / "), o.b.join(" / "), "", "", (o.g || []).join(" / "), new Date()]]);
    return { ok: true, dup: false };
  } finally {
    lock.releaseLock();
  }
}
