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
    if (p.action === "base" || p.action === "votacion") {
      try { cerrarVotaciones(); } catch (err) { console.error(err); }
    }
    if (p.action === "base") return salida(leerBase());
    if (p.action === "add") return salida(agregarPartido(JSON.parse(p.p)));
    if (p.action === "votacion") return salida({ ok: true, votaciones: votacionesAbiertas() });
    if (p.action === "votar") return salida(votar(JSON.parse(p.p)));
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
    partidos: leerPartidos(),
    votaciones: (function () { try { return votacionesAbiertas(); } catch (err) { return []; } })()
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

// ---------- Votación de puntajes ----------
// El día siguiente al partido se abre una votación de VOTO_DIAS días. Cada jugador del partido
// elige su nombre y puntúa a los 10 jugadores de 1 a 5. Cada uno vota una sola vez.
// ANONIMATO: no se guarda qué votó cada uno. Mientras la votación está abierta solo se guardan,
// en las propiedades del script (no en la hoja), la suma y la cantidad de votos de cada jugador
// y la lista de quiénes ya votaron. Al cerrar se escriben en Partidos los promedios
// (puntajes_a / puntajes_b) y el MVP (promedio más alto; si hay empate, todos los empatados),
// y esos datos se borran. Si nadie votó, la fila queda igual.
// La votación se cierra sola la próxima vez que alguien abre la app después del plazo.
// Opcional: Activadores > Agregar activador > cerrarVotaciones, cada hora, para cerrarla sin esperar.
const VOTO_DIAS = 2;

function hoyISO() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function sumarDias(iso, n) {
  const d = new Date(iso + "T12:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

// Pestaña Partidos con sus columnas (por encabezado)
function tablaPartidos() {
  const h = hoja("Partidos");
  const v = h.getDataRange().getValues();
  const enc = (v[0] || []).map(function (c) { return String(c).trim().toLowerCase(); });
  const C = { f: enc.indexOf("fecha"), a: enc.indexOf("equipo_a"), b: enc.indexOf("equipo_b"),
    mvp: enc.indexOf("mvp"), pa: enc.indexOf("puntajes_a"), pb: enc.indexOf("puntajes_b") };
  ["fecha", "equipo_a", "equipo_b", "mvp", "puntajes_a", "puntajes_b"].forEach(function (n) {
    if (enc.indexOf(n) < 0) throw new Error('Falta la columna "' + n + '" en la pestaña Partidos');
  });
  return { h: h, v: v, C: C };
}

// Totales anónimos de una votación: { s: {jugador: suma}, c: {jugador: cantidad}, v: [quiénes votaron] }
function claveVoto(f) { return "voto_" + f; }
function leerTotales(f) {
  const t = PropertiesService.getScriptProperties().getProperty(claveVoto(f));
  return t ? JSON.parse(t) : { s: {}, c: {}, v: [] };
}

// Estado de la fila i de Partidos: "abierta", "vencida" o "" (ya tiene puntajes o es futura)
function estadoFila(T, i, hoy) {
  const r = T.v[i];
  if (!r[T.C.f]) return "";
  if (String(r[T.C.pa]).trim() || String(r[T.C.pb]).trim()) return "";
  const f = fechaISO(r[T.C.f]);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) return "";
  if (hoy < sumarDias(f, 1)) return "";
  if (hoy <= sumarDias(f, VOTO_DIAS)) return "abierta";
  return "vencida";
}

function votacionesAbiertas() {
  const T = tablaPartidos(), hoy = hoyISO(), out = [];
  for (let i = 1; i < T.v.length; i++) {
    if (estadoFila(T, i, hoy) !== "abierta") continue;
    const f = fechaISO(T.v[i][T.C.f]);
    const a = lista(T.v[i][T.C.a]), b = lista(T.v[i][T.C.b]);
    if (!a.length || !b.length) continue;
    out.push({ f: f, a: a, b: b, votaron: leerTotales(f).v, cierra: sumarDias(f, VOTO_DIAS) });
  }
  return out;
}

// o = { f: "2026-11-20", v: "Juan", n: { "Juan": 4, "Pelu": 3, ... } }
function votar(o) {
  if (!o || !/^\d{4}-\d{2}-\d{2}$/.test(o.f) || !o.v || !o.n) throw new Error("Datos del voto inválidos");
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const T = tablaPartidos(), hoy = hoyISO();
    let fila = -1;
    for (let i = 1; i < T.v.length; i++) {
      if (T.v[i][T.C.f] && fechaISO(T.v[i][T.C.f]) === o.f) { fila = i; break; }
    }
    if (fila < 0) throw new Error("No encontré el partido de esa fecha");
    if (estadoFila(T, fila, hoy) !== "abierta") throw new Error("La votación de ese partido está cerrada");
    const jug = lista(T.v[fila][T.C.a]).concat(lista(T.v[fila][T.C.b]));
    const votante = String(o.v).trim();
    if (jug.indexOf(votante) < 0) throw new Error(votante + " no jugó ese partido");
    const notas = jug.map(function (p) {
      const x = Number(o.n[p]);
      if (!(x >= 1 && x <= 5 && Math.round(x) === x)) throw new Error("Falta el puntaje de " + p + " (de 1 a 5)");
      return x;
    });
    const t = leerTotales(o.f);
    if (t.v.indexOf(votante) >= 0) throw new Error(votante + " ya votó en esta votación");
    jug.forEach(function (p, k) { t.s[p] = (t.s[p] || 0) + notas[k]; t.c[p] = (t.c[p] || 0) + 1; });
    t.v.push(votante);
    PropertiesService.getScriptProperties().setProperty(claveVoto(o.f), JSON.stringify(t));
    return { ok: true, votaron: t.v };
  } finally {
    lock.releaseLock();
  }
}

// Cierra las votaciones vencidas: escribe promedios y MVP en Partidos y borra los totales
function cerrarVotaciones() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const T = tablaPartidos(), hoy = hoyISO();
    for (let i = 1; i < T.v.length; i++) {
      if (estadoFila(T, i, hoy) !== "vencida") continue;
      const f = fechaISO(T.v[i][T.C.f]);
      if (!props.getProperty(claveVoto(f))) continue;
      const t = leerTotales(f);
      if (!t.v.length) { props.deleteProperty(claveVoto(f)); continue; }
      const a = lista(T.v[i][T.C.a]), b = lista(T.v[i][T.C.b]), prom = {};
      a.concat(b).forEach(function (p) { prom[p] = t.c[p] ? t.s[p] / t.c[p] : null; });
      const txt = function (eq) {
        return eq.map(function (p) { return prom[p] === null ? "-" : String(Math.round(prom[p] * 100) / 100); }).join(" / ");
      };
      // formato texto para que la hoja no convierta "3 / 4" en una fecha
      T.h.getRange(i + 1, T.C.pa + 1).setNumberFormat("@").setValue(txt(a));
      T.h.getRange(i + 1, T.C.pb + 1).setNumberFormat("@").setValue(txt(b));
      if (!String(T.v[i][T.C.mvp]).trim()) {
        const vals = Object.keys(prom).filter(function (p) { return prom[p] !== null; });
        const mx = Math.max.apply(null, vals.map(function (p) { return prom[p]; }));
        T.h.getRange(i + 1, T.C.mvp + 1).setValue(vals.filter(function (p) { return Math.abs(prom[p] - mx) < 1e-9; }).join(" / "));
      }
      props.deleteProperty(claveVoto(f));
    }
  } finally {
    lock.releaseLock();
  }
}
