/**
 * Planta de Emergencia PLANELEC 500 kW — CGR-PEM-001
 * Ficha técnica pública + bitácora de mantenimiento con modo técnico.
 * Google Apps Script vinculado a una Hoja de cálculo de Google.
 *
 * QR público:  <URL_DE_LA_APP>
 * QR técnico:  <URL_DE_LA_APP>?modo=tecnico
 *
 * La clave se valida AQUÍ, en el servidor de Google. No viaja al celular.
 */

const EQUIPO_ID   = 'CGR-PEM-001';
const PIN_INICIAL = '737070';   // Solo se usa la primera vez que corres setup()
const MAX_FALLOS  = 15;          // Intentos fallidos antes de bloquear
const BLOQUEO_SEG = 900;         // 15 minutos de bloqueo

const COLS = ['ID','Fecha','Tipo','Técnico','Horómetro (h)','Actividades',
  'Combustible %','Aceite bar','Temp °C','Baterías V','RPM','Hz','kWh',
  'V L-L','V L-N','Refacciones','Observaciones','Registrado'];

/* ───────── Página ───────── */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.modoTecnico = !!(e && e.parameter && e.parameter.modo === 'tecnico');
  return t.evaluate()
    .setTitle('Planta de Emergencia ' + EQUIPO_ID)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ───────── Lectura pública ───────── */
function getData() {
  const ss = SpreadsheetApp.getActive();
  const eq = ss.getSheetByName('Equipo').getDataRange().getValues().slice(1);
  const info = {};
  eq.forEach(r => info[String(r[0])] = txt_(r[1]));

  const ficha = ss.getSheetByName('Ficha').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() !== '')
    .map(r => [txt_(r[0]), txt_(r[1])]);

  const historial = ss.getSheetByName('Historial').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() !== '')
    .map(fila_)
    .sort((a, b) => a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0);

  return { id: EQUIPO_ID, info, ficha, historial };
}

/* ───────── Modo técnico (requiere clave) ───────── */
function login(pin) { return { ok: checkPin_(pin) }; }

function guardarRegistro(pin, r) {
  if (!checkPin_(pin)) return { ok: false, error: 'Clave incorrecta.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.fecha || '')) return { ok: false, error: 'Fecha inválida.' };
  if (!String(r.tecnico || '').trim()) return { ok: false, error: 'Falta el nombre del técnico.' };

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Historial');
    const l = r.lectura || {};
    const fila = [
      r.id || ('R' + Date.now().toString(36).toUpperCase()),
      r.fecha, limpio_(r.tipo, 40), limpio_(r.tecnico, 80), num_(r.horometro),
      (r.actividades || []).map(a => limpio_(a, 120)).join('\n'),
      num_(l.comb), num_(l.aceite), num_(l.temp), num_(l.bat), num_(l.rpm), num_(l.hz), num_(l.kwh),
      limpio_(l.vll, 40), limpio_(l.vln, 40),
      limpio_(r.refacciones, 400), limpio_(r.obs, 2000), new Date()
    ];
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(x => String(x[0]));
    const i = r.id ? ids.indexOf(String(r.id)) : -1;
    if (i > 0) sh.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function eliminarRegistro(pin, id) {
  if (!checkPin_(pin)) return { ok: false, error: 'Clave incorrecta.' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Historial');
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(x => String(x[0]));
    const i = ids.indexOf(String(id));
    if (i > 0) sh.deleteRow(i + 1);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function guardarFicha(pin, filas) {
  if (!checkPin_(pin)) return { ok: false, error: 'Clave incorrecta.' };
  const datos = (filas || []).filter(f => f[0] || f[1]).slice(0, 60)
    .map(f => [limpio_(f[0], 60), limpio_(f[1], 200)]);
  const sh = SpreadsheetApp.getActive().getSheetByName('Ficha');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  if (datos.length) sh.getRange(2, 1, datos.length, 2).setValues(datos);
  return { ok: true };
}

/* ───────── Seguridad de la clave ───────── */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('fallos') || 0);
  if (fallos >= MAX_FALLOS) { Utilities.sleep(1500); return false; }
  const real = PropertiesService.getScriptProperties().getProperty('PIN');
  const ok = real && String(pin || '').trim() === real;
  if (!ok) { cache.put('fallos', String(fallos + 1), BLOQUEO_SEG); Utilities.sleep(800); }
  return !!ok;
}

/** Para cambiar la clave: escribe la nueva aquí, guarda y ejecuta esta función una vez. */
function cambiarClave() {
  const NUEVA_CLAVE = '737070';
  PropertiesService.getScriptProperties().setProperty('PIN', NUEVA_CLAVE);
  CacheService.getScriptCache().remove('fallos');
}

/* ───────── Utilidades ───────── */
function fila_(r) {
  return {
    id: String(r[0]), fecha: fecha_(r[1]), tipo: txt_(r[2]), tecnico: txt_(r[3]), horometro: txt_(r[4]),
    actividades: txt_(r[5]) ? txt_(r[5]).split('\n') : [],
    lectura: { comb: txt_(r[6]), aceite: txt_(r[7]), temp: txt_(r[8]), bat: txt_(r[9]), rpm: txt_(r[10]),
               hz: txt_(r[11]), kwh: txt_(r[12]), vll: txt_(r[13]), vln: txt_(r[14]) },
    refacciones: txt_(r[15]), obs: txt_(r[16])
  };
}
function fecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
function txt_(v) { return v === null || v === undefined ? '' : String(v); }
function num_(v) { return v === null || v === undefined || v === '' || isNaN(Number(v)) ? '' : Number(v); }
function limpio_(v, max) { return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, max); }

/* ───────── Instalación (ejecutar UNA vez) ───────── */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const hoja = (n) => ss.getSheetByName(n) || ss.insertSheet(n);

  const eq = hoja('Equipo');
  if (eq.getLastRow() === 0) eq.getRange(1, 1, 5, 2).setValues([
    ['Campo', 'Valor'],
    ['nombre', 'Planta de Emergencia PLANELEC 500 kW'],
    ['ubicacion', 'Planta 1 · Muebles Liz'],
    ['area', 'Mantenimiento Industrial'],
    ['subtitulo', 'Equipo de respaldo']
  ]);

  const fi = hoja('Ficha');
  if (fi.getLastRow() === 0) {
    const ficha = [
      ['Campo', 'Valor'],
      ['Fabricante', 'PLANELEC · Plantas Eléctricas México, S.A. de C.V.'],
      ['Modelo', '1S069D0500B1600'],
      ['Tablero de control', 'PEM-CAP'],
      ['Motor', '69X18-69'],
      ['No. serie motor', 'B0393007'],
      ['Generador', '54L1D-641'],
      ['No. serie generador', 'X23B082603'],
      ['Potencia emergencia', '500 kW / 625 kVA'],
      ['Potencia continua', '452 kW / 565 kVA'],
      ['Voltaje', '220 / 127 V'],
      ['Fases / hilos', '3 fases · 4 hilos'],
      ['Frecuencia', '60 Hz'],
      ['Velocidad', '1800 RPM'],
      ['Factor de potencia', '0.8'],
      ['Voltaje de baterías', '24 VCD'],
      ['Año de fabricación', 'Agosto 2023'],
      ['Peso', '5524 kg']
    ];
    fi.getRange(1, 1, ficha.length, 2).setValues(ficha);
  }

  const h = hoja('Historial');
  if (h.getLastRow() === 0) {
    h.appendRow(COLS);
    h.getRange('B:B').setNumberFormat('@');
    h.appendRow(['R1', '2026-03-20', 'Lectura semanal', 'Lectura inicial (fotografía del panel)', 103.25, '',
      55, 2.72, 62, 28.15, 1798, 60, 2603, '219 / 221 / 221', '127 / 127 / 128', '',
      'Datos capturados de fotografías del panel PEM-CAP.', new Date()]);
  }
  [eq, fi, h].forEach(s => { s.setFrozenRows(1); s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold'); });

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN')) props.setProperty('PIN', PIN_INICIAL);

  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}
