/**
 * Backend del módulo de Auditoría (Inspecciones AAA).
 *
 * Instalación (una sola vez):
 * 1. Crea un Google Sheet nuevo (será la base de datos) → Extensiones → Apps Script → pega este archivo.
 * 2. Configuración del proyecto → Propiedades de la secuencia de comandos → agrega:
 *      ADMIN_KEY = clave del administrador (tú)
 *      SUP_KEY   = clave que compartes con los supervisores
 * 3. Implementar → Nueva implementación → Tipo "Aplicación web"
 *      Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
 * 4. Copia la URL que termina en /exec y pégala en API_URL dentro de auditoria/index.html.
 * Las pestañas ORDENES y AUDITORIAS se crean solas. Las fotos quedan en tu Drive (carpeta "Auditoria AAA Fotos"), privadas.
 */
var ORDEN_COLS = ['Orden', 'Supervisor', 'Dirección', 'Tipo de trabajo', 'Inspector', 'Observación', 'Estado', 'Cargada el', 'Detalle'];
var AUD_COLS = ['ID', 'Orden', 'Supervisor', 'Fecha', 'Datos', 'Latitud', 'Longitud', 'Fotos'];

function doGet() { return out_({ ok: true, msg: 'API Auditoría AAA activa' }); }

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var role = role_(req.key);
    if (!role) return out_({ ok: false, error: 'Clave incorrecta' });
    var a = req.action;
    if (a === 'login') return out_({ ok: true, role: role, supervisores: supervisores_() });
    if (a === 'ordenes') return out_({ ok: true, ordenes: ordenes_(role === 'admin' ? null : req.supervisor), auditorias: auditorias_(role === 'admin' ? null : req.supervisor) });
    if (role !== 'admin' && a === 'cargar') return out_({ ok: false, error: 'Solo administrador' });
    if (a === 'cargar') return out_(cargar_(req));
    if (a === 'guardar') return out_(guardar_(req));
    return out_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function role_(key) {
  var p = PropertiesService.getScriptProperties();
  if (key && key === p.getProperty('ADMIN_KEY')) return 'admin';
  if (key && key === p.getProperty('SUP_KEY')) return 'supervisor';
  return null;
}

function sheet_(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols.length).setFontWeight('bold'); }
  return sh;
}

function rows_(sh, cols) {
  var v = sh.getDataRange().getValues();
  var res = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var j = 0; j < cols.length; j++) o[cols[j]] = v[i][j] instanceof Date ? v[i][j].toISOString() : v[i][j];
    res.push(o);
  }
  return res;
}

function supervisores_() {
  var seen = {}, list = [];
  rows_(sheet_('ORDENES', ORDEN_COLS), ORDEN_COLS).forEach(function (r) {
    var s = String(r['Supervisor']).trim();
    if (s && !seen[s]) { seen[s] = 1; list.push(s); }
  });
  return list.sort();
}

function ordenes_(sup) {
  return rows_(sheet_('ORDENES', ORDEN_COLS), ORDEN_COLS).filter(function (r) {
    return r['Orden'] !== '' && (!sup || String(r['Supervisor']).trim() === sup);
  });
}

function auditorias_(sup) {
  return rows_(sheet_('AUDITORIAS', AUD_COLS), AUD_COLS).filter(function (r) { return !sup || String(r['Supervisor']).trim() === sup; });
}

// req.rows: [{Orden, Supervisor, Dirección, ...otras columnas}]; req.reemplazar borra las órdenes actuales
function cargar_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sheet_('ORDENES', ORDEN_COLS);
    if (req.reemplazar && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ORDEN_COLS.length).clearContent();
    var existentes = {};
    if (!req.reemplazar) rows_(sh, ORDEN_COLS).forEach(function (r) { existentes[String(r['Orden'])] = 1; });
    var now = new Date(), nuevas = [], omitidas = 0;
    (req.rows || []).forEach(function (r) {
      var id = String(r['Orden'] || '').trim();
      if (!id || existentes[id]) { omitidas++; return; }
      existentes[id] = 1;
      var extra = {};
      Object.keys(r).forEach(function (k) { if (ORDEN_COLS.indexOf(k) < 0) extra[k] = r[k]; });
      nuevas.push([id, String(r['Supervisor'] || '').trim(), r['Dirección'] || '', r['Tipo de trabajo'] || '', r['Inspector'] || '', r['Observación'] || '', 'Pendiente', now, JSON.stringify(extra)]);
    });
    if (nuevas.length) {
      var rng = sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, ORDEN_COLS.length);
      sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 1).setNumberFormat('@'); // Orden como texto (evita notación científica)
      rng.setValues(nuevas);
    }
    return { ok: true, cargadas: nuevas.length, omitidas: omitidas };
  } finally { lock.releaseLock(); }
}

function folder_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Auditoria AAA Fotos');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

// req: {id, orden, supervisor, datos:{...}, lat, lng, fotos:[{name,data(base64 jpeg)}]}
function guardar_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var aud = sheet_('AUDITORIAS', AUD_COLS);
    var ya = rows_(aud, AUD_COLS).filter(function (r) { return r['ID'] === req.id; });
    if (ya.length) return { ok: true, duplicado: true };   // reintento de un envío ya recibido
    var urls = [];
    if (req.fotos && req.fotos.length) {
      var root = folder_(), it = root.getFoldersByName(String(req.orden));
      var sub = it.hasNext() ? it.next() : root.createFolder(String(req.orden));
      req.fotos.forEach(function (f, i) {
        var blob = Utilities.newBlob(Utilities.base64Decode(f.data), 'image/jpeg', req.orden + '_' + new Date().getTime() + '_' + (i + 1) + '.jpg');
        urls.push(sub.createFile(blob).getUrl());
      });
    }
    var fila = aud.getLastRow() + 1;
    aud.getRange(fila, 2, 1, 1).setNumberFormat('@');
    aud.getRange(fila, 1, 1, AUD_COLS.length).setValues([[req.id, String(req.orden), req.supervisor, new Date(), JSON.stringify(req.datos || {}), req.lat || '', req.lng || '', urls.join('\n')]]);
    var ord = sheet_('ORDENES', ORDEN_COLS);
    rows_(ord, ORDEN_COLS).forEach(function (r) {
      if (String(r['Orden']) === String(req.orden)) ord.getRange(r._row, 7).setValue('Auditada');
    });
    return { ok: true, fotos: urls.length };
  } finally { lock.releaseLock(); }
}
