/**
 * Backend del módulo de Auditoría (Inspecciones AAA).
 *
 * Instalación (una sola vez):
 * 1. Crea un Google Sheet nuevo (será la base de datos) → Extensiones → Apps Script → pega este archivo.
 * 2. Configuración del proyecto → Propiedades de la secuencia de comandos → agrega:
 *      ADMIN_KEY = clave del administrador (tú)
 *    Cada supervisor tiene su propia clave: se crean desde el panel de administrador (pestaña Supervisores).
 * 3. Implementar → Nueva implementación → Tipo "Aplicación web"
 *      Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
 * 4. Copia la URL que termina en /exec y pégala en API_URL dentro de auditoria/index.html.
 * Las pestañas ORDENES y AUDITORIAS se crean solas. Las fotos quedan en tu Drive (carpeta "Auditoria AAA Fotos"), privadas.
 */
var ORDEN_COLS = ['Orden', 'Supervisor', 'Dirección', 'Tipo de trabajo', 'Inspector', 'Observación', 'Estado', 'Cargada el', 'Detalle'];
var AUD_COLS = ['ID', 'Orden', 'Supervisor', 'Fecha', 'Datos', 'Latitud', 'Longitud', 'Fotos', 'Dirección', 'Tipo de trabajo', 'Inspector', 'Detalle'];
var SUP_COLS = ['Supervisor', 'Clave'];

function doGet() { return out_({ ok: true, msg: 'API Auditoría AAA activa' }); }

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var who = auth_(req.key);
    if (!who) return out_({ ok: false, error: 'Clave incorrecta' });
    var a = req.action, admin = who.role === 'admin';
    if (a === 'login') return out_({ ok: true, role: who.role, supervisor: who.supervisor || null });
    if (a === 'ordenes') {
      var sup = admin ? null : who.supervisor;
      var r = { ok: true, ordenes: ordenes_(sup), auditorias: auditorias_(sup) };
      if (admin) r.supervisores = supervisoresAdmin_();
      return out_(r);
    }
    if (a === 'guardar') return out_(guardar_(req, who));
    if (!admin) return out_({ ok: false, error: 'Solo administrador' });
    if (a === 'cargar') return out_(cargar_(req));
    if (a === 'eliminarAuditadas') return out_(eliminarAuditadas_());
    if (a === 'guardarSupervisor') return out_(guardarSupervisor_(req));
    if (a === 'eliminarSupervisor') return out_(eliminarSupervisor_(req));
    return out_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Devuelve {role:'admin'} o {role:'supervisor', supervisor:nombre} según la clave
function auth_(key) {
  key = String(key || '').trim();
  if (!key) return null;
  if (key === PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) return { role: 'admin' };
  var m = rows_(sheet_('SUPERVISORES', SUP_COLS), SUP_COLS).filter(function (r) { return String(r['Clave']).trim() === key; });
  return m.length ? { role: 'supervisor', supervisor: String(m[0]['Supervisor']).trim() } : null;
}

function sheet_(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols.length).setFontWeight('bold'); }
  if (sh.getLastColumn() < cols.length) { sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold'); }
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

// Supervisores con clave + nombres que aparecen en órdenes y aún no tienen clave
function supervisoresAdmin_() {
  var con = rows_(sheet_('SUPERVISORES', SUP_COLS), SUP_COLS).map(function (r) { return { Supervisor: String(r['Supervisor']).trim(), Clave: String(r['Clave']).trim() }; });
  var have = {}; con.forEach(function (c) { have[c.Supervisor] = 1; });
  ordenes_(null).forEach(function (r) {
    var n = String(r['Supervisor']).trim();
    if (n && !have[n]) { have[n] = 1; con.push({ Supervisor: n, Clave: '' }); }
  });
  return con.sort(function (x, y) { return x.Supervisor < y.Supervisor ? -1 : 1; });
}

function guardarSupervisor_(req) {
  var nombre = String(req.nombre || '').trim(), clave = String(req.clave || '').trim();
  if (!nombre || clave.length < 4) return { ok: false, error: 'Nombre y clave (mínimo 4 caracteres) son obligatorios' };
  if (clave === PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) return { ok: false, error: 'Esa clave no está disponible' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sheet_('SUPERVISORES', SUP_COLS), rs = rows_(sh, SUP_COLS), fila = 0;
    for (var i = 0; i < rs.length; i++) {
      if (String(rs[i]['Supervisor']).trim() === nombre) fila = rs[i]._row;
      else if (String(rs[i]['Clave']).trim() === clave) return { ok: false, error: 'Esa clave ya la usa otro supervisor' };
    }
    if (!fila) fila = sh.getLastRow() + 1;
    sh.getRange(fila, 1, 1, 2).setNumberFormat('@').setValues([[nombre, clave]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function eliminarSupervisor_(req) {
  var sh = sheet_('SUPERVISORES', SUP_COLS), nombre = String(req.nombre || '').trim();
  rows_(sh, SUP_COLS).reverse().forEach(function (r) { if (String(r['Supervisor']).trim() === nombre) sh.deleteRow(r._row); });
  return { ok: true };
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

function datosOrden_(orden) {
  var m = ordenes_(null).filter(function (r) { return String(r['Orden']) === String(orden); });
  return m.length ? m[0] : null;
}

// req: {id, orden, datos:{...}, lat, lng, fotos:[{name,data(base64 jpeg)}]}. El supervisor sale de la clave, no del cliente.
function guardar_(req, who) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var o = datosOrden_(req.orden);
    var sup = who.role === 'supervisor' ? who.supervisor : (o ? String(o['Supervisor']).trim() : '');
    if (who.role === 'supervisor' && (!o || String(o['Supervisor']).trim() !== sup)) return { ok: false, error: 'La orden no está asignada a este supervisor' };
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
    escribirAud_(aud, [req.id, String(req.orden), sup, new Date(), JSON.stringify(req.datos || {}), req.lat || '', req.lng || '', urls.join('\n'), o ? o['Dirección'] : '', o ? o['Tipo de trabajo'] : '', o ? o['Inspector'] : '', o ? o['Detalle'] : '']);
    marcarAuditada_(req.orden);
    return { ok: true, fotos: urls.length };
  } finally { lock.releaseLock(); }
}

function escribirAud_(aud, fila) {
  var n = aud.getLastRow() + 1;
  aud.getRange(n, 2, 1, 1).setNumberFormat('@');
  aud.getRange(n, 1, 1, AUD_COLS.length).setValues([fila]);
}

function marcarAuditada_(orden) {
  var ord = sheet_('ORDENES', ORDEN_COLS);
  rows_(ord, ORDEN_COLS).forEach(function (r) {
    if (String(r['Orden']) === String(orden)) ord.getRange(r._row, 7).setValue('Auditada');
  });
}

// Borra de ORDENES las órdenes ya auditadas. El historial queda en AUDITORIAS (con dirección, tipo e inspector).
function eliminarAuditadas_() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sheet_('ORDENES', ORDEN_COLS), n = 0;
    rows_(sh, ORDEN_COLS).reverse().forEach(function (r) { if (r['Estado'] === 'Auditada') { sh.deleteRow(r._row); n++; } });
    return { ok: true, eliminadas: n };
  } finally { lock.releaseLock(); }
}
