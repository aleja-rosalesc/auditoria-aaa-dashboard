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
var AUD_COLS = ['ID', 'Orden', 'Supervisor', 'Fecha', 'Datos', 'Latitud', 'Longitud', 'Fotos', 'Dirección', 'Tipo de trabajo', 'Inspector', 'Detalle', 'Firma', 'Editada el', 'EdicionID'];
var SUP_COLS = ['Supervisor', 'Clave'];
var DISP_COLS = ['Token', 'Rol', 'Supervisor', 'Actualizado'];

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
    if (a === 'registrarPush') return out_(registrarPush_(req, who));
    if (a === 'desregistrarPush') return out_(desregistrarPush_(req));
    if (!admin) return out_({ ok: false, error: 'Solo administrador' });
    if (a === 'cargar') return out_(cargar_(req));
    if (a === 'eliminarAuditadas') return out_(eliminarAuditadas_());
    if (a === 'eliminarOrdenes') return out_(eliminarOrdenes_(req));
    if (a === 'reasignar') return out_(reasignar_(req));
    if (a === 'vaciarTodo') return out_(vaciarTodo_());
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
    var reg = registrados_(), desconocidos = {};
    (req.rows || []).forEach(function (r) { if (!reg[norm_(r['Supervisor'])]) desconocidos[String(r['Supervisor'] || '').trim() || '(vacío)'] = 1; });
    var faltan = Object.keys(desconocidos);
    if (faltan.length) return { ok: false, error: 'No se cargó nada. Estos supervisores no existen: ' + faltan.join(', ') + '. Créalos primero en la pestaña Supervisores o corrige el nombre.' };
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
      nuevas.push([id, reg[norm_(r['Supervisor'])], r['Dirección'] || '', r['Tipo de trabajo'] || '', r['Inspector'] || '', r['Observación'] || '', 'Pendiente', now, JSON.stringify(extra)]);
    });
    if (nuevas.length) {
      var rng = sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, ORDEN_COLS.length);
      sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 1).setNumberFormat('@'); // Orden como texto (evita notación científica)
      rng.setValues(nuevas);
    }
    avisarSupervisores_(nuevas.map(function (n) { return n[1]; }), 'Nuevas órdenes para auditar', function (c) { return 'Tienes ' + c + (c === 1 ? ' orden nueva' : ' órdenes nuevas'); });
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

// req: {id, orden, datos:{...}, lat, lng, fotos:[{name,data(base64 jpeg)}], firma(base64 png, opcional), editar(true para corregir la auditoría existente)}. El supervisor sale de la clave, no del cliente.
function guardar_(req, who) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var o = datosOrden_(req.orden);
    var sup = who.role === 'supervisor' ? who.supervisor : (o ? String(o['Supervisor']).trim() : '');
    if (who.role === 'supervisor' && (!o || String(o['Supervisor']).trim() !== sup)) return { ok: false, error: 'La orden no está asignada a este supervisor' };
    var aud = sheet_('AUDITORIAS', AUD_COLS), previa = null, filas = rows_(aud, AUD_COLS);
    if (req.editar) {
      var propias = filas.filter(function (r) { return String(r['Orden']) === String(req.orden); });
      if (!propias.length) return { ok: false, error: 'No hay una auditoría para editar' };
      previa = propias[propias.length - 1];
      if (who.role === 'supervisor' && String(previa['Supervisor']).trim() !== sup) return { ok: false, error: 'La auditoría no es de este supervisor' };
      if (String(previa['EdicionID']) === String(req.id)) return { ok: true, duplicado: true };   // reintento de una edición ya recibida
    } else if (filas.some(function (r) { return r['ID'] === req.id; })) return { ok: true, duplicado: true };   // reintento de un envío ya recibido
    var urls = [], firmaUrl = '';
    if ((req.fotos && req.fotos.length) || req.firma) {
      var root = folder_(), it = root.getFoldersByName(String(req.orden));
      var sub = it.hasNext() ? it.next() : root.createFolder(String(req.orden));
      if (req.firma) firmaUrl = sub.createFile(Utilities.newBlob(Utilities.base64Decode(req.firma), 'image/png', 'firma_' + req.orden + '_' + new Date().getTime() + '.png')).getUrl();
    }
    if (req.fotos && req.fotos.length) {
      req.fotos.forEach(function (f, i) {
        var blob = Utilities.newBlob(Utilities.base64Decode(f.data), 'image/jpeg', req.orden + '_' + new Date().getTime() + '_' + (i + 1) + '.jpg');
        urls.push(sub.createFile(blob).getUrl());
      });
    }
    if (previa) {
      var col = function (n) { return AUD_COLS.indexOf(n) + 1; }, fila = previa._row;
      aud.getRange(fila, col('Datos')).setValue(JSON.stringify(req.datos || {}));
      if (req.lat) { aud.getRange(fila, col('Latitud')).setValue(req.lat); aud.getRange(fila, col('Longitud')).setValue(req.lng); }
      if (urls.length) aud.getRange(fila, col('Fotos')).setValue((previa['Fotos'] ? previa['Fotos'] + '\n' : '') + urls.join('\n'));
      if (firmaUrl) aud.getRange(fila, col('Firma')).setValue(firmaUrl);
      aud.getRange(fila, col('Editada el')).setValue(new Date());
      aud.getRange(fila, col('EdicionID')).setValue(req.id);
      avisarAdmin_('Auditoría editada', sup + ' · Orden ' + req.orden + ' · ' + ((req.datos && req.datos.resultado) || 'Auditoría'));
      return { ok: true, editada: true, fotos: urls.length };
    }
    escribirAud_(aud, [req.id, String(req.orden), sup, new Date(), JSON.stringify(req.datos || {}), req.lat || '', req.lng || '', urls.join('\n'), o ? o['Dirección'] : '', o ? o['Tipo de trabajo'] : '', o ? o['Inspector'] : '', o ? o['Detalle'] : '', firmaUrl, '', '']);
    marcarAuditada_(req.orden);
    avisarAdmin_('Nueva auditoría', sup + ' · Orden ' + req.orden + ' · ' + ((req.datos && req.datos.resultado) || 'Auditoría'));
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

// Quita de la hoja las filas cuya columna `col` esté en `set`; devuelve cuántas quitó
function borrarFilas_(sh, cols, col, set) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, 1, last - 1, cols.length), v = rng.getValues(), idx = cols.indexOf(col), keep = [];
  v.forEach(function (row) { if (!set[String(row[idx])]) keep.push(row); });
  var n = v.length - keep.length;
  if (!n) return 0;
  rng.clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, cols.length).setValues(keep);
  return n;
}

// req.ordenes: lista de números de orden. Borra la orden, sus auditorías y manda a la papelera de Drive sus fotos.
function eliminarOrdenes_(req) {
  var lista = req.ordenes || [], set = {};
  lista.forEach(function (o) { set[String(o)] = 1; });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var a = borrarFilas_(sheet_('ORDENES', ORDEN_COLS), ORDEN_COLS, 'Orden', set);
    var b = borrarFilas_(sheet_('AUDITORIAS', AUD_COLS), AUD_COLS, 'Orden', set);
    var root = folder_();
    lista.forEach(function (o) { var it = root.getFoldersByName(String(o)); while (it.hasNext()) it.next().setTrashed(true); });
    return { ok: true, ordenes: a, auditorias: b };
  } finally { lock.releaseLock(); }
}

// Deja ORDENES y AUDITORIAS vacías y manda a la papelera la carpeta de fotos. Los supervisores y sus claves se conservan.
function vaciarTodo_() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var n = 0;
    [['ORDENES', ORDEN_COLS], ['AUDITORIAS', AUD_COLS]].forEach(function (x) {
      var sh = sheet_(x[0], x[1]);
      if (sh.getLastRow() > 1) { n += sh.getLastRow() - 1; sh.getRange(2, 1, sh.getLastRow() - 1, x[1].length).clearContent(); }
    });
    var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
    if (id) { try { DriveApp.getFolderById(id).setTrashed(true); } catch (e) {} p.deleteProperty('FOLDER_ID'); }
    return { ok: true, filas: n };
  } finally { lock.releaseLock(); }
}

// Nombre sin tildes, mayúsculas ni espacios de más, para comparar nombres de supervisores
function norm_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }

// Supervisores registrados en la hoja SUPERVISORES: { nombreNormalizado: nombreOficial }
function registrados_() {
  var m = {};
  rows_(sheet_('SUPERVISORES', SUP_COLS), SUP_COLS).forEach(function (r) { var n = String(r['Supervisor']).trim(); if (n) m[norm_(n)] = n; });
  return m;
}

// Cambia el supervisor de las filas cuya Orden esté en `set`; devuelve cuántas cambió
function cambiarSup_(sh, cols, set, nuevo) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var io = cols.indexOf('Orden'), is = cols.indexOf('Supervisor'), n = last - 1;
  var ord = sh.getRange(2, io + 1, n, 1).getValues(), rng = sh.getRange(2, is + 1, n, 1), sup = rng.getValues(), c = 0;
  for (var i = 0; i < n; i++) if (set[String(ord[i][0])]) { sup[i][0] = nuevo; c++; }
  if (c) rng.setValues(sup);
  return c;
}

// req: {ordenes:[...], supervisor}. El destino debe estar registrado en la pestaña Supervisores.
function reasignar_(req) {
  var destino = registrados_()[norm_(req.supervisor)];
  if (!destino) return { ok: false, error: 'El supervisor "' + req.supervisor + '" no existe. Créalo primero en la pestaña Supervisores.' };
  var set = {}; (req.ordenes || []).forEach(function (o) { set[String(o)] = 1; });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var a = cambiarSup_(sheet_('ORDENES', ORDEN_COLS), ORDEN_COLS, set, destino);
    var b = cambiarSup_(sheet_('AUDITORIAS', AUD_COLS), AUD_COLS, set, destino);
    if (a) avisarSupervisores_(Array.apply(null, Array(a)).map(function () { return destino; }), 'Órdenes reasignadas', function (c) { return 'Te reasignaron ' + c + (c === 1 ? ' orden' : ' órdenes'); });
    return { ok: true, ordenes: a, auditorias: b, supervisor: destino };
  } finally { lock.releaseLock(); }
}

/* ================= Notificaciones push (Firebase Cloud Messaging) =================
 * Configuración (una sola vez):
 *  1. Firebase → Configuración del proyecto → Cuentas de servicio → "Generar nueva clave privada" (descarga un .json).
 *  2. Apps Script → Configuración del proyecto → Propiedades de la secuencia de comandos → agrega
 *       FCM_SERVICE_ACCOUNT = el contenido completo de ese .json   (es secreto: no lo compartas ni lo subas a GitHub)
 *  3. Ejecuta una vez la función autorizarPush desde el editor y acepta los permisos.
 *  4. Implementar → Administrar implementaciones → lápiz → Versión nueva → Implementar.
 */
function autorizarPush() {
  var sa = JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT') || 'null');
  if (!sa) throw new Error('Falta la propiedad FCM_SERVICE_ACCOUNT');
  CacheService.getScriptCache().remove('fcm_at');
  fcmToken_();
  Logger.log('Listo: Firebase responde para el proyecto ' + sa.project_id);
}

// Token de acceso de la cuenta de servicio (JWT firmado con RS256), guardado en caché ~55 min
function fcmToken_() {
  var cache = CacheService.getScriptCache(), guardado = cache.get('fcm_at');
  if (guardado) return guardado;
  var sa = JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT') || 'null');
  if (!sa) return null;
  var ahora = Math.floor(Date.now() / 1000);
  var b64 = function (o) { return Utilities.base64EncodeWebSafe(typeof o === 'string' ? o : JSON.stringify(o)).replace(/=+$/, ''); };
  var cabecera = b64({ alg: 'RS256', typ: 'JWT' });
  var reclamo = b64({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: ahora, exp: ahora + 3600 });
  var firma = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(cabecera + '.' + reclamo, sa.private_key)).replace(/=+$/, '');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', muteHttpExceptions: true, payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: cabecera + '.' + reclamo + '.' + firma } });
  var j = JSON.parse(r.getContentText());
  if (!j.access_token) throw new Error('Firebase no entregó el token: ' + r.getContentText());
  cache.put('fcm_at', j.access_token, 3300);
  return j.access_token;
}

// Envía el mismo aviso a una lista de tokens; borra los que Firebase reporta como inexistentes. Nunca lanza error.
function enviarPush_(tokens, titulo, cuerpo) {
  try {
    if (!tokens || !tokens.length) return 0;
    var at = fcmToken_(); if (!at) return 0;
    var sa = JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT'));
    var url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
    var reqs = tokens.map(function (tk) {
      return { url: url, method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + at },
        payload: JSON.stringify({ message: { token: tk, data: { titulo: String(titulo), cuerpo: String(cuerpo), tag: 'aud' }, webpush: { headers: { Urgency: 'high', TTL: '86400' } } } }) };
    });
    var resp = UrlFetchApp.fetchAll(reqs), muertos = {}, ok = 0;
    resp.forEach(function (r, i) {
      if (r.getResponseCode() === 200) ok++;
      else if (r.getResponseCode() === 404 || /UNREGISTERED|INVALID_ARGUMENT/.test(r.getContentText())) muertos[tokens[i]] = 1;
    });
    if (Object.keys(muertos).length) borrarFilas_(sheet_('DISPOSITIVOS', DISP_COLS), DISP_COLS, 'Token', muertos);
    return ok;
  } catch (e) { return 0; }
}

function tokensDe_(rol, supervisor) {
  var n = norm_(supervisor);
  return rows_(sheet_('DISPOSITIVOS', DISP_COLS), DISP_COLS).filter(function (r) {
    return r['Rol'] === rol && (rol === 'admin' || norm_(r['Supervisor']) === n);
  }).map(function (r) { return String(r['Token']); });
}
function avisarAdmin_(titulo, cuerpo) { enviarPush_(tokensDe_('admin'), titulo, cuerpo); }
// nombres: un nombre por cada orden; mensaje(c) arma el texto con la cantidad de órdenes de ese supervisor
function avisarSupervisores_(nombres, titulo, mensaje) {
  var cuenta = {}; nombres.forEach(function (n) { cuenta[n] = (cuenta[n] || 0) + 1; });
  Object.keys(cuenta).forEach(function (n) { enviarPush_(tokensDe_('supervisor', n), titulo, mensaje(cuenta[n])); });
}

// El teléfono entrega su token de Firebase; se asocia al rol y supervisor de la clave con la que entró
function registrarPush_(req, who) {
  var token = String(req.token || '').trim();
  if (!token) return { ok: false, error: 'Falta el token' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sheet_('DISPOSITIVOS', DISP_COLS), existente = rows_(sh, DISP_COLS).filter(function (r) { return String(r['Token']) === token; });
    var fila = [token, who.role, who.supervisor || '', new Date()];
    if (existente.length) sh.getRange(existente[0]._row, 1, 1, 4).setValues([fila]);
    else sh.appendRow(fila);
  } finally { lock.releaseLock(); }
  var enviados = enviarPush_([token], 'Notificaciones activadas', who.role === 'admin' ? 'Te avisaremos cuando lleguen auditorías.' : 'Te avisaremos cuando tengas órdenes nuevas.');
  return { ok: true, prueba: enviados > 0 };
}

// Al cerrar sesión el teléfono deja de recibir avisos: se borra su token de la lista de dispositivos
function desregistrarPush_(req) {
  var token = String(req.token || '').trim();
  if (!token) return { ok: false, error: 'Falta el token' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var set = {}; set[token] = 1;
    return { ok: true, quitados: borrarFilas_(sheet_('DISPOSITIVOS', DISP_COLS), DISP_COLS, 'Token', set) };
  } finally { lock.releaseLock(); }
}
