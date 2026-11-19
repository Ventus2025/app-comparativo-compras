/**
 * Repositorio en línea del Comparativo Comercial.
 * Cada proceso se guarda en la hoja "Procesos" de un archivo de Hojas de cálculo del propietario del script.
 * Perfiles (pestaña "Usuarios"), árbol COLOMBIA > grupo (OBRA, BOGOTÁ) > equipo (Malambo, Campano…):
 *   rol "admin" ve todo; rol "transversal" ve todos los procesos de su grupo;
 *   un comprador ve los de su equipo (o solo los suyos si no tiene equipo).
 * Los procesos ajenos siempre se abren en solo lectura; solo el creador modifica.
 * CMO: el creador envía el proceso al CMO del proyecto (hoja "CMO"); el CMO invitado solo puede cambiar la TRM, el presupuesto,
 *   el flujo de caja, "No hay presupuesto" y el rubro / presupuesto de lanzamiento de cada línea. Los CMO ven en solo lectura los procesos de su proyecto
 *   (según el centro de costo). Los cambios de dos personas a la vez se combinan campo por campo (mezcla_).
 * La hoja "Versiones" guarda copias periódicas de cada proceso para poder restaurarlas.
 * Aprobación del informe (hoja "Aprobaciones"): el creador envía el informe a un flujo de aprobadores, que aprueban uno tras otro.
 *   Mientras está en aprobación o aprobado, el proceso no se puede modificar. Si alguien rechaza (con motivo), vuelve al creador;
 *   al reenviarlo empieza una ronda nueva y aprueban todos otra vez. Los aprobadores ven el proceso en solo lectura.
 * Los informes PDF descargados se copian en la carpeta de Drive "Comparativos – Informes PDF" (una subcarpeta por CC)
 * y quedan registrados en la hoja "Informes".
 */
const REPO_PROP = 'REPO_SHEET_ID';
const REPO_PDF_PROP = 'REPO_PDF_FOLDER_ID';
const REPO_ADMIN_INICIAL = 'dbohorquez@ventus.global';
const REPO_COLS = ['id', 'cc', 'descripcion', 'gestor', 'proveedores', 'creador', 'creado', 'modificado', 'modificadoPor', 'version', 'archivado'];
const REPO_NC = REPO_COLS.length;
const REPO_VCOLS = ['idVersion', 'proceso', 'fecha', 'usuario', 'version', 'cc'];
const REPO_VNC = REPO_VCOLS.length;
const REPO_VMAX = 30; // versiones que se conservan por proceso
const REPO_VMIN = 600; // segundos mínimos entre versiones automáticas de un proceso
const REPO_CHUNK = 45000; // una celda admite 50.000 caracteres; el JSON se reparte en varias columnas

// CMO de cada proyecto (obra y Bogotá): correo, nombre, proyecto y códigos de centro de costo del proyecto
const REPO_CMO = [
  ['nparra@ventus.global', 'Nikolas', 'Suria', ['359', '129.01']],
  ['misilva@ventus.global', 'Miguel', 'Campano', ['147.01', '147.02', '147.03']],
  ['jmurcia@ventus.global', 'Juan Felipe', 'Malambo', ['104.01', '104.02']],
  ['cperdomo@ventus.global', 'Carlos', 'Carmelo II', ['121.52']],
  ['acano@ventus.global', 'Andrés', 'Wimke', ['91.01']]
];
// Los correos salen a nombre de Luis (procesos de obra) u Oscar (procesos de Bogotá), con respuesta y copia a ellos
const REPO_REMITE = { obra: ['lquintero@ventus.global', 'Luis Quintero'], bogota: ['ovargas@ventus.global', 'Oscar Vargas'] };
const REPO_CMOCOLS = ['proceso', 'cmo', 'nombre', 'enviadoPor', 'enviado', 'estado', 'completado'];
// Aprobación del informe: estado = 'en aprobacion' | 'aprobado' | 'devuelto' (alguien rechazó) | 'retirado' (el creador lo retiró)
const REPO_APRCOLS = ['proceso', 'estado', 'ronda', 'enviadoPor', 'enviado', 'version', 'flujo', 'historial'];
const REPO_APRMAX = 10; // aprobadores por flujo
const REPO_DOM = /^[^@\s]+@(ventus\.global|ventusconstrucciones\.com)$/;

/** Ejecutar una vez desde el editor para autorizar los permisos y crear la hoja del repositorio. */
function autorizar() {
  const ss = repo_();
  Logger.log('Repositorio listo: ' + ss.getUrl());
  Logger.log('Carpeta de informes PDF: ' + carpetaPDF_().getUrl());
  Logger.log('Usuario: ' + Session.getActiveUser().getEmail());
  cmoSh_(ss);
  aprSh_(ss);
  Logger.log('Correos disponibles hoy (solicitudes a CMO): ' + MailApp.getRemainingDailyQuota());
}

function carpetaPDF_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(REPO_PDF_PROP);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* la carpeta ya no existe: se crea otra */ }
  }
  const f = DriveApp.createFolder('Comparativos – Informes PDF');
  props.setProperty(REPO_PDF_PROP, f.getId());
  return f;
}

/** Guarda en Drive una copia del informe PDF que el usuario acaba de descargar y la registra en la hoja "Informes". */
function repoGuardarPDF(id, nombre, b64) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const m = meta_(sh.getRange(f, 1, 1, REPO_NC).getValues()[0]);
  if (!acceso_(ss, e, m)) throw new Error('No tiene permiso para ver este proceso.');
  const archivo = String(nombre || 'informe.pdf').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 150);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const raiz = carpetaPDF_(), sub = String(m.cc || 'Sin CC').replace(/[\\/]/g, '-').slice(0, 100);
    const it = raiz.getFoldersByName(sub), dir = it.hasNext() ? it.next() : raiz.createFolder(sub);
    const file = dir.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'application/pdf', archivo));
    let lg = ss.getSheetByName('Informes');
    if (!lg) {
      lg = ss.insertSheet('Informes');
      lg.getRange(1, 1, 1, 7).setValues([['fecha', 'proceso', 'cc', 'usuario', 'archivo', 'url', 'version']]).setFontWeight('bold');
      lg.setFrozenRows(1);
    }
    lg.appendRow([new Date(), id, txt_(m.cc), e, archivo, file.getUrl(), m.version]);
    return { url: file.getUrl(), carpeta: raiz.getName() + ' / ' + sub };
  } finally {
    lock.releaseLock();
  }
}

function repo_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(REPO_PROP);
  if (id) return SpreadsheetApp.openById(id);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const again = props.getProperty(REPO_PROP);
    if (again) return SpreadsheetApp.openById(again);
    const ss = SpreadsheetApp.create('Comparativos – Repositorio');
    const p = ss.getSheets()[0].setName('Procesos');
    p.getRange(1, 1, 1, REPO_NC + 1).setValues([REPO_COLS.concat(['datos'])]).setFontWeight('bold');
    p.setFrozenRows(1);
    const u = ss.insertSheet('Usuarios');
    u.getRange(1, 1, 2, 6).setValues([
      ['correo', 'rol', 'nota', 'equipo', 'grupo', 'nombre'],
      [REPO_ADMIN_INICIAL, 'admin', 'rol "admin" ve los procesos de todos (solo lectura). Quien no esté aquí ve solo los suyos.', '', 'OBRA', '']
    ]);
    u.getRange(1, 1, 1, 6).setFontWeight('bold');
    u.setFrozenRows(1);
    props.setProperty(REPO_PROP, ss.getId());
    return ss;
  } finally {
    lock.releaseLock();
  }
}

function yo_() {
  const e = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!e) throw new Error('No se pudo identificar su correo. Abra la aplicación con su cuenta @ventus.global o @ventusconstrucciones.com.');
  return e;
}

// Comparación sin mayúsculas ni tildes ("Bogotá" = "BOGOTA")
function nrm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function lista_(s) {
  return String(s == null ? '' : s).split(',').map(x => x.trim()).filter(Boolean);
}

// Lee la pestaña "Usuarios" por nombre de columna; agrega las columnas que falten (hojas creadas con versiones anteriores)
function usuarios_(ss) {
  const sh = ss.getSheetByName('Usuarios');
  const v = sh.getDataRange().getValues();
  const head = v[0].map(h => nrm_(h));
  ['equipo', 'grupo', 'nombre'].forEach(c => {
    if (head.indexOf(c) < 0) {
      head.push(c);
      sh.getRange(1, head.length).setValue(c).setFontWeight('bold');
    }
  });
  const col = c => head.indexOf(c);
  const rows = v.slice(1).map(r => ({
    e: String(r[Math.max(0, col('correo'))] || '').trim().toLowerCase(),
    rol: nrm_(r[col('rol') < 0 ? 1 : col('rol')]),
    eq: lista_(r[col('equipo')]),
    gr: lista_(r[col('grupo')]),
    n: String(r[col('nombre')] || '').trim()
  }));
  return { sh: sh, head: head, rows: rows };
}

/**
 * Perfil del usuario (árbol COLOMBIA > grupo > equipo):
 *   admin: ve todo; transversal: ve a todos los de su(s) grupo(s); equipo: ve a sus compañeros de equipo.
 * ve = correos cuyos procesos puede ver (siempre incluye el propio).
 */
function perfil_(ss, email) {
  const rows = usuarios_(ss).rows;
  const mine = rows.filter(r => r.e === email);
  const uniq = a => a.filter((x, i) => a.findIndex(y => nrm_(y) === nrm_(x)) === i);
  const equipos = uniq(mine.flatMap(r => r.eq)), eqs = equipos.map(nrm_);
  const tgr = uniq(mine.filter(r => r.rol === 'transversal').flatMap(r => r.gr)), tgs = tgr.map(nrm_);
  const ve = [email];
  rows.forEach(r => {
    if (!r.e || ve.indexOf(r.e) >= 0) return;
    if (r.eq.some(q => eqs.indexOf(nrm_(q)) >= 0) || r.gr.some(g => tgs.indexOf(nrm_(g)) >= 0)) ve.push(r.e);
  });
  // "equipos" es lo que la pantalla muestra junto al correo
  return { admin: mine.some(r => r.rol === 'admin'), equipos: equipos.concat(tgr.map(g => g + ' (transversal)')), grupos: uniq(mine.flatMap(r => r.gr)), ve: ve };
}

// Usuarios iniciales del árbol COLOMBIA (grupo, equipo, rol). Se cargan con cargarUsuarios().
const REPO_USUARIOS = [
  ['dfranco@ventus.global', 'Dayana Franco', 'OBRA', 'Malambo', 'comprador'],
  ['carguello@ventus.global', 'Camilo Arguello', 'OBRA', 'Malambo', 'comprador'],
  ['lmunoz@ventus.global', 'Linda Muñoz', 'OBRA', 'Malambo', 'comprador'],
  ['ghumo@ventus.global', 'Geraldine Humo', 'OBRA', 'Campano', 'comprador'],
  ['cmoreno@ventus.global', 'Jean Carlos Moreno', 'OBRA', 'Campano', 'comprador'],
  ['dmolina@ventus.global', 'Doris Molina', 'OBRA', 'Campano', 'comprador'],
  ['mrojas@ventus.global', 'Miguel Rojas', 'OBRA', 'Suria', 'comprador'],
  ['emoreno@ventus.global', 'Edgar Moreno', 'OBRA', 'Suria', 'comprador'],
  ['kromero@ventus.global', 'Karina Romero', 'OBRA', 'Wimke', 'comprador'],
  ['cgalvis@ventus.global', 'Claudia Galvis', 'OBRA', 'Wimke', 'comprador'],
  ['jocardona@ventus.global', 'Johanna Cardona', 'OBRA', 'Carmelo II', 'comprador'],
  ['dbohorquez@ventus.global', 'Daniela Bohórquez M', 'OBRA', '', 'transversal'],
  ['lquintero@ventus.global', 'Luis Ernesto Quintero', 'OBRA', '', 'transversal'],
  ['ahernandez@ventus.global', 'Angela Hernandez', 'OBRA', '', 'transversal'],
  ['ovargas@ventus.global', 'Oscar Vargas', 'BOGOTÁ', '', 'transversal'],
  ['ccortes@ventus.global', 'Carlos Alfredo Cortés', 'BOGOTÁ', '', 'comprador'],
  ['jbermudez@ventus.global', 'Johan Bermúdez', 'BOGOTÁ', '', 'comprador'],
  ['jjaimes@ventus.global', 'Jessica Jaimes', 'BOGOTÁ', '', 'comprador'],
  ['ctraslavina@ventus.global', 'Claudia Traslaviña', 'BOGOTÁ', '', 'comprador'],
  ['csalazar@ventus.global', 'Cindy Salazar', 'BOGOTÁ', '', 'comprador']
];

/**
 * Ejecutar una vez desde el editor: escribe en "Usuarios" el grupo, equipo, rol y nombre de REPO_USUARIOS.
 * Actualiza las filas que ya existen (por correo) y agrega las que faltan; no borra a nadie
 * y no le quita el rol "admin" a quien ya lo tenga.
 */
function cargarUsuarios() {
  const ss = repo_(), u = usuarios_(ss), sh = u.sh, head = u.head;
  const col = c => head.indexOf(c) + 1;
  const datos = sh.getDataRange().getValues();
  const filaDe = {};
  datos.slice(1).forEach((r, i) => { const e = String(r[Math.max(0, col('correo') - 1)] || '').trim().toLowerCase(); if (e) filaDe[e] = i + 2; });
  let nuevos = 0, actualizados = 0, sig = sh.getLastRow() + 1;
  REPO_USUARIOS.forEach(([correo, nombre, grupo, equipo, rol]) => {
    let f = filaDe[correo];
    if (f) actualizados++; else { f = sig++; nuevos++; sh.getRange(f, col('correo')).setValue(correo); }
    const actual = nrm_(sh.getRange(f, col('rol')).getValue());
    sh.getRange(f, col('rol')).setValue(actual === 'admin' ? 'admin' : rol);
    sh.getRange(f, col('grupo')).setValue(grupo);
    sh.getRange(f, col('equipo')).setValue(equipo);
    sh.getRange(f, col('nombre')).setValue(nombre);
  });
  Logger.log(`Usuarios: ${actualizados} actualizados, ${nuevos} agregados. Hoja: ${ss.getUrl()}`);
}

function puedeVer_(pf, creador) {
  return pf.admin || pf.ve.indexOf(creador) >= 0;
}

// ---- CMO
function cmoSh_(ss) {
  let sh = ss.getSheetByName('CMO');
  if (!sh) {
    sh = ss.insertSheet('CMO');
    sh.getRange(1, 1, 1, REPO_CMOCOLS.length).setValues([REPO_CMOCOLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function cmoFila_(r) {
  const o = {};
  REPO_CMOCOLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].toISOString() : String(r[i] == null ? '' : r[i]); });
  return o;
}

// Invitaciones de todos los procesos: { idProceso: {cmo, nombre, enviadoPor, enviado, estado, completado} }
function cmoTodas_(ss) {
  const sh = cmoSh_(ss), n = sh.getLastRow() - 1, o = {};
  if (n > 0) sh.getRange(2, 1, n, REPO_CMOCOLS.length).getValues().forEach(r => { if (r[0]) o[r[0]] = cmoFila_(r); });
  return o;
}

// Invitación de un proceso (se guarda en caché porque el pulso la consulta seguido)
function cmoDe_(ss, id) {
  const c = CacheService.getScriptCache(), k = 'pc_' + id, x = c.get(k);
  if (x) return JSON.parse(x) || null;
  const sh = cmoSh_(ss), f = fila_(sh, id), inv = f ? cmoFila_(sh.getRange(f, 1, 1, REPO_CMOCOLS.length).getValues()[0]) : null;
  c.put(k, JSON.stringify(inv), 21600);
  return inv;
}

function cmoGuardar_(ss, inv) {
  const sh = cmoSh_(ss), f = fila_(sh, inv.proceso) || sh.getLastRow() + 1;
  sh.getRange(f, 1, 1, REPO_CMOCOLS.length).setValues([REPO_CMOCOLS.map(c => inv[c] || '')]);
  CacheService.getScriptCache().put('pc_' + inv.proceso, JSON.stringify(inv), 21600);
}

// ¿el correo es CMO del proyecto de este centro de costo? (el CC se guarda como "359 PV Suria")
function cmoProy_(e, cc) {
  const cod = String(cc || '').trim().split(/\s+/)[0];
  return REPO_CMO.some(c => c[0] === e && c[3].indexOf(cod) >= 0);
}

/** Qué puede hacer el usuario con el proceso: 'dueno' (todo), 'cmo' (sus campos), 'aprobador' (ver y aprobar), 'lector' (ver) o '' (nada). */
function acceso_(ss, e, m, inv, pf, apr) {
  if (m.creador === e) return 'dueno';
  if (inv === undefined) inv = cmoDe_(ss, m.id);
  if (inv && inv.cmo === e) return 'cmo';
  if (apr === undefined) apr = aprDe_(ss, m.id);
  if (apr && (apr.flujo.some(x => x.e === e) || apr.historial.some(h => h.e === e))) return 'aprobador';
  if (puedeVer_(pf || perfil_(ss, e), m.creador) || cmoProy_(e, m.cc)) return 'lector';
  return '';
}

// Campos que puede cambiar el CMO: TRM, hoja 2 (presupuesto y "No hay presupuesto"), hoja 3 (flujo de caja) y rubro / presupuesto de lanzamiento de cada línea
function cmoCampo_(p) {
  return p[0] === 'trm' || p[0] === 'bud' || p[0] === 'fl' || p[0] === 'nb' || (p[0] === 'items' && p.length === 3 && ['bi', 'bt', 'pb'].indexOf(p[2]) >= 0);
}

/**
 * Mezcla de tres vías: b = lo que el usuario tenía al empezar, m = lo suyo ahora, t = lo guardado por otro.
 * Lo que el usuario no tocó queda como lo dejó el otro; si ambos cambiaron el mismo dato, gana el usuario.
 * ok(ruta) limita qué rutas puede cambiar el usuario (CMO); las demás quedan como están guardadas.
 */
function mezcla_(b, m, t, p, ok) {
  const js = JSON.stringify, obj = x => x !== null && typeof x === 'object', libre = !ok || ok(p);
  if (js(m) === js(b)) return t;
  if (libre) {
    if (js(t) === js(b)) return m;
    if (!obj(m) || !obj(t) || !obj(b) || Array.isArray(m) !== Array.isArray(t) || Array.isArray(m) !== Array.isArray(b)) return m;
  } else if (!obj(m) || !obj(t) || !obj(b) || Array.isArray(m) !== Array.isArray(t)) return t;
  if (Array.isArray(t)) {
    const n = libre && m.length !== b.length ? m.length : t.length, r = [];
    for (let i = 0; i < n; i++) r.push(mezcla_(b[i], m[i], t[i], p.concat(i), ok));
    return r;
  }
  const r = {};
  Object.keys(Object.assign({}, t, libre ? m : {})).forEach(k => {
    const v = mezcla_(b[k], m[k], t[k], p.concat(k), ok);
    if (v !== undefined) r[k] = v;
  });
  return r;
}

function nombreDe_(ss, e) {
  const c = REPO_CMO.find(x => x[0] === e);
  if (c) return c[1];
  const u = usuarios_(ss).rows.find(r => r.e === e && r.n) || REPO_USUARIOS.map(r => ({ e: r[0], n: r[1] })).find(r => r.e === e);
  return u ? u.n : e.split('@')[0];
}

// Los procesos de compradores de Bogotá salen a nombre de Oscar; los demás (obra), a nombre de Luis
function remite_(ss, creador) {
  const b = perfil_(ss, creador).grupos.some(g => nrm_(g) === 'bogota');
  return REPO_REMITE[b ? 'bogota' : 'obra'];
}

function correo_(ss, creador, para, asunto, titulo, cuerpo, url, boton) {
  const r = remite_(ss, creador);
  const html = `<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:auto;border:1px solid #d5dbd8;border-radius:10px;overflow:hidden">
<div style="background:#33634a;color:#fff;padding:16px 22px;border-bottom:4px solid #e07a00"><div style="font-size:12px;opacity:.8">VENTUS · Comparativo Comercial</div><div style="font-size:19px;font-weight:bold;margin-top:4px">${titulo}</div></div>
<div style="padding:20px 22px;color:#1d2421;font-size:14px;line-height:1.55">${cuerpo}
${url ? `<p style="margin:22px 0 8px"><a href="${url}" style="background:#009a44;color:#fff;text-decoration:none;padding:11px 20px;border-radius:8px;font-weight:bold;display:inline-block">${boton || 'Abrir el proceso'}</a></p>` : ''}
<p style="color:#66716c;font-size:12px;margin-top:20px">Enviado desde el Comparativo Comercial de Ventus a nombre de ${r[1]}. Si tiene dudas, responda este correo.</p></div></div>`;
  MailApp.sendEmail({ to: para, cc: r[0] === para ? '' : r[0], replyTo: r[0], name: r[1] + ' – Compras Ventus', subject: asunto, htmlBody: html });
}

function enlace_(id) {
  let u = '';
  try { u = ScriptApp.getService().getUrl() || ''; } catch (e) { /* sin implementación publicada */ }
  return u ? u + '?p=' + encodeURIComponent(id) : '';
}

const esc_ = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function procesoCmo_(ss, id) {
  const sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const r = sh.getRange(f, 1, 1, sh.getLastColumn()).getValues()[0];
  return { m: meta_(r), S: JSON.parse(unir_(r.slice(REPO_NC)) || '{}') };
}

/** El creador envía el proceso a un CMO: queda pendiente y le llega un correo con el enlace. */
function repoEnviarCmo(id, para) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id), m = x.m, S = x.S;
  if (m.creador !== e) throw new Error('Solo quien creó el proceso puede enviarlo a CMO.');
  para = String(para || '').trim().toLowerCase();
  if (!/^[^@\s]+@(ventus\.global|ventusconstrucciones\.com)$/.test(para)) throw new Error('Escriba un correo @ventus.global o @ventusconstrucciones.com válido.');
  if (para === e) throw new Error('El CMO debe ser otra persona.');
  const inv = { proceso: id, cmo: para, nombre: nombreDe_(ss, para), enviadoPor: e, enviado: new Date().toISOString(), estado: 'pendiente', completado: '' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { cmoGuardar_(ss, inv); } finally { lock.releaseLock(); }
  const proc = [S.cc, S.desc].filter(Boolean).join(' – ') || 'sin nombre';
  let aviso = '';
  try {
    correo_(ss, m.creador, para, `Te han solicitado llenar el presupuesto del proceso ${S.cc || ''}`.trim(), 'Te han solicitado llenar el presupuesto',
      `<p>Hola ${esc_(inv.nombre)},</p><p><b>${esc_(nombreDe_(ss, e))}</b> te solicitó llenar el presupuesto del proceso:</p>
<p style="background:#f2f7f4;border-left:4px solid #009a44;padding:10px 14px;margin:12px 0"><b>${esc_(proc)}</b>${S.gestor ? `<br><span style="color:#66716c">Gestor de compra: ${esc_(S.gestor)}</span>` : ''}</p>
<p>Debes completar (aparece sombreado en amarillo):</p><ul style="padding-left:20px;margin:6px 0"><li><b>1. Datos:</b> la TRM del proyecto (si ya está, revisa que sea la correcta)</li><li><b>2. Presupuesto</b></li><li><b>3. Flujo de caja</b> (opcional)</li><li><b>5. Precios:</b> el rubro y el presupuesto de lanzamiento de cada línea</li></ul>
<p>Puedes trabajar al mismo tiempo que el gestor. Cuando termines, haz clic en <b>“Proceso por CMO completo”</b>.</p>`, enlace_(id));
  } catch (err) { aviso = 'La solicitud quedó registrada, pero no se pudo enviar el correo: ' + (err && err.message || err); }
  return { inv: inv, aviso: aviso };
}

/** El CMO invitado marca su parte como completa; se le avisa al creador por correo. */
function repoCmoCompleto(id) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id), m = x.m, S = x.S, inv = cmoDe_(ss, id);
  if (!inv || inv.cmo !== e) throw new Error('Solo el CMO al que se le envió el proceso puede marcar el proceso como completo.');
  inv.estado = 'completo';
  inv.completado = new Date().toISOString();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { cmoGuardar_(ss, inv); } finally { lock.releaseLock(); }
  let aviso = '';
  try {
    correo_(ss, m.creador, m.creador, `CMO completó el presupuesto del proceso ${S.cc || ''}`.trim(), 'El CMO completó su parte',
      `<p><b>${esc_(inv.nombre)}</b> (${esc_(e)}) marcó como completo el presupuesto del proceso:</p>
<p style="background:#f2f7f4;border-left:4px solid #009a44;padding:10px 14px;margin:12px 0"><b>${esc_([S.cc, S.desc].filter(Boolean).join(' – ') || 'sin nombre')}</b></p>
<p>Ya puedes revisar la TRM, el presupuesto, el flujo de caja y el rubro / presupuesto de lanzamiento de cada línea.</p>`, enlace_(id));
  } catch (err) { aviso = 'Quedó marcado como completo, pero no se pudo avisar por correo: ' + (err && err.message || err); }
  return { inv: inv, aviso: aviso };
}

// ---- Aprobación del informe de adjudicación
// Una fila por proceso en la hoja "Aprobaciones": flujo = aprobadores de la ronda actual, en orden
// ({e, n, cargo, estado: 'pendiente' | 'aprobado' | 'rechazado', fecha, motivo}); historial = trazabilidad de todas las rondas
// ({t, e, n, cargo, acc: 'envio' | 'aprobo' | 'rechazo' | 'retiro', ronda, motivo}).
function aprSh_(ss) {
  let sh = ss.getSheetByName('Aprobaciones');
  if (!sh) {
    sh = ss.insertSheet('Aprobaciones');
    sh.getRange(1, 1, 1, REPO_APRCOLS.length).setValues([REPO_APRCOLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function aprFila_(r) {
  const o = {}, js = v => { try { const x = JSON.parse(String(v || '[]')); return Array.isArray(x) ? x : []; } catch (e) { return []; } };
  REPO_APRCOLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].toISOString() : String(r[i] == null ? '' : r[i]); });
  o.ronda = Number(o.ronda) || 0;
  o.version = Number(o.version) || 0;
  o.flujo = js(o.flujo);
  o.historial = js(o.historial);
  return o;
}

function aprTodas_(ss) {
  const sh = aprSh_(ss), n = sh.getLastRow() - 1, o = {};
  if (n > 0) sh.getRange(2, 1, n, REPO_APRCOLS.length).getValues().forEach(r => { if (r[0]) o[r[0]] = aprFila_(r); });
  return o;
}

// Aprobación de un proceso (en caché, porque el pulso la consulta seguido); fresco = leer la hoja sin caché
function aprDe_(ss, id, fresco) {
  const c = CacheService.getScriptCache(), k = 'pa_' + id, x = fresco ? null : c.get(k);
  if (x) return JSON.parse(x) || null;
  const sh = aprSh_(ss), f = fila_(sh, id), a = f ? aprFila_(sh.getRange(f, 1, 1, REPO_APRCOLS.length).getValues()[0]) : null;
  c.put(k, JSON.stringify(a), 21600);
  return a;
}

function aprGuardar_(ss, a) {
  const sh = aprSh_(ss), f = fila_(sh, a.proceso) || sh.getLastRow() + 1;
  if (a.historial.length > 300) a.historial = a.historial.slice(-300);
  sh.getRange(f, 1, 1, REPO_APRCOLS.length).setValues([REPO_APRCOLS.map(c => c === 'flujo' || c === 'historial' ? JSON.stringify(a[c]) : a[c] == null ? '' : a[c])]);
  CacheService.getScriptCache().put('pa_' + a.proceso, JSON.stringify(a), 21600);
}

// en aprobación o aprobado: el proceso no se puede modificar
function aprBloq_(a) {
  return !!a && (a.estado === 'en aprobacion' || a.estado === 'aprobado');
}

function aprNoBloq_(ss, id) {
  const a = aprDe_(ss, id);
  if (!aprBloq_(a)) return;
  throw new Error(a.estado === 'aprobado'
    ? 'El informe ya está aprobado y el proceso no se puede modificar. Si necesita cambiarlo, use "Retirar aprobación" en la hoja 11 (se anulan las aprobaciones).'
    : 'El proceso está en aprobación y no se puede modificar. Si necesita cambiarlo, use "Retirar de aprobación" en la hoja 11 (se anulan las aprobaciones de esta ronda).');
}

// a quién le toca aprobar ahora (el primero pendiente, en orden)
function aprActual_(a) {
  return a && a.estado === 'en aprobacion' ? a.flujo.find(x => x.estado === 'pendiente') || null : null;
}

// resumen para el Historial
function aprRes_(a) {
  if (!a) return null;
  const cur = aprActual_(a);
  return { estado: a.estado, ronda: a.ronda, ok: a.flujo.filter(x => x.estado === 'aprobado').length, n: a.flujo.length, act: cur ? { e: cur.e, n: cur.n, cargo: cur.cargo } : null };
}

const aprProc_ = S => [S.cc, S.desc].filter(Boolean).join(' – ') || 'sin nombre';
const aprCaja_ = (S, extra) => `<p style="background:#f2f7f4;border-left:4px solid #009a44;padding:10px 14px;margin:12px 0"><b>${esc_(aprProc_(S))}</b>${S.gestor ? `<br><span style="color:#66716c">Gestor de compra: ${esc_(S.gestor)}</span>` : ''}${extra || ''}</p>`;
const aprQuien_ = x => esc_(x.n) + (x.cargo ? ' (' + esc_(x.cargo) + ')' : '');

// correo a quien le toca aprobar
function aprAvisar_(ss, m, S, a) {
  const cur = aprActual_(a);
  if (!cur) return;
  const i = a.flujo.indexOf(cur), antes = a.flujo.slice(0, i);
  correo_(ss, m.creador, cur.e, `Aprobación pendiente: informe de adjudicación ${S.cc || ''}`.trim(), 'Tienes un informe de adjudicación por aprobar',
    `<p>Hola ${esc_(cur.n)},</p><p><b>${esc_(nombreDe_(ss, m.creador))}</b> envió a aprobación el informe de adjudicación del proceso:</p>${aprCaja_(S)}
<p>Eres el paso <b>${i + 1} de ${a.flujo.length}</b> del flujo de aprobación${cur.cargo ? ` (${esc_(cur.cargo)})` : ''}.${antes.length ? ` Ya aprobaron: ${antes.map(aprQuien_).join(', ')}.` : ''}</p>
<p>Abre el proceso (lo verás en solo lectura), revisa el informe en la hoja <b>11. Informe de adjudicación</b> y usa <b>“Aprobar”</b> o <b>“Rechazar”</b>. Si rechazas, escribe el motivo: el proceso vuelve al gestor para que lo corrija.</p>`, enlace_(m.id), 'Revisar y aprobar');
}

/** El creador envía el informe al flujo de aprobación (lista en orden de {e, cargo}). Empieza una ronda nueva. */
function repoEnviarAprobacion(id, lista) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id), m = x.m, S = x.S;
  if (m.creador !== e) throw new Error('Solo quien creó el proceso puede enviarlo a aprobación.');
  const vistos = [], flujo = (lista || []).map(p => {
    const c = String(p && p.e || '').trim().toLowerCase(), cargo = String(p && p.cargo || '').trim().slice(0, 80);
    if (!REPO_DOM.test(c)) throw new Error(`"${c || '(vacío)'}" no es un correo @ventus.global o @ventusconstrucciones.com válido.`);
    if (c === e) throw new Error('Usted no puede aprobar su propio proceso: quite su correo del flujo.');
    if (vistos.indexOf(c) >= 0) throw new Error(`${c} está dos veces en el flujo.`);
    vistos.push(c);
    return { e: c, n: nombreDe_(ss, c), cargo: cargo, estado: 'pendiente', fecha: '', motivo: '' };
  });
  if (!flujo.length) throw new Error('Agregue al menos una persona al flujo de aprobación.');
  if (flujo.length > REPO_APRMAX) throw new Error(`El flujo admite máximo ${REPO_APRMAX} aprobadores.`);
  let a;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const prev = aprDe_(ss, id, true);
    if (aprBloq_(prev)) throw new Error(prev.estado === 'aprobado' ? 'El informe ya está aprobado.' : 'El informe ya está en aprobación.');
    const now = new Date().toISOString(), ronda = (prev ? prev.ronda : 0) + 1;
    a = { proceso: id, estado: 'en aprobacion', ronda: ronda, enviadoPor: e, enviado: now, version: m.version, flujo: flujo,
      historial: (prev ? prev.historial : []).concat([{ t: now, e: e, n: nombreDe_(ss, e), cargo: 'Gestor de compra', acc: 'envio', ronda: ronda, motivo: flujo.map(f => f.n + (f.cargo ? ' (' + f.cargo + ')' : '')).join(' → ') }]) };
    aprGuardar_(ss, a);
    snap_(ss, id, e, m.version, m.cc, JSON.stringify(S)); // copia de lo que se envió a aprobar
  } finally {
    lock.releaseLock();
  }
  let aviso = '';
  try { aprAvisar_(ss, m, S, a); } catch (err) { aviso = 'El flujo quedó registrado, pero no se pudo enviar el correo al primer aprobador: ' + (err && err.message || err); }
  return { apr: a, aviso: aviso };
}

/** Quien tiene el turno aprueba (comentario opcional) o rechaza (motivo obligatorio). */
function repoDecidirAprobacion(id, aprueba, motivo) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id), m = x.m, S = x.S;
  motivo = String(motivo || '').trim().slice(0, 2000);
  if (!aprueba && !motivo) throw new Error('Escriba el motivo del rechazo.');
  let a, cur;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    a = aprDe_(ss, id, true);
    cur = aprActual_(a);
    if (!cur) throw new Error(a && a.estado === 'aprobado' ? 'El informe ya está aprobado.' : 'El informe ya no está en aprobación (quizás el gestor lo retiró o alguien lo rechazó).');
    if (cur.e !== e) throw new Error(a.flujo.some(f => f.e === e && f.estado === 'pendiente') ? `Todavía no es su turno: falta la aprobación de ${cur.n}.` : 'Usted no tiene una aprobación pendiente en este proceso.');
    const now = new Date().toISOString();
    cur.estado = aprueba ? 'aprobado' : 'rechazado';
    cur.fecha = now;
    cur.motivo = motivo;
    a.historial.push({ t: now, e: e, n: cur.n, cargo: cur.cargo, acc: aprueba ? 'aprobo' : 'rechazo', ronda: a.ronda, motivo: motivo });
    if (!aprueba) a.estado = 'devuelto';
    else if (!aprActual_(a)) a.estado = 'aprobado';
    aprGuardar_(ss, a);
  } finally {
    lock.releaseLock();
  }
  let aviso = '';
  try {
    if (!aprueba) {
      correo_(ss, m.creador, m.creador, `Informe de adjudicación devuelto: ${S.cc || ''}`.trim(), 'Te devolvieron el informe de adjudicación',
        `<p><b>${aprQuien_(cur)}</b> rechazó el informe de adjudicación del proceso:</p>${aprCaja_(S)}
<p style="background:#fdeeee;border-left:4px solid #d9534f;padding:10px 14px;margin:12px 0"><b>Motivo:</b><br>${esc_(motivo).replace(/\n/g, '<br>')}</p>
<p>El proceso quedó desbloqueado: corrígelo y vuelve a enviarlo a aprobación desde la hoja <b>11. Informe de adjudicación</b>. Todos los aprobadores tendrán que aprobar de nuevo.</p>`, enlace_(id));
    } else if (a.estado === 'aprobado') {
      correo_(ss, m.creador, m.creador, `Informe de adjudicación aprobado: ${S.cc || ''}`.trim(), 'El informe de adjudicación quedó aprobado',
        `<p>Todos los aprobadores aprobaron el informe de adjudicación del proceso:</p>${aprCaja_(S)}
<ul style="padding-left:20px;margin:6px 0">${a.flujo.map(f => `<li>${aprQuien_(f)}</li>`).join('')}</ul>
<p>Ya puedes descargar el informe en PDF desde la hoja <b>11</b>: las aprobaciones salen en el recuadro <b>APROBACIONES</b> del resumen.</p>`, enlace_(id));
    } else aprAvisar_(ss, m, S, a);
  } catch (err) { aviso = 'Su decisión quedó registrada, pero no se pudo enviar el correo: ' + (err && err.message || err); }
  return { apr: a, aviso: aviso };
}

/** El creador retira el informe de la aprobación (o retira la aprobación ya dada) para poder modificar el proceso. */
function repoRetirarAprobacion(id, motivo) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id), m = x.m, S = x.S;
  if (m.creador !== e) throw new Error('Solo quien creó el proceso puede retirarlo de la aprobación.');
  motivo = String(motivo || '').trim().slice(0, 2000);
  let a, cur;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    a = aprDe_(ss, id, true);
    if (!aprBloq_(a)) throw new Error('El informe no está en aprobación.');
    cur = aprActual_(a);
    const now = new Date().toISOString();
    a.estado = 'retirado';
    a.historial.push({ t: now, e: e, n: nombreDe_(ss, e), cargo: 'Gestor de compra', acc: 'retiro', ronda: a.ronda, motivo: motivo });
    aprGuardar_(ss, a);
  } finally {
    lock.releaseLock();
  }
  let aviso = '';
  if (cur) {
    try {
      correo_(ss, m.creador, cur.e, `Ya no se requiere tu aprobación: ${S.cc || ''}`.trim(), 'El gestor retiró el informe de la aprobación',
        `<p>Hola ${esc_(cur.n)},</p><p><b>${esc_(nombreDe_(ss, e))}</b> retiró de la aprobación el informe de adjudicación del proceso:</p>${aprCaja_(S)}
${motivo ? `<p><b>Motivo:</b> ${esc_(motivo)}</p>` : ''}<p>Por ahora no tienes que hacer nada. Si lo vuelve a enviar, te llegará un correo nuevo.</p>`, '');
    } catch (err) { aviso = 'Quedó retirado, pero no se pudo avisar por correo al aprobador: ' + (err && err.message || err); }
  }
  return { apr: a, aviso: aviso };
}

/** Pulso cada ~12 s mientras el proceso está abierto: versión guardada, quién más lo tiene abierto y estado del CMO. */
function repoPulso(id) {
  const e = yo_(), c = CacheService.getScriptCache(), k = 'pp_' + id, now = Date.now();
  let pp = {};
  try { pp = JSON.parse(c.get(k) || '{}') || {}; } catch (x) { pp = {}; }
  pp[e] = now;
  Object.keys(pp).forEach(x => { if (now - pp[x] > 40000) delete pp[x]; });
  c.put(k, JSON.stringify(pp), 300);
  let ver = c.get('pv_' + id), ss = null;
  if (!ver) {
    ss = repo_();
    const sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
    ver = String(f ? meta_(sh.getRange(f, 1, 1, REPO_NC).getValues()[0]).version : 0);
    c.put('pv_' + id, ver, 21600);
  }
  const ic = c.get('pc_' + id), ia = c.get('pa_' + id);
  return { version: Number(ver), otros: Object.keys(pp).filter(x => x !== e), cmo: ic ? JSON.parse(ic) : cmoDe_(ss || (ss = repo_()), id), apr: ia ? JSON.parse(ia) : aprDe_(ss || repo_(), id) };
}

/** Aprobación del Gerente de Compras (único ofertante): se guarda en Drive junto a los informes del CC. */
function repoAdjunto(id, nombre, b64) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id);
  if (x.m.creador !== e) throw new Error('Solo quien creó el proceso puede adjuntar la aprobación.');
  aprNoBloq_(ss, id);
  const raiz = carpetaPDF_(), sub = String(x.m.cc || 'Sin CC').replace(/[\\/]/g, '-').slice(0, 100);
  const it = raiz.getFoldersByName(sub), dir = it.hasNext() ? it.next() : raiz.createFolder(sub);
  const nm = ('Aprobación único ofertante – ' + String(nombre || 'aprobacion.pdf')).replace(/[\\/:*?"<>|]+/g, '_').slice(0, 150);
  const file = dir.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'application/pdf', nm));
  return { id: file.getId(), url: file.getUrl() };
}

/**
 * Adjuntos del proceso (ofertas, aprobación técnica, anexos, documento general): se guardan en Drive
 * (carpeta del CC / Adjuntos) y el proceso guarda su referencia en S.fx, así siguen ahí al volver a abrirlo.
 */
function repoArchivo(id, nombre, tipo, b64) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id);
  if (x.m.creador !== e) throw new Error('Solo quien creó el proceso puede adjuntar archivos.');
  aprNoBloq_(ss, id);
  const raiz = carpetaPDF_(), sub = String(x.m.cc || 'Sin CC').replace(/[\\/]/g, '-').slice(0, 100);
  const it = raiz.getFoldersByName(sub), dir = it.hasNext() ? it.next() : raiz.createFolder(sub);
  const ia = dir.getFoldersByName('Adjuntos'), ad = ia.hasNext() ? ia.next() : dir.createFolder('Adjuntos');
  const nm = String(nombre || 'adjunto').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 150);
  const file = ad.createFile(Utilities.newBlob(Utilities.base64Decode(b64), tipo || 'application/octet-stream', nm));
  return { id: file.getId(), url: file.getUrl() };
}

/** Devuelve un adjunto del proceso (solo si el proceso lo tiene registrado) para mostrarlo o armar el informe. */
function repoLeerArchivo(id, fileId) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id);
  if (!acceso_(ss, e, x.m)) throw new Error('No tiene permiso para ver este proceso.');
  const fx = x.S.fx || {}, ids = [];
  Object.keys(fx).forEach(c => Object.keys(fx[c] || {}).forEach(k => (fx[c][k] || []).forEach(f => ids.push(f.id))));
  if (ids.indexOf(fileId) < 0) throw new Error('El archivo no corresponde a este proceso.');
  return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes());
}

/** Devuelve la aprobación adjunta al proceso (solo ese archivo) para armar el informe. */
function repoLeerAdjunto(id, fileId) {
  const ss = repo_(), e = yo_(), x = procesoCmo_(ss, id);
  if (!acceso_(ss, e, x.m)) throw new Error('No tiene permiso para ver este proceso.');
  if (!x.S.uo || !x.S.uo.f || x.S.uo.f.id !== fileId) throw new Error('El archivo no corresponde a este proceso.');
  return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes());
}

function fila_(sh, id, col) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || !id) return 0;
  const ids = sh.getRange(2, col || 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

// google.script.run no transporta fechas: se devuelven como texto ISO
function meta_(r) {
  const o = {};
  REPO_COLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
  o.creador = String(o.creador || '').toLowerCase();
  o.version = Number(o.version) || 0;
  o.archivado = o.archivado === true;
  return o;
}

// Texto para las columnas de índice: recortado (el contenido completo está en "datos") y sin que = + - @ se interprete como fórmula
function txt_(v) {
  v = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

// cada trozo lleva una "J" delante para que Sheets nunca lo tome como número o fórmula
function trozos_(datos) {
  const chunks = [];
  for (let i = 0; i < datos.length; i += REPO_CHUNK) chunks.push('J' + datos.slice(i, i + REPO_CHUNK));
  return chunks;
}

function unir_(celdas) {
  return celdas.map(String).filter(x => x).map(x => x.slice(1)).join('');
}

function asegurarTam_(sh, fila, ancho) {
  if (fila > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  if (ancho > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), ancho - sh.getMaxColumns());
}

function versiones_(ss) {
  let sh = ss.getSheetByName('Versiones');
  if (!sh) {
    sh = ss.insertSheet('Versiones');
    sh.getRange(1, 1, 1, REPO_VNC + 1).setValues([REPO_VCOLS.concat(['datos'])]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Guarda una copia del proceso y deja solo las últimas REPO_VMAX de ese proceso
function snap_(ss, proceso, usuario, version, cc, datos) {
  const sh = versiones_(ss);
  const chunks = trozos_(datos);
  const f = sh.getLastRow() + 1;
  asegurarTam_(sh, f, REPO_VNC + chunks.length);
  sh.getRange(f, 1, 1, REPO_VNC + chunks.length).setValues([[Utilities.getUuid(), proceso, new Date(), usuario, version, txt_(cc)].concat(chunks)]);
  const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues(), filas = [];
  ids.forEach((r, i) => { if (r[0] === proceso) filas.push(i + 2); });
  for (let i = filas.length - REPO_VMAX - 1; i >= 0; i--) sh.deleteRow(filas[i]);
  CacheService.getScriptCache().put('v_' + proceso, '1', REPO_VMIN);
}

function repoInicio() {
  const ss = repo_(), e = yo_(), pf = perfil_(ss, e), inv = cmoTodas_(ss), apr = aprTodas_(ss), nombres = {};
  REPO_USUARIOS.forEach(r => { nombres[r[0]] = r[1]; });
  usuarios_(ss).rows.forEach(r => { if (r.e && r.n) nombres[r.e] = r.n; });
  REPO_CMO.forEach(c => { nombres[c[0]] = c[1]; });
  return {
    email: e, admin: pf.admin, equipos: pf.equipos, grupos: pf.grupos, nombres: nombres,
    cmo: REPO_CMO.map(c => ({ e: c[0], n: c[1], p: c[2], cs: c[3] })),
    pend: Object.keys(inv).filter(k => inv[k].cmo === e && inv[k].estado === 'pendiente').length,
    aprPend: Object.keys(apr).filter(k => { const a = aprActual_(apr[k]); return a && a.e === e; }).length
  };
}

function repoListar(verArchivados) {
  const ss = repo_(), e = yo_(), pf = perfil_(ss, e), sh = ss.getSheetByName('Procesos');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const inv = cmoTodas_(ss), apr = aprTodas_(ss);
  return sh.getRange(2, 1, n, REPO_NC).getValues().map(meta_)
    .filter(m => {
      if (!m.id || !(verArchivados || !m.archivado)) return false;
      m.cmo = inv[m.id] || null;
      m.rol = acceso_(ss, e, m, m.cmo, pf, apr[m.id] || null);
      m.apr = aprRes_(apr[m.id]);
      return !!m.rol;
    })
    .sort((a, b) => String(b.modificado).localeCompare(String(a.modificado)));
}

function repoAbrir(id) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const r = sh.getRange(f, 1, 1, sh.getLastColumn()).getValues()[0], m = meta_(r);
  m.cmo = cmoDe_(ss, id);
  m.apr = aprDe_(ss, id);
  m.rol = acceso_(ss, e, m, m.cmo, undefined, m.apr);
  if (!m.rol) throw new Error('No tiene permiso para ver este proceso.');
  m.datos = unir_(r.slice(REPO_NC));
  m.bloqueo = aprBloq_(m.apr);
  m.soloLectura = m.rol === 'lector' || m.rol === 'aprobador' || m.bloqueo;
  return m;
}

/**
 * Guarda el proceso. base = lo que el usuario tenía guardado antes de sus cambios: si alguien guardó entretanto,
 * los cambios se combinan (mezcla_) en vez de pisarse, y se devuelve en "datos" el resultado combinado.
 * El CMO invitado solo puede cambiar sus campos (cmoCampo_); el resto queda como está guardado.
 */
function repoGuardar(id, datos, version, forzar, base) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos');
  let S = JSON.parse(datos), out = datos;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    let f = fila_(sh, id), creado = now, archivado = false, ver = 0, creador = e;
    if (f) {
      const r = sh.getRange(f, 1, 1, sh.getLastColumn()).getValues()[0], m = meta_(r), rol = acceso_(ss, e, m);
      if (rol !== 'dueno' && rol !== 'cmo') throw new Error('Solo quien creó el proceso (o el CMO al que se le envió) puede modificarlo. Use "Duplicar" para trabajar una copia propia.');
      aprNoBloq_(ss, id);
      const cur = unir_(r.slice(REPO_NC));
      if (rol === 'cmo') S = mezcla_(JSON.parse(base || cur), S, JSON.parse(cur), [], cmoCampo_);
      else if (!forzar && m.version !== Number(version)) {
        if (!base) return { conflicto: true, modificado: m.modificado, modificadoPor: m.modificadoPor, version: m.version };
        S = mezcla_(JSON.parse(base), S, JSON.parse(cur), []);
      }
      if (rol === 'cmo' || (!forzar && m.version !== Number(version))) out = JSON.stringify(S);
      creado = r[REPO_COLS.indexOf('creado')];
      archivado = m.archivado;
      ver = m.version;
      creador = m.creador;
    } else {
      id = Utilities.getUuid();
      f = sh.getLastRow() + 1;
    }
    const chunks = trozos_(out);
    const width = Math.max(chunks.length, sh.getLastColumn() - REPO_NC, 1);
    while (chunks.length < width) chunks.push('');
    asegurarTam_(sh, f, REPO_NC + width);
    const prov = (S.prov || []).map(p => p && p.n).filter((x, i, a) => x && a.indexOf(x) === i).join(', ');
    const row = [id, txt_(S.cc), txt_(S.desc), txt_(S.gestor), txt_(prov), creador, creado, now, e, ver + 1, archivado].concat(chunks);
    sh.getRange(f, 1, 1, row.length).setValues([row]);
    const cache = CacheService.getScriptCache();
    cache.put('pv_' + id, String(ver + 1), 21600);
    if (!cache.get('v_' + id)) snap_(ss, id, e, ver + 1, S.cc, out);
    const res = { id: id, version: ver + 1, modificado: now.toISOString() };
    if (out !== datos) res.datos = out;
    return res;
  } finally {
    lock.releaseLock();
  }
}

function repoArchivar(id, archivar) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const m = meta_(sh.getRange(f, 1, 1, REPO_NC).getValues()[0]);
  if (m.creador !== e && !perfil_(ss, e).admin) throw new Error('No tiene permiso para archivar este proceso.');
  sh.getRange(f, REPO_COLS.indexOf('archivado') + 1).setValue(!!archivar);
  return true;
}

function repoVersiones(id) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const m = meta_(sh.getRange(f, 1, 1, REPO_NC).getValues()[0]);
  if (!acceso_(ss, e, m)) throw new Error('No tiene permiso para ver este proceso.');
  const vs = versiones_(ss), n = vs.getLastRow() - 1;
  if (n < 1) return [];
  return vs.getRange(2, 1, n, REPO_VNC).getValues()
    .filter(r => r[1] === id)
    .map(r => ({ idVersion: r[0], fecha: r[2] instanceof Date ? r[2].toISOString() : String(r[2]), usuario: r[3], version: Number(r[4]) || 0 }))
    .sort((a, b) => b.fecha.localeCompare(a.fecha));
}

function repoVersion(idVersion) {
  const ss = repo_(), e = yo_(), vs = versiones_(ss), fv = fila_(vs, idVersion);
  if (!fv) throw new Error('La versión ya no existe.');
  const r = vs.getRange(fv, 1, 1, vs.getLastColumn()).getValues()[0];
  const sh = ss.getSheetByName('Procesos'), f = fila_(sh, r[1]);
  if (!f) throw new Error('El proceso de esta versión ya no existe.');
  const m = meta_(sh.getRange(f, 1, 1, REPO_NC).getValues()[0]);
  if (!acceso_(ss, e, m)) throw new Error('No tiene permiso para ver este proceso.');
  return { datos: unir_(r.slice(REPO_VNC)), version: Number(r[4]) || 0 };
}

/** Guarda ya una versión del estado actual del proceso (antes de restaurar otra, o a pedido del creador). */
function repoInstantanea(id) {
  const ss = repo_(), e = yo_(), sh = ss.getSheetByName('Procesos'), f = fila_(sh, id);
  if (!f) throw new Error('El proceso ya no existe en el repositorio.');
  const r = sh.getRange(f, 1, 1, sh.getLastColumn()).getValues()[0], m = meta_(r);
  if (m.creador !== e) throw new Error('Solo quien creó el proceso puede guardar o restaurar versiones.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    snap_(ss, id, e, m.version, m.cc, unir_(r.slice(REPO_NC)));
  } finally {
    lock.releaseLock();
  }
  return true;
}
