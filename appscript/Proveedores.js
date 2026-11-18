/**
 * Base de proveedores (solo lectura): "BD Proveedores - Rev 1", pestaña "Ventus".
 * El app se ejecuta con la cuenta del propietario, así que basta con que esa cuenta pueda ver el archivo.
 * Al cliente solo van nombre, NIT, texto de la columna de compliance (tal como se ve en la hoja), nombre de fantasía y evaluación.
 */
const PROV_SHEET_ID = '18CyghWnHXBulKpBsCUkgLO_Jd-ywtDv28dkQDTKHOSE';
const PROV_GID = 1715702941;
const PROV_CACHE = 'prov_v3'; // v3: cada fila trae también la evaluación
const PROV_TTL = 600; // segundos que se guarda la lista antes de volver a leer la hoja
const PROV_PART = 90000; // la caché admite 100 KB por valor; la lista se reparte en varias partes

// [[nombre, nit, compliance, fantasía, evaluación], ...] sin filas vacías; si un proveedor aparece varias veces queda el registro más reciente con compliance
function proveedoresCompliance() {
  const c = CacheService.getScriptCache();
  const n = +(c.get(PROV_CACHE + '_n') || 0);
  if (n) {
    const ks = [];
    for (let i = 0; i < n; i++) ks.push(PROV_CACHE + '_' + i);
    const g = c.getAll(ks);
    if (ks.every(k => g[k] != null)) return JSON.parse(ks.map(k => g[k]).join(''));
  }
  const L = provLeer_();
  const s = JSON.stringify(L), m = {};
  let i = 0;
  for (; i * PROV_PART < s.length; i++) m[PROV_CACHE + '_' + i] = s.slice(i * PROV_PART, (i + 1) * PROV_PART);
  m[PROV_CACHE + '_n'] = String(i);
  try { c.putAll(m, PROV_TTL); } catch (e) { /* sin caché: la próxima vez se vuelve a leer */ }
  return L;
}

function provLeer_() {
  const sh = SpreadsheetApp.openById(PROV_SHEET_ID).getSheets().find(s => s.getSheetId() === PROV_GID);
  if (!sh) throw new Error('No se encontró la pestaña de proveedores en "BD Proveedores".');
  const nr = sh.getLastRow() - 1, nc = sh.getLastColumn();
  if (nr < 1) return [];
  const H = sh.getRange(1, 1, 1, nc).getDisplayValues()[0].map(h => h.replace(/\s+/g, ' ').trim());
  const col = (rx, nom) => {
    const i = H.findIndex(h => rx.test(h));
    if (i < 0) throw new Error('No se encontró la columna "' + nom + '" en la base de proveedores.');
    return sh.getRange(2, i + 1, nr, 1).getDisplayValues().map(r => r[0].trim());
  };
  const N = col(/^Nombre de la Empresa/i, 'Nombre de la Empresa');
  const T = col(/ID Fiscal/i, 'Número ID Fiscal');
  const C = col(/compliance/i, 'Fecha realización compliance');
  const F = col(/^Nombre Fantas/i, 'Nombre Fantasía de la empresa');
  // la evaluación es opcional: si la hoja no tiene esa columna, el app la deja en blanco
  const ie = H.findIndex(h => /^evaluaci/i.test(h)), V = ie < 0 ? N.map(() => '') : sh.getRange(2, ie + 1, nr, 1).getDisplayValues().map(r => r[0].replace(/\s+/g, ' ').trim());
  const idx = {}, out = [];
  for (let r = 0; r < nr; r++) {
    const nom = N[r].replace(/\s+/g, ' ');
    if (!nom) continue;
    const k = nom.toUpperCase() + '|' + T[r].replace(/\D/g, '');
    const fan = F[r].replace(/\s+/g, ' ');
    const fila = [nom, T[r], C[r], fan.toUpperCase() === nom.toUpperCase() ? '' : fan, V[r]];
    if (k in idx) { if (C[r] || !out[idx[k]][2]) out[idx[k]] = fila; }
    else { idx[k] = out.length; out.push(fila); }
  }
  return out;
}
