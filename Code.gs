// ===== CONFIGURACIÓN (ya cargada con tus datos) =====
const SHEET_ID  = '14-TXgKnV27BXdNDVDw8OEWURHWQKBhBJyBfGElcQidk';
const FOLDER_ID = '12Zvkp43EJOosQdVVhXPHhcSQovAZIFG4';
// ====================================================

const sh = n => SpreadsheetApp.openById(SHEET_ID).getSheetByName(n);
const read = n => { const [h, ...v] = sh(n).getDataRange().getValues(); return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]]))); };
const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const low = s => String(s || '').trim().toLowerCase();
const b64 = b => Utilities.base64EncodeWebSafe(b);
const pub = p => ({
  id: String(p.id), titulo: p.titulo, cuerpo: p.cuerpo, categoria: p.categoria, link: p.link,
  fotos: String(p.fotos).split(' ').filter(Boolean), estado: p.estado, autor: p.autorNombre,
  autorUsuario: p.autorUsuario, fecha: new Date(p.fecha).toISOString(), motivo: p.motivoRechazo, likes: Number(p.likes) || 0
});
const byDate = (a, b) => new Date(b.fecha) - new Date(a.fecha);

// ---------- Contraseñas y sesiones ----------
function secret() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRETO');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRETO', s); }
  return s;
}
const sign = s => b64(Utilities.computeHmacSha256Signature(s, secret()));
function hash(clave, salt) {
  let h = salt + clave;
  for (let i = 0; i < 300; i++) h = b64(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
const mkClave = clave => { const salt = Utilities.getUuid().slice(0, 12); return salt + ':' + hash(clave, salt); };
function check(clave, stored) { const [salt, h] = String(stored).split(':'); return !!h && hash(String(clave), salt) === h; }
function mkTok(us) {
  const body = b64(Utilities.newBlob(JSON.stringify({ u: us, e: Date.now() + 8 * 36e5 })).getBytes());
  return body + '.' + sign(body);
}
function who(tok) {
  try {
    const [body, sig] = String(tok).split('.');
    if (sign(body) !== sig) return null;
    const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString());
    if (o.e < Date.now()) return null;
    const u = read('Usuarios').find(x => low(x.usuario) === o.u);
    return u ? { usuario: o.u, nombre: u.nombre, rol: u.rol } : null;
  } catch (err) { return null; }
}
function login(us, clave) {
  us = low(us);
  const c = CacheService.getScriptCache(), k = 'f_' + us.slice(0, 40), n = Number(c.get(k)) || 0;
  if (n >= 5) return { error: 'Demasiados intentos. Probá de nuevo en 10 minutos.' };
  const u = read('Usuarios').find(x => low(x.usuario) === us);
  if (!u || !u.clave || !check(clave, u.clave)) { c.put(k, String(n + 1), 600); return { error: 'Usuario o contraseña incorrectos.' }; }
  c.remove(k);
  return { token: mkTok(us), me: { usuario: us, nombre: u.nombre, rol: u.rol } };
}
function userRow(us) {
  const s = sh('Usuarios');
  if (s.getLastRow() < 2) return 0;
  const i = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().flat().map(low).indexOf(us);
  return i < 0 ? 0 : i + 2;
}
function setUser(us, nombre, rol, clave, email) {
  const s = sh('Usuarios'), r = userRow(us), row = [us, nombre, rol, mkClave(clave)];
  if (r) { s.getRange(r, 1, 1, 4).setValues([row]); if (email) s.getRange(r, 5).setValue(email); }
  else s.appendRow([...row, email || '']);
}
// Ejecutala UNA vez desde el editor para crear el primer usuario de Dirección.
// Después de correrla, borrá la contraseña de acá.
function crearPrimerUsuario() {
  setUser('monty', 'Alejandro', 'direccion', 'borrada', '');
}

// ---------- API ----------
function doGet() { // público: publicaciones aprobadas
  return json(read('Publicaciones').filter(p => p.estado === 'aprobada').map(pub).sort(byDate));
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.a === 'like') return json(like(d.id));
  if (d.a === 'login') return json(login(d.usuario, d.clave));
  const u = who(d.token);
  if (!u) return json({ error: 'auth' });
  if (d.a === 'me') return json({ me: u });

  if (d.a === 'pass') {
    const row = read('Usuarios').find(x => low(x.usuario) === u.usuario);
    if (!check(d.vieja, row.clave)) return json({ error: 'La contraseña actual no es correcta.' });
    if (String(d.nueva).length < 8) return json({ error: 'La contraseña nueva debe tener al menos 8 caracteres.' });
    sh('Usuarios').getRange(userRow(u.usuario), 4).setValue(mkClave(String(d.nueva)));
    return json({ ok: true });
  }
  if (!u.rol) return json({ error: 'Tu usuario no tiene rol asignado.' });

  if (d.a === 'user' && u.rol === 'direccion') {
    const us = low(d.usuario), cl = String(d.clave || '');
    if (!/^[a-z0-9._-]{3,30}$/.test(us) || cl.length < 8) return json({ error: 'Usuario: de 3 a 30 letras o números, sin espacios. Contraseña: mínimo 8 caracteres.' });
    if (['redaccion', 'direccion'].indexOf(d.rol) < 0) return json({ error: 'Rol no válido.' });
    setUser(us, String(d.nombre || us).slice(0, 60), d.rol, cl, String(d.email || '').trim());
    return json({ ok: true });
  }
  if (d.a === 'panel') {
    const all = read('Publicaciones').map(pub).sort(byDate);
    return json({ items: u.rol === 'direccion' ? all.filter(p => p.estado === 'pendiente') : all.filter(p => p.autorUsuario === u.usuario) });
  }
  if (d.a === 'create') {
    const titulo = String(d.titulo || '').trim().slice(0, 120), cuerpo = String(d.cuerpo || '').trim().slice(0, 2000);
    if (!titulo || !cuerpo) return json({ error: 'Completá el título y el texto.' });
    const link = /^https?:\/\//.test(d.link) ? d.link : '';
    const fotos = (d.fotos || []).slice(0, 5).map(saveImg).join(' ');
    const auto = u.rol === 'direccion';
    sh('Publicaciones').appendRow([Utilities.getUuid().slice(0, 8), titulo, cuerpo, d.categoria, link, fotos,
      auto ? 'aprobada' : 'pendiente', u.usuario, u.nombre, new Date().toISOString(), '', 0]);
    if (!auto) notify(titulo, u.nombre);
    return json({ ok: true, estado: auto ? 'aprobada' : 'pendiente' });
  }
  if (d.a === 'review' && u.rol === 'direccion') {
    const r = rowOf(d.id);
    if (!r) return json({ error: 'No se encontró la publicación.' });
    sh('Publicaciones').getRange(r, 7).setValue(d.ok ? 'aprobada' : 'rechazada');
    sh('Publicaciones').getRange(r, 11).setValue(d.ok ? '' : String(d.motivo || '').slice(0, 300));
    return json({ ok: true });
  }
  return json({ error: 'Acción no válida.' });
}

function rowOf(id) {
  const s = sh('Publicaciones');
  if (s.getLastRow() < 2) return 0;
  const i = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().flat().map(String).indexOf(String(id));
  return i < 0 ? 0 : i + 2;
}
function like(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    const r = rowOf(id); if (!r) return { error: 'no existe' };
    const c = sh('Publicaciones').getRange(r, 12);
    c.setValue((Number(c.getValue()) || 0) + 1);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function saveImg(dataUrl) {
  const blob = Utilities.newBlob(Utilities.base64Decode(dataUrl.split(',')[1]), 'image/jpeg', 'foto.jpg');
  const f = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + f.getId() + '=w1080';
}
function notify(titulo, autor) {
  try {
    const to = read('Usuarios').filter(u => u.rol === 'direccion' && u.email).map(u => u.email).join(',');
    if (to) MailApp.sendEmail(to, 'Nueva publicación para revisar', autor + ' envió "' + titulo + '" a la Cartelera. Entrá a la app, pestaña Revisar, para aprobarla o rechazarla.');
  } catch (err) {}
}
