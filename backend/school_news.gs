/**
 * «Ұлан» – «Біздің мектепте» рубрикасының бэкенді (Google Apps Script).
 *
 * Оқушы тіркеледі (логин + құпиясөз), мектебінің жаңалығын 70 сөзге дейін жазып,
 * бір сурет жүктейді. Жаңалық алдымен «Тексерілуде» күйінде тұрады,
 * редакция мақұлдағаннан кейін ғана сайтта көрінеді.
 *
 * Орнату:
 *  1) Осы кодты жаңа Apps Script жобасына қойыңыз.
 *  2) setup() функциясын бір рет іске қосыңыз (Google кесте мен Drive папкасын жасайды).
 *  3) Deploy → New deployment → Web app: Execute as «Me», Who has access «Anyone».
 *  4) Шыққан /exec сілтемесін сайттағы SCHOOL_URL-ге қойыңыз.
 *
 * Деректер: «Ұлан – Біздің мектепте (деректер)» кестесі (Users, Posts парақтары)
 * және «Ұлан – Біздің мектепте суреттері» папкасы сіздің Google Drive-та сақталады.
 * Құпиясөздер ашық сақталмайды – тұзды SHA-256 хэші ғана.
 */

// Редакторларды тексеру үшін негізгі бэкенд (checkAdmin әрекеті)
const MAIN_URL = 'https://script.google.com/macros/s/AKfycbzxsVDRrzou_ooGE5BlN88F84jFCVtX-Gc41cDuPRX-wTnGt_GvKzDJb8qB4spiaH-e/exec';

const MAX_WORDS = 70;
const MAX_CHARS = 700;
const MAX_PHOTO_BYTES = 3 * 1024 * 1024;
const TOKEN_DAYS = 60;
const MAX_PENDING_PER_USER = 3;
const MAX_POSTS_PER_DAY = 5;

const USER_COLS = ['login', 'salt', 'hash', 'name', 'school', 'grade', 'region', 'created', 'token', 'tokenExp', 'blocked'];
const POST_COLS = ['id', 'created', 'login', 'name', 'school', 'grade', 'region', 'text', 'photoId', 'status', 'moderator', 'moderatedAt'];

// ======================= Орнату =======================
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('Ұлан – Біздің мектепте (деректер)');
    const u = ss.getSheets()[0];
    u.setName('Users'); u.appendRow(USER_COLS); u.setFrozenRows(1);
    const p = ss.insertSheet('Posts');
    p.appendRow(POST_COLS); p.setFrozenRows(1);
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder('Ұлан – Біздің мектепте суреттері').getId());
  }
  Logger.log('Дайын. Кесте: ' + props.getProperty('SHEET_ID') + ', папка: ' + props.getProperty('FOLDER_ID'));
}

// ======================= Кіру нүктелері =======================
// GET (JSONP) – тек көпшілікке ашық тізім. Құпиясөздер URL-де жүрмеуі үшін қалғаны POST арқылы.
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    out = (p.action === 'list') ? listPublic() : { ok: false, error: 'Белгісіз әрекет' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return respond(out, p.callback);
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents || '{}'); } catch (x) {}
  let out;
  try {
    const routes = { register, login, me, submit, modList, modSet, modResetPass, list: listPublic };
    const fn = routes[p.action];
    out = fn ? fn(p) : { ok: false, error: 'Белгісіз әрекет' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return respond(out);
}

function respond(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb && /^[A-Za-z0-9_$.]{1,64}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ======================= Көмекшілер =======================
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')); }
function sheet_(name) { return ss_().getSheetByName(name); }
function rows_(sh) {
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.map(function (r, i) { const o = { _row: i + 2 }; head.forEach(function (h, j) { o[h] = r[j]; }); return o; });
}
function setCell_(sh, row, cols, col, val) { sh.getRange(row, cols.indexOf(col) + 1).setValue(val); }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function hash_(salt, pass) {
  let h = salt + '|' + pass;
  for (let i = 0; i < 300; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  return h;
}
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max); }
function words_(s) { return String(s || '').trim().split(/\s+/).filter(Boolean).length; }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function rateLimit_(key, max, seconds) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(key) || 0);
  if (n >= max) throw new Error('Тым жиі. Біраздан кейін қайталап көр.');
  c.put(key, String(n + 1), seconds);
}
function userByToken_(token) {
  if (!token) return null;
  const u = rows_(sheet_('Users')).filter(function (r) { return r.token && r.token === token; })[0];
  if (!u || u.blocked === true || new Date(u.tokenExp) < new Date()) return null;
  return u;
}
function publicUser_(u) { return { login: u.login, name: u.name, school: u.school, grade: u.grade, region: u.region }; }
function photoUrl_(id) { return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1200' : ''; }

// ======================= Оқушы =======================
function register(p) {
  rateLimit_('reg_all', 40, 3600);
  const login = clean_(p.login, 20).toLowerCase();
  const pass = String(p.pass || '');
  const name = clean_(p.name, 60), school = clean_(p.school, 100), grade = clean_(p.grade, 20), region = clean_(p.region, 60);
  if (!/^[a-z0-9_.]{3,20}$/.test(login)) return { ok: false, error: 'Логин 3–20 таңба болсын: латын әріптері, сандар, _ немесе .' };
  if (pass.length < 6 || pass.length > 64) return { ok: false, error: 'Құпиясөз кемінде 6 таңба болсын.' };
  if (name.length < 2) return { ok: false, error: 'Аты-жөніңді жаз.' };
  if (school.length < 2) return { ok: false, error: 'Мектебіңді жаз.' };
  if (!grade) return { ok: false, error: 'Сыныбыңды жаз.' };
  return withLock_(function () {
    const sh = sheet_('Users');
    if (rows_(sh).some(function (r) { return String(r.login) === login; })) return { ok: false, error: 'Бұл логин бос емес. Басқасын ойлап тап.' };
    const salt = Utilities.getUuid();
    const token = Utilities.getUuid();
    const exp = new Date(Date.now() + TOKEN_DAYS * 864e5);
    sh.appendRow([login, salt, hash_(salt, pass), name, school, grade, region, new Date(), token, exp, false]);
    return { ok: true, token: token, user: { login: login, name: name, school: school, grade: grade, region: region } };
  });
}

function login(p) {
  const login = clean_(p.login, 20).toLowerCase();
  rateLimit_('login_' + login, 10, 600);
  const sh = sheet_('Users');
  const u = rows_(sh).filter(function (r) { return String(r.login) === login; })[0];
  if (!u || u.blocked === true || hash_(u.salt, String(p.pass || '')) !== u.hash) return { ok: false, error: 'Логин не құпиясөз қате.' };
  const token = Utilities.getUuid();
  setCell_(sh, u._row, USER_COLS, 'token', token);
  setCell_(sh, u._row, USER_COLS, 'tokenExp', new Date(Date.now() + TOKEN_DAYS * 864e5));
  return { ok: true, token: token, user: publicUser_(u) };
}

function me(p) {
  const u = userByToken_(p.token);
  if (!u) return { ok: false, error: 'Қайта кіру керек.', relogin: true };
  const posts = rows_(sheet_('Posts'))
    .filter(function (r) { return r.login === u.login; })
    .sort(function (a, b) { return new Date(b.created) - new Date(a.created); })
    .slice(0, 20)
    .map(function (r) { return { id: r.id, date: r.created, text: r.text, status: r.status, photo: r.status === 'approved' ? photoUrl_(r.photoId) : '' }; });
  return { ok: true, user: publicUser_(u), posts: posts };
}

function submit(p) {
  const u = userByToken_(p.token);
  if (!u) return { ok: false, error: 'Қайта кіру керек.', relogin: true };
  const text = String(p.text || '').replace(/\r\n/g, '\n').replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, ' ').trim().slice(0, MAX_CHARS);
  const n = words_(text);
  if (n < 5) return { ok: false, error: 'Жаңалықты кемінде 5 сөзбен жаз.' };
  if (n > MAX_WORDS) return { ok: false, error: 'Мәтін ' + MAX_WORDS + ' сөзден аспауы керек. Қазір: ' + n + ' сөз.' };

  const posts = rows_(sheet_('Posts')).filter(function (r) { return r.login === u.login; });
  if (posts.filter(function (r) { return r.status === 'pending'; }).length >= MAX_PENDING_PER_USER)
    return { ok: false, error: 'Тексерілуді күтіп тұрған жаңалықтарың көп. Редакция қарағанша күте тұр.' };
  const day = Date.now() - 864e5;
  if (posts.filter(function (r) { return new Date(r.created).getTime() > day; }).length >= MAX_POSTS_PER_DAY)
    return { ok: false, error: 'Бір күнде ' + MAX_POSTS_PER_DAY + ' жаңалықтан артық жіберуге болмайды.' };

  let photoId = '';
  if (p.photo) {
    const m = String(p.photo).match(/^data:(image\/(jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
    if (!m) return { ok: false, error: 'Сурет форматы қолдау таппайды (JPG, PNG не WEBP болсын).' };
    const bytes = Utilities.base64Decode(m[3]);
    if (bytes.length > MAX_PHOTO_BYTES) return { ok: false, error: 'Сурет тым үлкен (3 МБ-тан аспасын).' };
    const id = 'S' + Date.now();
    const file = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'))
      .createFile(Utilities.newBlob(bytes, m[1], id + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2])));
    photoId = file.getId();   // жарияланғанша сурет жабық тұрады
  }
  return withLock_(function () {
    const id = 'S' + Date.now() + Math.floor(Math.random() * 1000);
    sheet_('Posts').appendRow([id, new Date(), u.login, u.name, u.school, u.grade, u.region, text, photoId, 'pending', '', '']);
    return { ok: true, id: id, status: 'pending' };
  });
}

// ======================= Сайттағы көпшілікке ашық тізім =======================
function listPublic() {
  const c = CacheService.getScriptCache();
  const hit = c.get('public_list');
  if (hit) return JSON.parse(hit);
  const posts = rows_(sheet_('Posts'))
    .filter(function (r) { return r.status === 'approved'; })
    .sort(function (a, b) { return new Date(b.moderatedAt || b.created) - new Date(a.moderatedAt || a.created); })
    .slice(0, 30)
    .map(function (r) {
      return { id: r.id, date: r.created, text: r.text, photo: photoUrl_(r.photoId), name: r.name, school: r.school, grade: r.grade, region: r.region };
    });
  const out = { ok: true, posts: posts };
  c.put('public_list', JSON.stringify(out), 60);
  return out;
}

// ======================= Редакция =======================
function isModerator_(login, pass) {
  if (!login || !pass) return false;
  const key = 'mod_' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, login + '|' + pass, Utilities.Charset.UTF_8));
  const c = CacheService.getScriptCache();
  if (c.get(key) === '1') return true;
  const url = MAIN_URL + '?action=checkAdmin&login=' + encodeURIComponent(login) + '&pass=' + encodeURIComponent(pass) + '&callback=cb';
  const txt = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getContentText();
  const m = txt.match(/cb\(([\s\S]*)\)\s*;?\s*$/);
  let ok = false;
  try { ok = !!(m && JSON.parse(m[1]).ok === true); } catch (x) {}
  if (ok) c.put(key, '1', 600);
  return ok;
}

function modList(p) {
  if (!isModerator_(p.modLogin, p.modPass)) return { ok: false, error: 'Редакция логині қате.' };
  const pending = rows_(sheet_('Posts'))
    .filter(function (r) { return r.status === 'pending'; })
    .sort(function (a, b) { return new Date(a.created) - new Date(b.created); })
    .slice(0, 30)
    .map(function (r) {
      let photo = '';
      if (r.photoId) {
        try { const b = DriveApp.getFileById(r.photoId).getBlob(); photo = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (x) {}
      }
      return { id: r.id, date: r.created, text: r.text, photo: photo, name: r.name, school: r.school, grade: r.grade, region: r.region, login: r.login };
    });
  return { ok: true, posts: pending };
}

function modSet(p) {
  if (!isModerator_(p.modLogin, p.modPass)) return { ok: false, error: 'Редакция логині қате.' };
  if (['approved', 'rejected'].indexOf(p.status) < 0) return { ok: false, error: 'Күй қате.' };
  return withLock_(function () {
    const sh = sheet_('Posts');
    const r = rows_(sh).filter(function (x) { return x.id === p.id; })[0];
    if (!r) return { ok: false, error: 'Жаңалық табылмады.' };
    if (r.photoId) {
      try {
        const f = DriveApp.getFileById(r.photoId);
        if (p.status === 'approved') f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        else f.setTrashed(true);
      } catch (x) {}
    }
    setCell_(sh, r._row, POST_COLS, 'status', p.status);
    setCell_(sh, r._row, POST_COLS, 'moderator', clean_(p.modLogin, 40));
    setCell_(sh, r._row, POST_COLS, 'moderatedAt', new Date());
    CacheService.getScriptCache().remove('public_list');
    return { ok: true };
  });
}

// Оқушы құпиясөзін ұмытса – редакция уақытша құпиясөз береді
function modResetPass(p) {
  if (!isModerator_(p.modLogin, p.modPass)) return { ok: false, error: 'Редакция логині қате.' };
  const login = clean_(p.login, 20).toLowerCase();
  return withLock_(function () {
    const sh = sheet_('Users');
    const u = rows_(sh).filter(function (r) { return String(r.login) === login; })[0];
    if (!u) return { ok: false, error: 'Мұндай логин жоқ.' };
    const temp = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const salt = Utilities.getUuid();
    setCell_(sh, u._row, USER_COLS, 'salt', salt);
    setCell_(sh, u._row, USER_COLS, 'hash', hash_(salt, temp));
    setCell_(sh, u._row, USER_COLS, 'token', '');
    return { ok: true, login: login, tempPass: temp };
  });
}
