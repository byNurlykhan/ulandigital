/**
 * «Ұлан» – «Біздің мектепте» рубрикасының бэкенді (Google Apps Script).
 *
 * Оқушы тіркеледі (логин + құпиясөз), мектебінің жаңалығын 70 сөзге дейін жазып,
 * бір сурет жүктейді. Жаңалық алдымен «Тексерілуде» күйінде тұрады.
 * Редакция оны сайттағы кабинеттен НЕМЕСЕ тікелей Google кестеден жариялайды:
 * «ЖАРИЯЛАУ» бағанында тізімнен «Жарияланды» не «Қабылданбады» таңдалады.
 *
 * Орнату (бір рет):
 *  1) Кодты Apps Script жобасына қойыңыз.
 *  2) setup() – кесте мен Drive папкасын жасайды (бұрыннан болса, тек безендіреді).
 *  3) Deploy → Web app: Execute as «Me», Who has access «Anyone».
 *  Ескі (ағылшынша) кестені жаңа түрге көшіру: migrate().
 *
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

// Кестедегі күйлер (қазақша) және олардың жүйелік кодтары
const ST = { pending: 'Тексерілуде', approved: 'Жарияланды', rejected: 'Қабылданбады' };
function stCode_(v) {
  v = String(v || '').trim();
  for (const k in ST) if (ST[k] === v || k === v) return k;
  return 'pending';
}

// Бағандар: [жүйелік кілт, кестедегі қазақша атауы]
const POST_COLS = [
  ['id', 'ID'], ['created', 'Жіберілді'], ['name', 'Аты-жөні'], ['school', 'Мектебі'], ['grade', 'Сыныбы'],
  ['region', 'Қала / аудан'], ['text', 'Мәтіні'], ['photoLink', 'Сурет'],
  ['status', 'ЖАРИЯЛАУ (тізімнен таңдаңыз)'], ['moderator', 'Тексерген'], ['moderatedAt', 'Тексерілген уақыты'],
  ['login', 'Логин'], ['photoId', 'Сурет ID (жүйелік)'], ['shared', 'Сурет күйі (жүйелік)']
];
const USER_COLS = [
  ['login', 'Логин'], ['name', 'Аты-жөні'], ['school', 'Мектебі'], ['grade', 'Сыныбы'], ['region', 'Қала / аудан'],
  ['created', 'Тіркелді'], ['blocked', 'Бұғатталған (иә / бос)'],
  ['salt', 'salt (жүйелік)'], ['hash', 'hash (жүйелік)'], ['token', 'token (жүйелік)'], ['tokenExp', 'Токен мерзімі (жүйелік)']
];
const KEYS = function (cols) { return cols.map(function (c) { return c[0]; }); };
const POSTS = 'Жаңалықтар', USERS = 'Оқушылар';

// ======================= Орнату =======================
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('Ұлан – Біздің мектепте (деректер)');
    const p = ss.getSheets()[0]; p.setName(POSTS);
    p.getRange(1, 1, 1, POST_COLS.length).setValues([POST_COLS.map(function (c) { return c[1]; })]);
    const u = ss.insertSheet(USERS);
    u.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS.map(function (c) { return c[1]; })]);
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder('Ұлан – Біздің мектепте суреттері').getId());
  }
  format_();
  Logger.log('Дайын. Кесте: https://docs.google.com/spreadsheets/d/' + props.getProperty('SHEET_ID'));
}

// Кестені безендіру: тақырып, ені, «ЖАРИЯЛАУ» тізімі мен түстері
function format_() {
  const ss = ss_();
  const p = ss.getSheetByName(POSTS), u = ss.getSheetByName(USERS);
  [[p, POST_COLS], [u, USER_COLS]].forEach(function (x) {
    const sh = x[0], cols = x[1];
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#1C1633').setFontColor('#FFFFFF').setWrap(true);
    cols.forEach(function (c, i) { if (/жүйелік/.test(c[1])) sh.getRange(1, i + 1).setBackground('#9B97AF'); });
  });
  const col = function (cols, key) { return KEYS(cols).indexOf(key) + 1; };
  p.setColumnWidth(col(POST_COLS, 'text'), 420);
  p.getRange(2, col(POST_COLS, 'text'), p.getMaxRows() - 1, 1).setWrap(true);
  p.setColumnWidth(col(POST_COLS, 'status'), 200);
  const stRange = p.getRange(2, col(POST_COLS, 'status'), p.getMaxRows() - 1, 1);
  stRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([ST.pending, ST.approved, ST.rejected], true).setAllowInvalid(false)
    .setHelpText('Жаңалықты сайтқа шығару үшін «Жарияланды» таңдаңыз.').build());
  stRange.setFontWeight('bold');
  const rules = [
    [ST.pending, '#FFF1D0', '#8A6A00'], [ST.approved, '#DDF7E7', '#11893F'], [ST.rejected, '#ECE8F5', '#5A5574']
  ].map(function (r) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setFontColor(r[2]).setRanges([stRange]).build();
  });
  p.setConditionalFormatRules(rules);
}

// Бұрынғы (ағылшынша бағанды) кестені жаңа түрге көшіреді, сынақ деректерін өшіреді
function migrate() {
  const ss = ss_();
  const isTest = function (login) { return /^test_ulan_check/.test(String(login || '')); };
  const oldP = ss.getSheetByName('Posts'), oldU = ss.getSheetByName('Users');
  const readOld = function (sh) {
    if (!sh) return [];
    const v = sh.getDataRange().getValues(); const head = v.shift();
    return v.map(function (r) { const o = {}; head.forEach(function (h, j) { o[h] = r[j]; }); return o; });
  };
  const posts = readOld(oldP), users = readOld(oldU);
  let removed = 0;
  posts.forEach(function (o) {
    if (isTest(o.login) && o.photoId) { try { DriveApp.getFileById(o.photoId).setTrashed(true); } catch (x) {} }
  });
  const p = ss.getSheetByName(POSTS) || ss.insertSheet(POSTS);
  const u = ss.getSheetByName(USERS) || ss.insertSheet(USERS);
  p.clear(); u.clear();
  p.getRange(1, 1, 1, POST_COLS.length).setValues([POST_COLS.map(function (c) { return c[1]; })]);
  u.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS.map(function (c) { return c[1]; })]);
  posts.forEach(function (o) {
    if (isTest(o.login)) { removed++; return; }
    o.status = ST[stCode_(o.status)];
    o.photoLink = o.photoId ? photoFormula_(o.photoId) : '';
    o.shared = '';
    p.appendRow(rowOf_(POST_COLS, o));
  });
  users.forEach(function (o) {
    if (isTest(o.login)) { removed++; return; }
    o.blocked = (o.blocked === true) ? 'иә' : '';
    u.appendRow(rowOf_(USER_COLS, o));
  });
  if (oldP) ss.deleteSheet(oldP);
  if (oldU) ss.deleteSheet(oldU);
  format_();
  CacheService.getScriptCache().remove('public_list');
  Logger.log('Көшірілді. Өшірілген сынақ жолдары: ' + removed);
}

// ======================= Кіру нүктелері =======================
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
function rows_(sh, cols) {
  const keys = KEYS(cols);
  const v = sh.getDataRange().getValues();
  v.shift();
  return v.map(function (r, i) { const o = { _row: i + 2 }; keys.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
// Формула енгізуден қорғау: = + - @ белгісінен басталған мәтіннің алдына апостроф қосылады (кестеде көрінбейді). Тек жүйенің өз формуласы (photoLink) өтеді.
const FORMULA_OK = { photoLink: true };
function safeCell_(key, v) { return (typeof v === 'string' && !FORMULA_OK[key] && /^\s*[=+\-@\t\r]/.test(v)) ? String.fromCharCode(39) + v : v; }
function rowOf_(cols, o) { return KEYS(cols).map(function (k) { return safeCell_(k, o[k] == null ? '' : o[k]); }); }
function setCell_(sh, row, cols, key, val) { sh.getRange(row, KEYS(cols).indexOf(key) + 1).setValue(safeCell_(key, val)); }
function photoFormula_(id) { return '=HYPERLINK("https://drive.google.com/file/d/' + id + '/view","Суретті ашу")'; }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function hash_(salt, pass) {
  let h = salt + '|' + pass;
  for (let i = 0; i < 300; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  return h;
}
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max); }
function words_(s) { return String(s || '').trim().split(/\s+/).filter(Boolean).length; }
function isBlocked_(u) { return u.blocked === true || /^(иә|ия|да|yes|true)$/i.test(String(u.blocked || '').trim()); }
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
  const u = rows_(sheet_(USERS), USER_COLS).filter(function (r) { return r.token && r.token === token; })[0];
  if (!u || isBlocked_(u) || new Date(u.tokenExp) < new Date()) return null;
  return u;
}
function publicUser_(u) { return { login: u.login, name: u.name, school: u.school, grade: u.grade, region: u.region }; }
function photoUrl_(id) { return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1200' : ''; }

// Сурет күйін жаңалық күйіне сәйкестендіру (кестеден өзгертілсе де)
function syncPhoto_(sh, r) {
  if (!r.photoId) return;
  const code = stCode_(r.status);
  try {
    if (code === 'approved' && r.shared !== 'ашық') {
      DriveApp.getFileById(r.photoId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      setCell_(sh, r._row, POST_COLS, 'shared', 'ашық'); r.shared = 'ашық';
    } else if (code === 'rejected' && r.shared !== 'өшірілді') {
      DriveApp.getFileById(r.photoId).setTrashed(true);
      setCell_(sh, r._row, POST_COLS, 'shared', 'өшірілді'); r.shared = 'өшірілді';
    }
  } catch (x) {}
}

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
    const sh = sheet_(USERS);
    if (rows_(sh, USER_COLS).some(function (r) { return String(r.login) === login; })) return { ok: false, error: 'Бұл логин бос емес. Басқасын ойлап тап.' };
    const salt = Utilities.getUuid();
    const token = Utilities.getUuid();
    sh.appendRow(rowOf_(USER_COLS, { login: login, name: name, school: school, grade: grade, region: region, created: new Date(),
      blocked: '', salt: salt, hash: hash_(salt, pass), token: token, tokenExp: new Date(Date.now() + TOKEN_DAYS * 864e5) }));
    return { ok: true, token: token, user: { login: login, name: name, school: school, grade: grade, region: region } };
  });
}

function login(p) {
  const login = clean_(p.login, 20).toLowerCase();
  rateLimit_('login_' + login, 10, 600);
  const sh = sheet_(USERS);
  const u = rows_(sh, USER_COLS).filter(function (r) { return String(r.login) === login; })[0];
  if (!u || isBlocked_(u) || hash_(u.salt, String(p.pass || '')) !== u.hash) return { ok: false, error: 'Логин не құпиясөз қате.' };
  const token = Utilities.getUuid();
  setCell_(sh, u._row, USER_COLS, 'token', token);
  setCell_(sh, u._row, USER_COLS, 'tokenExp', new Date(Date.now() + TOKEN_DAYS * 864e5));
  return { ok: true, token: token, user: publicUser_(u) };
}

function me(p) {
  const u = userByToken_(p.token);
  if (!u) return { ok: false, error: 'Қайта кіру керек.', relogin: true };
  const posts = rows_(sheet_(POSTS), POST_COLS)
    .filter(function (r) { return r.login === u.login; })
    .sort(function (a, b) { return new Date(b.created) - new Date(a.created); })
    .slice(0, 20)
    .map(function (r) { const c = stCode_(r.status); return { id: r.id, date: r.created, text: r.text, status: c, photo: c === 'approved' ? photoUrl_(r.photoId) : '' }; });
  return { ok: true, user: publicUser_(u), posts: posts };
}

function submit(p) {
  const u = userByToken_(p.token);
  if (!u) return { ok: false, error: 'Қайта кіру керек.', relogin: true };
  const text = String(p.text || '').replace(/\r\n/g, '\n').replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, ' ').trim().slice(0, MAX_CHARS);
  const n = words_(text);
  if (n < 5) return { ok: false, error: 'Жаңалықты кемінде 5 сөзбен жаз.' };
  if (n > MAX_WORDS) return { ok: false, error: 'Мәтін ' + MAX_WORDS + ' сөзден аспауы керек. Қазір: ' + n + ' сөз.' };

  const posts = rows_(sheet_(POSTS), POST_COLS).filter(function (r) { return r.login === u.login; });
  if (posts.filter(function (r) { return stCode_(r.status) === 'pending'; }).length >= MAX_PENDING_PER_USER)
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
    const file = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'))
      .createFile(Utilities.newBlob(bytes, m[1], 'S' + Date.now() + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2])));
    photoId = file.getId();   // жарияланғанша сурет жабық тұрады
  }
  return withLock_(function () {
    const id = 'S' + Date.now() + Math.floor(Math.random() * 1000);
    sheet_(POSTS).appendRow(rowOf_(POST_COLS, { id: id, created: new Date(), name: u.name, school: u.school, grade: u.grade,
      region: u.region, text: text, photoLink: photoId ? photoFormula_(photoId) : '', status: ST.pending,
      moderator: '', moderatedAt: '', login: u.login, photoId: photoId, shared: '' }));
    return { ok: true, id: id, status: 'pending' };
  });
}

// ======================= Сайттағы көпшілікке ашық тізім =======================
function listPublic() {
  const c = CacheService.getScriptCache();
  const hit = c.get('public_list');
  if (hit) return JSON.parse(hit);
  const sh = sheet_(POSTS);
  const all = rows_(sh, POST_COLS);
  all.forEach(function (r) { syncPhoto_(sh, r); });   // кестеден өзгертілген күйлерді де ескереміз
  const posts = all
    .filter(function (r) { return stCode_(r.status) === 'approved'; })
    .sort(function (a, b) { return new Date(b.created) - new Date(a.created); })
    .slice(0, 30)
    .map(function (r) {
      return { id: r.id, date: r.created, text: r.text, photo: photoUrl_(r.photoId), name: r.name, school: r.school, grade: r.grade, region: r.region };
    });
  const out = { ok: true, posts: posts };
  c.put('public_list', JSON.stringify(out), 60);
  return out;
}

// ======================= Редакция (сайттағы кабинет) =======================
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
  const pending = rows_(sheet_(POSTS), POST_COLS)
    .filter(function (r) { return stCode_(r.status) === 'pending'; })
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
    const sh = sheet_(POSTS);
    const r = rows_(sh, POST_COLS).filter(function (x) { return x.id === p.id; })[0];
    if (!r) return { ok: false, error: 'Жаңалық табылмады.' };
    setCell_(sh, r._row, POST_COLS, 'status', ST[p.status]);
    setCell_(sh, r._row, POST_COLS, 'moderator', clean_(p.modLogin, 40));
    setCell_(sh, r._row, POST_COLS, 'moderatedAt', new Date());
    r.status = ST[p.status];
    syncPhoto_(sh, r);
    CacheService.getScriptCache().remove('public_list');
    return { ok: true };
  });
}

// Оқушы құпиясөзін ұмытса – редакция уақытша құпиясөз береді
function modResetPass(p) {
  if (!isModerator_(p.modLogin, p.modPass)) return { ok: false, error: 'Редакция логині қате.' };
  const login = clean_(p.login, 20).toLowerCase();
  return withLock_(function () {
    const sh = sheet_(USERS);
    const u = rows_(sh, USER_COLS).filter(function (r) { return String(r.login) === login; })[0];
    if (!u) return { ok: false, error: 'Мұндай логин жоқ.' };
    const temp = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const salt = Utilities.getUuid();
    setCell_(sh, u._row, USER_COLS, 'salt', salt);
    setCell_(sh, u._row, USER_COLS, 'hash', hash_(salt, temp));
    setCell_(sh, u._row, USER_COLS, 'token', '');
    return { ok: true, login: login, tempPass: temp };
  });
}
