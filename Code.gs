/**
 * ระบบธุรการในชั้นเรียน - วิทยาลัยเทคโนโลยีหลวงปู่สรวงวัดไพรพัฒนา
 * ฝั่งเซิร์ฟเวอร์ (Google Apps Script) เก็บข้อมูลใน Google Sheets
 * แผ่นงาน: Users (ครู/ผู้ดูแล), Rooms, Students, Docs (รายวิชา เช็กชื่อ คะแนน งานที่ส่ง)
 */
var DEFAULT_SHEET_URLS = [
  'https://docs.google.com/spreadsheets/d/1vKKyaKgyBSfTSGXv5uIHU7RL7TOuSmijnodyDBa8QYw/edit',
  'https://docs.google.com/spreadsheets/d/14qmiVVydOo3OAUY7r7XV77b4GUptoET6V7mQ4LPJJ70/edit'
];
var SHEETS = {
  Users: ['username', 'name', 'role', 'salt', 'hash'],
  Rooms: ['rid', 'code', 'major', 'level', 'advisor', 'order'],
  Students: ['id', 'first', 'last', 'rid', 'salt', 'hash'],
  Docs: ['coll', 'id', 'owner', 'sid', 'json', 'updated']
};
var SESSION_SECONDS = 21600; // 6 ชั่วโมง
var MAX_FILE_BYTES = 5 * 1024 * 1024;
var UPLOAD_FOLDER = 'งานที่นักศึกษาส่ง (ระบบธุรการในชั้นเรียน)';


function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'api') {
    var out = api(e.parameter.name || '', e.parameter.payload || '{}');
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบธุรการในชั้นเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function doPost(e) {
  var res;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = api(body.name || '', JSON.stringify(body.payload || {}));
  } catch (err) {
    res = JSON.stringify({ ok: false, error: 'คำขอไม่ถูกต้อง' });
  }
  return ContentService.createTextOutput(res).setMimeType(ContentService.MimeType.JSON);
}


function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('ระบบธุรการ')
      .addItem('ติดตั้งระบบ (ทำครั้งแรกครั้งเดียว)', 'setup')
      .addItem('ตั้งรหัสผ่าน admin ใหม่', 'resetAdminPassword')
      .addToUi();
  } catch (e) {
    /* กรณีรันจากสคริปต์แยกเดี่ยว (Standalone Script) จะไม่มี UI ของชีต */
  }
}


/* ---------- ไฟล์ Google Sheets ที่เก็บข้อมูล ---------- */
var SS_ = null;
function ss_() {
  if (SS_) return SS_;
  try { SS_ = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { SS_ = null; }
  if (SS_) return SS_;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) {
    try { return (SS_ = SpreadsheetApp.openById(id)); } catch (e) { SS_ = null; }
  }
  for (var i = 0; i < DEFAULT_SHEET_URLS.length; i++) {
    try {
      SS_ = SpreadsheetApp.openByUrl(DEFAULT_SHEET_URLS[i]);
      props.setProperty('SHEET_ID', SS_.getId());
      return SS_;
    } catch (e) {
      SS_ = null;
    }
  }
  SS_ = SpreadsheetApp.create('ระบบธุรการในชั้นเรียน (ฐานข้อมูล)');
  props.setProperty('SHEET_ID', SS_.getId());
  return SS_;
}


/* ---------- ติดตั้ง ---------- */
function setup() {
  var ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
    }
  });
  var msg = [];
  var hasAdmin = readRows_('Users').some(function (u) { return u.role === 'admin'; });
  if (!hasAdmin) {
    var pw = randomPassword_();
    var salt = Utilities.getUuid();
    var ush = ss.getSheetByName('Users');
    ush.appendRow(['admin', 'ผู้ดูแลระบบ', 'admin', salt, hash_(salt, pw)]);
    msg.push('สร้างผู้ดูแลระบบแล้ว\nชื่อผู้ใช้: admin\nรหัสผ่าน: ' + pw + '\n(จดไว้ แล้วเปลี่ยนได้หลังเข้าสู่ระบบ)');
  }
  if (ss.getSheetByName('Rooms').getLastRow() < 2) {
    msg.push('ขั้นต่อไป: เข้าระบบด้วย admin แล้วกด "นำเข้ารายชื่อนักศึกษา" เลือกไฟล์ใบรายชื่อ Excel');
  }
  if (!msg.length) msg.push('ระบบติดตั้งไว้แล้ว ไม่มีอะไรต้องทำเพิ่ม');
  msg.push('ไฟล์ข้อมูล: ' + ss.getUrl());
  notify_(msg.join('\n\n'));
}


function resetAdminPassword() {
  var pw = randomPassword_(), salt = Utilities.getUuid();
  var sh = ss_().getSheetByName('Users');
  if (!sh || sh.getLastRow() < 2) {
    notify_('ยังไม่มีผู้ดูแลระบบ ให้กด "ติดตั้งระบบ" ก่อน');
    return;
  }
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === 'admin') {
      sh.getRange(i + 1, 4, 1, 2).setValues([[salt, hash_(salt, pw)]]);
      notify_('รหัสผ่านใหม่ของ ' + rows[i][0] + ' คือ: ' + pw);
      return;
    }
  }
  notify_('ยังไม่มีผู้ดูแลระบบ ให้กด "ติดตั้งระบบ" ก่อน');
}


function notify_(text) {
  Logger.log(text);
  try { SpreadsheetApp.getUi().alert(text); } catch (e) { /* เรียกจากตัวแก้ไขสคริปต์ ดูผลใน Execution log */ }
}


/* ---------- เครื่องมือพื้นฐาน ---------- */
function sheet_(name) { return ss_().getSheetByName(name); }
function readRows_(name) {
  var sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), cols = SHEETS[name], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var c = 0; c < cols.length; c++) o[cols[c]] = v[i][c];
    out.push(o);
  }
  return out;
}
function hash_(salt, pass) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function randomPassword_() {
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789', s = '';
  var seed = Utilities.getUuid().replace(/-/g, '');
  for (var i = 0; i < 8; i++) s += chars.charAt(parseInt(seed.substr(i * 2, 2), 16) % chars.length);
  return s;
}
function fail_(msg) { var e = new Error(msg); e.userMessage = msg; throw e; }
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}


/* ---------- Docs: ที่เก็บข้อมูลแบบเอกสาร ---------- */
function allDocs_() {
  return readRows_('Docs').map(function (d) {
    return { _row: d._row, coll: String(d.coll), id: String(d.id), owner: String(d.owner), sid: String(d.sid), json: d.json };
  });
}
function parse_(d) { try { return JSON.parse(d.json); } catch (e) { return {}; } }
function putDoc_(coll, id, owner, sid, data, existingDocs) {
  var json = JSON.stringify(data);
  if (json.length > 45000) fail_('ข้อมูลรายการนี้ใหญ่เกินไป');
  var sh = sheet_('Docs'), row = [coll, id, owner, sid, json, new Date().toISOString()];
  var list = existingDocs || allDocs_();
  var found = list.filter(function (d) { return d.coll === coll && d.id === id; })[0];
  if (found) sh.getRange(found._row, 1, 1, 6).setValues([row]); else sh.appendRow(row);
}
function deleteDocs_(pred) {
  var sh = sheet_('Docs');
  allDocs_().filter(pred).map(function (d) { return d._row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (r) { sh.deleteRow(r); });
}


/* ---------- เข้าสู่ระบบ ---------- */
function session_(token) {
  if (!token) fail_('SESSION');
  var cache = CacheService.getScriptCache(), raw = cache.get('sess_' + token);
  if (!raw) fail_('SESSION');
  cache.put('sess_' + token, raw, SESSION_SECONDS);
  return JSON.parse(raw);
}
function need_(sess, roles) { if (roles.indexOf(sess.role) < 0) fail_('ไม่มีสิทธิ์ใช้งานส่วนนี้'); }
function checkTries_(key) {
  var cache = CacheService.getScriptCache(), n = Number(cache.get('try_' + key) || 0);
  if (n >= 8) fail_('ใส่รหัสผ่านผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่');
  return function (ok) { if (ok) cache.remove('try_' + key); else cache.put('try_' + key, String(n + 1), 600); };
}
function openSession_(sess) {
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(sess), SESSION_SECONDS);
  return token;
}


var API = {};


API.login = function (p) {
  var user = String(p.user || '').trim(), pass = String(p.pass || '');
  if (!user || !pass) fail_('กรอกข้อมูลให้ครบ');
  var done = checkTries_(p.kind + '_' + user.toLowerCase());
  if (p.kind === 'student') {
    var st = readRows_('Students').filter(function (s) { return String(s.id).trim() === user; })[0];
    var stuId = st ? String(st.id).trim() : '';
    var first = st && !st.hash;
    var ok = st && (first ? pass === stuId : hash_(st.salt, pass) === st.hash);
    done(!!ok);
    if (!ok) fail_('รหัสนักศึกษาหรือรหัสผ่านไม่ถูกต้อง');
    var name = String(st.first || '').trim() + ' ' + String(st.last || '').trim();
    return { token: openSession_({ role: 'student', user: stuId, name: name, rid: String(st.rid).trim() }), role: 'student', name: name, mustChange: first };
  }
  var u = readRows_('Users').filter(function (x) { return String(x.username).trim().toLowerCase() === user.toLowerCase(); })[0];
  var good = u && hash_(u.salt, pass) === u.hash;
  done(!!good);
  if (!good) fail_('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  return { token: openSession_({ role: u.role, user: String(u.username).trim(), name: u.name }), role: u.role, name: u.name, mustChange: false };
};


API.logout = function (p) { if (p && p.token) CacheService.getScriptCache().remove('sess_' + p.token); return {}; };


API.changePassword = function (p) {
  var s = session_(p.token), np = String(p.newPass || '');
  if (np.length < 6) fail_('รหัสผ่านใหม่ต้องยาวอย่างน้อย 6 ตัว');
  return withLock_(function () {
    var isStu = s.role === 'student', name = isStu ? 'Students' : 'Users';
    var rec = readRows_(name).filter(function (r) { return String(isStu ? r.id : r.username).trim() === s.user; })[0];
    if (!rec) fail_('ไม่พบบัญชี');
    var recId = String(isStu ? rec.id : rec.username).trim();
    var oldOk = isStu && !rec.hash ? String(p.oldPass) === recId : hash_(rec.salt, String(p.oldPass || '')) === rec.hash;
    if (!oldOk) fail_('รหัสผ่านเดิมไม่ถูกต้อง');
    if (isStu && np === recId) fail_('รหัสผ่านใหม่ต้องไม่ใช่รหัสนักศึกษา');
    var salt = Utilities.getUuid();
    sheet_(name).getRange(rec._row, isStu ? 5 : 4, 1, 2).setValues([[salt, hash_(salt, np)]]);
    return {};
  });
};


/* ---------- เริ่มต้นตามบทบาท ---------- */
function roomsWithStudents_() {
  var by = {};
  var rooms = readRows_('Rooms').map(function (r) {
    var o = { _id: String(r.rid).trim(), code: String(r.code || '').trim(), major: String(r.major || '').trim(), level: String(r.level || '').trim(), advisor: String(r.advisor || '').trim(), order: Number(r.order) || 0, students: [] };
    by[o._id] = o; return o;
  });
  readRows_('Students').forEach(function (s) {
    var rid = String(s.rid).trim();
    if (by[rid]) by[rid].students.push([String(s.id).trim(), String(s.first || '').trim(), String(s.last || '').trim()]);
  });
  return rooms;
}
function today_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'); }


API.boot = function (p) {
  var s = session_(p.token), docs = allDocs_();
  if (s.role === 'teacher') {
    var mine = docs.filter(function (d) { return d.coll === 'subjects' && d.owner === s.user; });
    var ids = {}; mine.forEach(function (d) { ids[d.id] = 1; });
    var t = today_();
    return {
      role: s.role, name: s.name, today: t, rooms: roomsWithStudents_(),
      subjects: mine.map(function (d) { var o = parse_(d); o._id = d.id; return o; }),
      todaySids: docs.filter(function (d) { return d.coll === 'att' && ids[d.sid] && d.id === d.sid + '_' + t; }).map(function (d) { return d.sid; })
    };
  }
  if (s.role === 'admin') {
    return {
      role: s.role, name: s.name, today: today_(), rooms: roomsWithStudents_(),
      users: readRows_('Users').map(function (u) {
        var uname = String(u.username).trim();
        return { username: uname, name: u.name, role: u.role,
          subjects: docs.filter(function (d) { return d.coll === 'subjects' && d.owner === uname; }).length };
      })
    };
  }
  // นักศึกษา: เห็นเฉพาะข้อมูลของตัวเอง
  var room = readRows_('Rooms').filter(function (r) { return String(r.rid).trim() === s.rid; })[0] || {};
  var teachers = {}; readRows_('Users').forEach(function (u) { teachers[String(u.username).trim()] = u.name; });
  var subjects = docs.filter(function (d) { return d.coll === 'subjects'; }).map(function (d) { var o = parse_(d); o._id = d.id; o._owner = d.owner; return o; })
    .filter(function (o) { return (o.rooms || []).indexOf(s.rid) >= 0; })
    .map(function (o) {
      var out = { _id: o._id, name: o.name, code: o.code, sched: o.sched, color: o.color, teacher: teachers[o._owner] || '', weights: o.weights, tasks: o.tasks || [], att: {}, subs: {}, work: {}, mid: null, fin: null };
      docs.forEach(function (d) {
        if (d.sid !== o._id) return;
        var j;
        if (d.coll === 'att') { j = parse_(d); out.att[j.date] = (j.m || {})[s.user] || 'P'; }
        else if (d.coll === 'subs') { j = parse_(d); if ((j.m || {})[s.user]) out.subs[j.tid] = j.m[s.user]; }
        else if (d.coll === 'exams') { j = parse_(d); if ((j.mid || {})[s.user] != null) out.mid = j.mid[s.user]; if ((j.fin || {})[s.user] != null) out.fin = j.fin[s.user]; }
        else if (d.coll === 'work' && d.owner === s.user) { j = parse_(d); out.work[j.tid] = j; }
      });
      return out;
    });
  return { role: s.role, name: s.name, today: today_(), me: { id: s.user, name: s.name, room: room.code || '', major: room.major || '', advisor: room.advisor || '' }, subjects: subjects };
};


/* ---------- ครู ---------- */
function ownSubject_(sess, sid, docs) {
  var d = (docs || allDocs_()).filter(function (x) { return x.coll === 'subjects' && x.id === sid; })[0];
  if (!d || d.owner !== sess.user) fail_('ไม่พบรายวิชานี้ในบัญชีของคุณ');
  return d;
}


API.openSubject = function (p) {
  var s = session_(p.token); need_(s, ['teacher']);
  var docs = allDocs_(); ownSubject_(s, p.sid, docs);
  var out = { att: [], subs: [], exams: null, work: [] };
  docs.forEach(function (d) {
    if (d.sid !== p.sid) return;
    var j = parse_(d);
    if (d.coll === 'att') out.att.push(j);
    else if (d.coll === 'subs') out.subs.push(j);
    else if (d.coll === 'exams') out.exams = j;
    else if (d.coll === 'work') out.work.push(j);
  });
  return out;
};


API.save = function (p) {
  var s = session_(p.token); need_(s, ['teacher']);
  var coll = p.coll, id = String(p.id), data = p.data || {};
  if (['subjects', 'att', 'subs', 'exams'].indexOf(coll) < 0 || !/^[A-Za-z0-9_\-]{1,80}$/.test(id)) fail_('คำขอไม่ถูกต้อง');
  return withLock_(function () {
    var docs = allDocs_();
    if (coll === 'subjects') {
      var ex = docs.filter(function (x) { return x.coll === 'subjects' && x.id === id; })[0];
      if (ex && ex.owner !== s.user) fail_('ไม่พบรายวิชานี้ในบัญชีของคุณ');
      if (!String(data.name || '').trim() || !(data.rooms || []).length) fail_('รายวิชาต้องมีชื่อและห้องเรียน');
      data.owner = s.user;
      putDoc_('subjects', id, s.user, id, data, docs);
    } else {
      var sid = String(data.sid || '');
      ownSubject_(s, sid, docs);
      if (id.indexOf(sid) !== 0) fail_('คำขอไม่ถูกต้อง');
      putDoc_(coll, id, s.user, sid, data, docs);
    }
    return {};
  });
};


API.remove = function (p) {
  var s = session_(p.token); need_(s, ['teacher']);
  return withLock_(function () {
    var docs = allDocs_();
    if (p.coll === 'subjects') {
      ownSubject_(s, p.id, docs);
      deleteDocs_(function (d) { return d.sid === p.id; });
    } else if (p.coll === 'subs' || p.coll === 'att') {
      var sid = String(p.sid || (p.id || '').split('_')[0] || '');
      var d = docs.filter(function (x) { return x.coll === p.coll && x.id === p.id; })[0];
      var targetSid = d ? d.sid : sid;
      if (targetSid) {
        ownSubject_(s, targetSid, docs);
        deleteDocs_(function (x) {
          if (x.coll === p.coll && x.id === p.id) return true;
          if (p.coll === 'subs' && x.coll === 'work' && x.sid === targetSid && x.id.indexOf(p.id + '_') === 0) return true;
          return false;
        });
      }
    } else fail_('คำขอไม่ถูกต้อง');
    return {};
  });
};


/* ---------- นักศึกษาส่งงาน ---------- */
API.submit = function (p) {
  var s = session_(p.token); need_(s, ['student']);
  var text = String(p.text || '').slice(0, 5000), link = String(p.link || '').trim().slice(0, 500);
  if (link && !/^https?:\/\//i.test(link)) fail_('ลิงก์ต้องขึ้นต้นด้วย http:// หรือ https://');
  if (!text.trim() && !link && !p.fileB64) fail_('พิมพ์คำตอบ แนบลิงก์ หรือแนบไฟล์อย่างน้อยหนึ่งอย่าง');
  var file = null;
  if (p.fileB64) {
    var bytes = Utilities.base64Decode(p.fileB64);
    if (bytes.length > MAX_FILE_BYTES) fail_('ไฟล์ใหญ่เกิน 5 MB');
    var it = DriveApp.getFoldersByName(UPLOAD_FOLDER);
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER);
    var safe = String(p.fileName || 'file').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80);
    var f = folder.createFile(Utilities.newBlob(bytes, p.mime || 'application/octet-stream', s.user + '_' + safe));
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* องค์กรอาจไม่อนุญาตให้แชร์ภายนอก */ }
    file = { name: safe, url: f.getUrl() };
  }
  return withLock_(function () {
    var docs = allDocs_();
    var subj = docs.filter(function (d) { return d.coll === 'subjects' && d.id === p.sid; })[0];
    var so = subj && parse_(subj);
    if (!so || (so.rooms || []).indexOf(s.rid) < 0) fail_('ไม่พบรายวิชานี้');
    var task = (so.tasks || []).filter(function (t) { return t.id === p.tid; })[0];
    if (!task) fail_('ไม่พบงานนี้');
    var id = p.sid + '_' + p.tid + '_' + s.user;
    var prev = docs.filter(function (d) { return d.coll === 'work' && d.id === id; })[0];
    var isLate = !!(task.due && today_() > task.due);
    var data = { sid: p.sid, tid: p.tid, stu: s.user, text: text, link: link, file: file || (prev ? parse_(prev).file : null) || null, at: new Date().toISOString(), late: isLate };
    putDoc_('work', id, s.user, p.sid, data, docs);

    // หากครูเคยกด "ยังไม่ส่ง" (rej: true) ไว้ก่อนหน้า เมื่อส่งงานใหม่ให้อัปเดตสถานะใน subs ด้วย
    var subsId = p.sid + '_' + p.tid;
    var subsDoc = docs.filter(function (d) { return d.coll === 'subs' && d.id === subsId; })[0];
    if (subsDoc) {
      var sj = parse_(subsDoc);
      if (sj && sj.m && sj.m[s.user] && (sj.m[s.user].rej || sj.m[s.user].st === 'none')) {
        delete sj.m[s.user].rej;
        sj.m[s.user].st = isLate ? 'late' : 'sent';
        putDoc_('subs', subsId, subsDoc.owner, p.sid, sj, docs);
      }
    }
    return data;
  });
};


/* ---------- ผู้ดูแลระบบ ---------- */
API.saveUser = function (p) {
  var s = session_(p.token); need_(s, ['admin']);
  var username = String(p.username || '').trim().toLowerCase(), name = String(p.name || '').trim();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) fail_('ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 . _ - ยาว 3-30 ตัว');
  if (!name) fail_('กรอกชื่อ-สกุลของครู');
  return withLock_(function () {
    if (readRows_('Users').some(function (u) { return String(u.username).trim().toLowerCase() === username; })) fail_('มีชื่อผู้ใช้นี้แล้ว');
    var pw = String(p.password || '') || randomPassword_();
    if (pw.length < 6) fail_('รหัสผ่านต้องยาวอย่างน้อย 6 ตัว');
    var salt = Utilities.getUuid();
    var sh = sheet_('Users');
    sh.appendRow([username, name, 'teacher', salt, hash_(salt, pw)]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('@');
    return { username: username, password: pw };
  });
};


API.resetUser = function (p) {
  var s = session_(p.token); need_(s, ['admin']);
  return withLock_(function () {
    var u = readRows_('Users').filter(function (x) { return String(x.username).trim() === String(p.username).trim(); })[0];
    if (!u) fail_('ไม่พบผู้ใช้');
    var pw = randomPassword_(), salt = Utilities.getUuid();
    sheet_('Users').getRange(u._row, 4, 1, 2).setValues([[salt, hash_(salt, pw)]]);
    return { username: String(u.username).trim(), password: pw };
  });
};


API.deleteUser = function (p) {
  var s = session_(p.token); need_(s, ['admin']);
  return withLock_(function () {
    var u = readRows_('Users').filter(function (x) { return String(x.username).trim() === String(p.username).trim(); })[0];
    if (!u) fail_('ไม่พบผู้ใช้');
    if (u.role === 'admin') fail_('ลบผู้ดูแลระบบไม่ได้');
    var uname = String(u.username).trim();
    if (allDocs_().some(function (d) { return d.coll === 'subjects' && d.owner === uname; })) fail_('ครูท่านนี้ยังมีรายวิชาอยู่ ให้ครูลบรายวิชาก่อน');
    sheet_('Users').deleteRow(u._row);
    return {};
  });
};


API.importRoster = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var rooms = p.rooms || [];
  if (!rooms.length) fail_('ไม่พบรายชื่อในไฟล์');
  return withLock_(function () {
    var old = {};
    readRows_('Students').forEach(function (x) { old[String(x.id).trim()] = [x.salt, x.hash]; });
    var oldRooms = readRows_('Rooms'), byCode = {}, maxN = 0;
    oldRooms.forEach(function (r) {
      byCode[String(r.code).trim()] = String(r.rid).trim();
      var n = Number(String(r.rid).replace(/\D/g, '')) || 0;
      if (n > maxN) maxN = n;
    });
    var keep = {}, roomRows = [], stuRows = [], seen = {};
    rooms.forEach(function (r, i) {
      var code = String(r.code || '').trim().slice(0, 40);
      if (!code || !(r.students || []).length) return;
      var nextN = byCode[code] ? 0 : (++maxN);
      var rid = byCode[code] || ('r' + (nextN < 10 ? '0' + nextN : String(nextN)));
      byCode[code] = rid;
      keep[rid] = 1;
      roomRows.push([rid, code, String(r.major || '').slice(0, 80), String(r.level || '').slice(0, 20), String(r.advisor || '').slice(0, 120), i + 1]);
      r.students.forEach(function (x) {
        var id = String(x[0] || '').replace(/\D/g, '');
        if (!id || seen[id]) return;
        seen[id] = 1;
        var pw = old[id] || ['', ''];
        stuRows.push([id, String(x[1] || '').trim().slice(0, 80), String(x[2] || '').trim().slice(0, 80), rid, pw[0], pw[1]]);
      });
    });
    oldRooms.forEach(function (r) {
      var rid = String(r.rid).trim();
      if (!keep[rid]) roomRows.push([rid, r.code, r.major, r.level, r.advisor, 9999]);
    });
    var rs = sheet_('Rooms'), st = sheet_('Students');
    if (rs.getLastRow() > 1) rs.getRange(2, 1, rs.getLastRow() - 1, 6).clearContent();
    if (st.getLastRow() > 1) st.getRange(2, 1, st.getLastRow() - 1, 6).clearContent();
    if (roomRows.length) rs.getRange(2, 1, roomRows.length, 6).setValues(roomRows);
    if (stuRows.length) {
      st.getRange(2, 1, stuRows.length, 1).setNumberFormat('@');
      st.getRange(2, 1, stuRows.length, 6).setValues(stuRows);
    }
    return { rooms: Object.keys(keep).length, students: stuRows.length, list: roomsWithStudents_() };
  });
};


API.saveRoom = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var code = String(p.code || '').trim().slice(0, 40);
  var major = String(p.major || '').trim().slice(0, 80);
  var level = String(p.level || '').trim().slice(0, 20);
  var advisor = String(p.advisor || '').trim().slice(0, 120);
  if (!code) fail_('กรุณากรอกรหัสห้องเรียน เช่น ชย.11');
  return withLock_(function () {
    var rows = readRows_('Rooms');
    var dup = rows.filter(function (r) { return String(r.code).trim().toLowerCase() === code.toLowerCase() && String(r.rid).trim() !== String(p.rid || '').trim(); })[0];
    if (dup) fail_('มีรหัสห้องเรียน "' + code + '" อยู่ในระบบแล้ว');
    var sh = sheet_('Rooms');
    var rid = String(p.rid || '').trim();
    if (rid) {
      var ex = rows.filter(function (r) { return String(r.rid).trim() === rid; })[0];
      if (!ex) fail_('ไม่พบห้องเรียนนี้');
      sh.getRange(ex._row, 2, 1, 4).setValues([[code, major, level, advisor]]);
    } else {
      var maxN = 0;
      rows.forEach(function (r) {
        var n = Number(String(r.rid).replace(/\D/g, '')) || 0;
        if (n > maxN) maxN = n;
      });
      var nextN = maxN + 1;
      rid = 'r' + (nextN < 10 ? '0' + nextN : String(nextN));
      sh.appendRow([rid, code, major, level, advisor, rows.length + 1]);
    }
    return { rid: rid, list: roomsWithStudents_() };
  });
};


API.deleteRoom = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var rid = String(p.rid || '').trim();
  if (!rid) fail_('ไม่พบรหัสห้องเรียน');
  return withLock_(function () {
    var rs = sheet_('Rooms');
    var rm = readRows_('Rooms').filter(function (r) { return String(r.rid).trim() === rid; })[0];
    if (!rm) fail_('ไม่พบห้องเรียนนี้');
    rs.deleteRow(rm._row);
    var st = sheet_('Students');
    readRows_('Students').filter(function (x) { return String(x.rid).trim() === rid; })
      .map(function (x) { return x._row; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (r) { st.deleteRow(r); });
    return { list: roomsWithStudents_() };
  });
};


API.saveStudent = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var origId = String(p.origId || '').trim();
  var id = String(p.id || '').trim();
  var first = String(p.first || '').trim().slice(0, 80);
  var last = String(p.last || '').trim().slice(0, 80);
  var rid = String(p.rid || '').trim();
  if (!id || !first || !rid) fail_('กรุณากรอกรหัสนักศึกษา ชื่อ และห้องเรียนให้ครบ');
  return withLock_(function () {
    var rm = readRows_('Rooms').filter(function (r) { return String(r.rid).trim() === rid; })[0];
    if (!rm) fail_('ไม่พบห้องเรียนที่เลือก');
    var allStu = readRows_('Students');
    var sh = sheet_('Students');
    if (origId) {
      var cur = allStu.filter(function (x) { return String(x.id).trim() === origId; })[0];
      if (!cur) fail_('ไม่พบข้อมูลนักศึกษาเดิม');
      if (id !== origId && allStu.some(function (x) { return String(x.id).trim() === id; })) {
        fail_('รหัสนักศึกษา ' + id + ' มีอยู่ในระบบแล้ว');
      }
      sh.getRange(cur._row, 1).setNumberFormat('@');
      sh.getRange(cur._row, 1, 1, 4).setValues([[id, first, last, rid]]);
    } else {
      if (allStu.some(function (x) { return String(x.id).trim() === id; })) {
        fail_('รหัสนักศึกษา ' + id + ' มีอยู่ในระบบแล้ว');
      }
      sh.appendRow([id, first, last, rid, '', '']);
      sh.getRange(sh.getLastRow(), 1).setNumberFormat('@');
    }
    return { list: roomsWithStudents_() };
  });
};


API.bulkAddStudents = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var rid = String(p.rid || '').trim();
  var items = p.students || [];
  if (!rid || !items.length) fail_('กรุณาระบุห้องเรียนและรายชื่อนักศึกษา');
  return withLock_(function () {
    var rm = readRows_('Rooms').filter(function (r) { return String(r.rid).trim() === rid; })[0];
    if (!rm) fail_('ไม่พบห้องเรียนที่เลือก');
    var allStu = readRows_('Students');
    var byId = {};
    allStu.forEach(function (x) { byId[String(x.id).trim()] = x; });
    var sh = sheet_('Students');
    var added = 0, updated = 0;
    items.forEach(function (item) {
      var id = String(item[0] || '').trim();
      var first = String(item[1] || '').trim().slice(0, 80);
      var last = String(item[2] || '').trim().slice(0, 80);
      if (!id || !first) return;
      if (byId[id]) {
        sh.getRange(byId[id]._row, 2, 1, 3).setValues([[first, last, rid]]);
        updated++;
      } else {
        sh.appendRow([id, first, last, rid, '', '']);
        sh.getRange(sh.getLastRow(), 1).setNumberFormat('@');
        byId[id] = { _row: sh.getLastRow(), id: id };
        added++;
      }
    });
    return { added: added, updated: updated, list: roomsWithStudents_() };
  });
};


API.deleteStudent = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  var id = String(p.id || '').trim();
  if (!id) fail_('ไม่พบรหัสนักศึกษา');
  return withLock_(function () {
    var st = readRows_('Students').filter(function (x) { return String(x.id).trim() === id; })[0];
    if (!st) fail_('ไม่พบนักศึกษา');
    sheet_('Students').deleteRow(st._row);
    return { list: roomsWithStudents_() };
  });
};


API.resetStudent = function (p) {
  var s = session_(p.token); need_(s, ['admin', 'teacher']);
  return withLock_(function () {
    var st = readRows_('Students').filter(function (x) { return String(x.id).trim() === String(p.id).trim(); })[0];
    if (!st) fail_('ไม่พบนักศึกษา');
    sheet_('Students').getRange(st._row, 5, 1, 2).setValues([['', '']]);
    return {};
  });
};


/* ---------- จุดเข้าเดียวจากหน้าเว็บ ---------- */
function api(name, json) {
  try {
    if (!API[name]) fail_('คำขอไม่ถูกต้อง');
    return JSON.stringify({ ok: true, data: API[name](JSON.parse(json || '{}')) });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.userMessage || 'เกิดข้อผิดพลาด กรุณาลองใหม่' });
  }
}
