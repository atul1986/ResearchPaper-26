/**
 * TE Research Paper Submission Portal: secure backend API
 * Ajeenkya D Y Patil University · School of Engineering · Third Year (TE)
 *
 * Frontend (GitHub Pages) → this Google Apps Script web app (JSON API) → your Google Sheet + Drive.
 * Students: open access (no login) to submit and track.
 * Faculty: username + password (salted & hashed), role-based access, division scoping,
 *          masked contact numbers, private Drive files, full audit log, login lockout.
 *
 * STORAGE IS FULLY SEPARATE from the DSL Lab portal: own Google Sheet, own Drive folder,
 * and every saved setting/counter/session uses an "RP_" prefix, so the two can never mix,
 * even if both scripts were accidentally placed in the same Apps Script project.
 *
 * Nothing secret is stored in this file, so it is safe to keep in a public GitHub repo.
 * See README.md for the 5-minute deployment steps.
 */

const CONFIG = {
  UNIVERSITY: 'Ajeenkya D Y Patil University',
  SCHOOL: 'School of Engineering',
  LAB_NAME: 'TE Research Paper Submission',
  CLASS_NAME: 'Third Year (TE) B.Tech',
  ACADEMIC_YEAR: '2026-27',
  DEADLINE: '2026-11-25T23:59:00+05:30',   // '' = no deadline
  ACCEPT_LATE: true,                        // false = close submissions after the deadline
  MAX_FILE_MB: 10,
  DIVISIONS: 'ABCDE'.split(''),
  PROGRAMS: [
    'B.Tech Computer Engineering',
    'B.Tech Computer Science & Engineering',
    'B.Tech CSE (Artificial Intelligence & Machine Learning)',
    'B.Tech CSE (Data Science)',
    'B.Tech CSE (Cyber Security)',
    'B.Tech Artificial Intelligence & Data Science',
    'B.Tech Information Technology',
    'B.Tech Electronics & Telecommunication Engineering',
    'B.Tech Robotics & Automation',
    'B.Tech Mechanical Engineering',
    'B.Tech Civil Engineering'
  ],
  DOMAINS: ['Artificial Intelligence & Machine Learning', 'Deep Learning', 'Data Science & Analytics', 'Natural Language Processing',
            'Computer Vision', 'Internet of Things', 'Cyber Security', 'Blockchain', 'Cloud & Edge Computing', 'Computer Networks',
            'Software Engineering', 'Embedded Systems & Robotics', 'Healthcare Informatics', 'Other'],
  MAX_PLAG_PERCENT: 15,                                   // students see a warning above this similarity %
  BLOCK_DUPLICATE_URN: true,
  STATUSES: ['Submitted', 'Under Review', 'Approved', 'Needs Revision'],
  SESSION_MINUTES: 120,
  // ---- Email notifications (sent from the Google account that deployed this script)
  EMAIL_ENABLED: true,
  NOTIFY_ON_REVIEW: true,                         // email the team when faculty update the status
  SENDER_NAME: 'TE Research Papers · ADYPU School of Engineering',
  REPLY_TO: 'atul.kathole1910@gmail.com',
  SITE_URL: 'https://atul1986.github.io/ResearchPaper-26/',
  LOGO_URL: 'https://adypu.edu.in/wp-content/themes/adypu/images/logo.png',
  MAX_LOGIN_FAILS: 5,
  LOCK_MINUTES: 15
};

const TZ = 'Asia/Kolkata';
const SH_SUB = 'Submissions', SH_FAC = 'Faculty', SH_LOG = 'Audit Log';
const HEADERS = [
  'Submission ID', 'Submitted On', 'Program', 'Division', 'Paper Title', 'Research Area', 'Objective', 'Plagiarism %', 'Team Size',
  'S1 Name (Leader)', 'S1 URN', 'S1 Contact', 'S2 Name', 'S2 URN', 'S2 Contact',
  'S3 Name', 'S3 URN', 'S3 Contact', 'S4 Name', 'S4 URN', 'S4 Contact',
  'File Name', 'File Link', 'File ID', 'Plag Report Name', 'Plag Report Link', 'Plag Report ID', 'Late', 'Status', 'Faculty Remarks', 'Reviewed On', 'Reviewed By',
  'S1 Email', 'S2 Email', 'S3 Email', 'S4 Email', 'Email Log'
];
const FAC_HEADERS = ['Username', 'Name', 'Role', 'Divisions', 'Salt', 'Hash', 'Active', 'Created On', 'Last Login'];
const LOG_HEADERS = ['Time', 'User', 'Action', 'Details'];
const COL = {}; HEADERS.forEach((h, i) => COL[h] = i);

/* ================================================================ SETUP */

/** Run once from the Apps Script editor. Prints your admin username + a one-time random password. */
function setup() {
  const p = PropertiesService.getScriptProperties();
  let ss = p.getProperty('RP_SHEET_ID') ? SpreadsheetApp.openById(p.getProperty('RP_SHEET_ID')) : null;
  if (!ss) { ss = SpreadsheetApp.create('TE Research Paper Submissions – AY ' + CONFIG.ACADEMIC_YEAR); p.setProperty('RP_SHEET_ID', ss.getId()); }
  const mk = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      const first = ss.getSheets()[0];
      sh = (name === SH_SUB && ss.getSheets().length === 1 && first.getLastRow() === 0) ? first.setName(name) : ss.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#111827').setFontColor('#fff');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    }
    return sh;
  };
  mk(SH_SUB, HEADERS); const fac = mk(SH_FAC, FAC_HEADERS); mk(SH_LOG, LOG_HEADERS);
  ss.getSheets().forEach(s => { if (s.getName().indexOf('Sheet') === 0 && s.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s); });

  if (!p.getProperty('RP_FOLDER_ID')) p.setProperty('RP_FOLDER_ID', DriveApp.createFolder('TE Research Paper Submissions – AY ' + CONFIG.ACADEMIC_YEAR).getId());

  if (fac.getLastRow() < 2) {
    const pw = randomPassword_();
    const salt = Utilities.getUuid();
    fac.appendRow(['admin', 'Research Coordinator', 'admin', 'ALL', salt, hash_(pw, salt), 'YES', now_(), '']);
    Logger.log('================  SAVE THIS  ================');
    Logger.log('Admin username : admin');
    Logger.log('Admin password : ' + pw + '   (change it after first login)');
    Logger.log('=============================================');
  }
  Logger.log('Portal       : TE RESEARCH PAPERS (separate from DSL Lab)');
  Logger.log('Google Sheet : ' + ss.getUrl());
  Logger.log('Drive folder : ' + DriveApp.getFolderById(p.getProperty('RP_FOLDER_ID')).getUrl());
}

/** Emergency: run from the editor to reset the "admin" password if it is forgotten. */
function resetAdminPassword() {
  const sh = db_().ss.getSheetByName(SH_FAC), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === 'admin') {
    const pw = randomPassword_(), salt = Utilities.getUuid();
    sh.getRange(i + 1, 5, 1, 3).setValues([[salt, hash_(pw, salt), 'YES']]);
    Logger.log('New admin password: ' + pw);
    return;
  }
}

const PORTAL_ID = 'TE-RESEARCH-PAPERS';

function db_() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('RP_SHEET_ID') || !p.getProperty('RP_FOLDER_ID')) setup();
  const ss = SpreadsheetApp.openById(p.getProperty('RP_SHEET_ID'));
  const sub = ss.getSheetByName(SH_SUB);
  if (!sub || sub.getRange(1, 5).getValue() !== 'Paper Title') throw new Error('Storage check failed: the linked sheet is not the TE Research Paper sheet.');
  if (sub.getLastColumn() < HEADERS.length) {             // upgrade older sheets with new columns
    sub.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#111827').setFontColor('#fff');
    sub.getRange(1, 1, sub.getMaxRows(), HEADERS.length).setNumberFormat('@');
  }
  return { p, ss, sub, folder: DriveApp.getFolderById(p.getProperty('RP_FOLDER_ID')) };
}

/* ================================================================ HTTP */

function doGet() { return json_({ ok: true, service: 'TE Research Paper Submission API', time: new Date().toISOString() }); }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const fn = ROUTES[req.action];
    if (!fn) throw new Error('Unknown action');
    out = { ok: true, data: fn(req) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const ROUTES = {
  // ---- public (students)
  init: () => ({ config: publicConfig_(), stats: stats_() }),
  submit: r => submit_(r.payload || {}),
  track: r => track_(r.id, r.urn),
  // ---- faculty (authenticated)
  login: r => login_(r.username, r.password),
  logout: r => { CacheService.getScriptCache().remove('rp_s_' + r.token); return true; },
  dashboard: r => dashboard_(auth_(r)),
  review: r => review_(auth_(r), r.id, r.status, r.remarks),
  reveal: r => reveal_(auth_(r), r.id),
  file: r => file_(auth_(r), r.id, r.which),
  changePassword: r => changePassword_(auth_(r), r.oldPassword, r.newPassword),
  // ---- admin only
  export: r => export_(auth_(r, true)),
  logs: r => logs_(auth_(r, true)),
  facultyList: r => facultyList_(auth_(r, true)),
  facultySave: r => facultySave_(auth_(r, true), r.account || {}),
  facultyDelete: r => facultyDelete_(auth_(r, true), r.username),
  testEmail: r => testEmail_(auth_(r, true)),
  saveSettings: r => saveSettings_(auth_(r, true), r.deadline, r.acceptLate),
  links: r => { auth_(r, true); const d = db_(); return { sheetUrl: d.ss.getUrl(), folderUrl: d.folder.getUrl() }; }
};

/* ================================================================ HELPERS */

/* Deadline & late rule can be changed from the faculty dashboard (stored in Script Properties). */
function deadline_() { const v = PropertiesService.getScriptProperties().getProperty('RP_DEADLINE'); return v == null ? CONFIG.DEADLINE : v; }
function acceptLate_() { const v = PropertiesService.getScriptProperties().getProperty('RP_ACCEPT_LATE'); return v == null ? CONFIG.ACCEPT_LATE : v === 'true'; }
const now_ = () => Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm:ss');
const clean_ = (s, max) => String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 300);
const safe_ = s => /^[=+\-@]/.test(String(s)) ? "'" + s : String(s);   // block spreadsheet formula injection
const un_ = v => String(v == null ? '' : v).replace(/^'/, '');
const mask_ = c => c ? c.slice(0, 2) + 'XXXXX' + c.slice(-3) : '';
const maskMail_ = e => { if (!e) return ''; const [u, d] = e.split('@'); return u.slice(0, 2) + '•••@' + (d || ''); };
function randomPassword_() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789@#$%';
  let s = ''; const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  for (let i = 0; i < 14; i++) s += a[(b[i] + 256) % a.length];
  return s;
}
function hash_(pw, salt) {
  let h = String(pw);
  for (let i = 0; i < 1000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + h + '|' + i));
  return h;
}
function log_(user, action, details) {
  try { db_().ss.getSheetByName(SH_LOG).appendRow([now_(), user, action, safe_(clean_(details, 500))]); } catch (e) {}
}
function publicConfig_() {
  const c = {}; ['UNIVERSITY', 'SCHOOL', 'LAB_NAME', 'CLASS_NAME', 'ACADEMIC_YEAR', 'DEADLINE', 'ACCEPT_LATE', 'MAX_FILE_MB', 'DIVISIONS', 'PROGRAMS', 'DOMAINS', 'STATUSES', 'MAX_PLAG_PERCENT'].forEach(k => c[k] = CONFIG[k]);
  c.DEADLINE = deadline_(); c.ACCEPT_LATE = acceptLate_();
  c.serverNow = new Date().toISOString();
  return c;
}
function rows_() {
  const sh = db_().sub, n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEADERS.length).getValues();
}
function rowObj_(r, full) {
  const members = [1, 2, 3, 4].map(i => ({
    name: un_(r[COL['S' + i + (i === 1 ? ' Name (Leader)' : ' Name')]]),
    urn: un_(r[COL['S' + i + ' URN']]),
    contact: full ? un_(r[COL['S' + i + ' Contact']]) : mask_(un_(r[COL['S' + i + ' Contact']])),
    email: full ? un_(r[COL['S' + i + ' Email']]) : maskMail_(un_(r[COL['S' + i + ' Email']]))
  })).filter(m => m.name || m.urn);
  return {
    id: un_(r[0]), submittedOn: un_(r[COL['Submitted On']]), program: un_(r[COL['Program']]), division: un_(r[COL['Division']]),
    title: un_(r[COL['Paper Title']]), domain: un_(r[COL['Research Area']]), objective: un_(r[COL['Objective']]), plag: un_(r[COL['Plagiarism %']]),
    members, fileName: un_(r[COL['File Name']]), plagFileName: un_(r[COL['Plag Report Name']]), late: r[COL['Late']] === 'YES',
    status: un_(r[COL['Status']]) || 'Submitted', remarks: un_(r[COL['Faculty Remarks']]),
    reviewedOn: un_(r[COL['Reviewed On']]), reviewedBy: un_(r[COL['Reviewed By']]), emailLog: un_(r[COL['Email Log']])
  };
}

/* ================================================================ PUBLIC: stats */

function stats_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('rp_stats');
  if (hit) return JSON.parse(hit);
  const byDiv = {}; CONFIG.DIVISIONS.forEach(d => byDiv[d] = 0);
  let students = 0;
  const rs = rows_();
  rs.forEach(r => { byDiv[r[COL['Division']]] = (byDiv[r[COL['Division']]] || 0) + 1; students += Number(r[COL['Team Size']]) || 0; });
  const s = { total: rs.length, students, byDiv, updated: new Date().toISOString() };
  cache.put('rp_stats', JSON.stringify(s), 30);
  return s;
}

/* ================================================================ PUBLIC: submit */

const RX_PHONE = /^[6-9]\d{9}$/, RX_URN = /^[A-Za-z0-9\/\-]{4,25}$/, RX_EMAIL = /^[^\s@<>()",;:]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;
const EXT = { pdf: 'application/pdf', doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
const EXT_PLAG = Object.assign({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg' }, EXT);

function validate_(p) {
  if (p.website) throw new Error('Rejected.');               // honeypot for bots
  const o = {};
  o.program = clean_(p.program, 120); if (CONFIG.PROGRAMS.indexOf(o.program) < 0) throw new Error('Please select a valid program.');
  o.division = clean_(p.division, 2).toUpperCase(); if (CONFIG.DIVISIONS.indexOf(o.division) < 0) throw new Error('Please select a valid division.');
  o.title = clean_(p.title, 250); if (o.title.length < 5) throw new Error('Paper title must be at least 5 characters.');
  o.domain = clean_(p.domain, 60);
  o.objective = clean_(p.objective, 1000); if (o.objective.length < 20) throw new Error('Please write the objective of the paper (at least 20 characters).');
  o.plag = clean_(p.plag, 6);
  if (o.plag !== '' && !(Number(o.plag) >= 0 && Number(o.plag) <= 100)) throw new Error('Plagiarism % must be between 0 and 100.');
  o.members = []; const seen = {};
  (p.members || []).slice(0, 4).forEach((m, i) => {
    const name = clean_(m && m.name, 80), urn = clean_(m && m.urn, 25).toUpperCase(), contact = clean_(m && m.contact, 10), email = clean_(m && m.email, 120).toLowerCase();
    if (i > 0 && !(name || urn || contact || email)) return;
    const L = i === 0 ? 'Student 1 (Team Leader)' : 'Student ' + (i + 1);
    if (name.length < 2) throw new Error(L + ': name is required.');
    if (!RX_URN.test(urn)) throw new Error(L + ': enter a valid URN.');
    if (i === 0 && !contact) throw new Error(L + ': contact number is required.');
    if (contact && !RX_PHONE.test(contact)) throw new Error(L + ': contact must be a 10-digit mobile number.');
    if (i === 0 && !email) throw new Error(L + ': email is required (confirmation is sent there).');
    if (email && !RX_EMAIL.test(email)) throw new Error(L + ': enter a valid email address.');
    if (seen[urn]) throw new Error('URN ' + urn + ' is entered twice.');
    seen[urn] = 1; o.members.push({ name, urn, contact, email });
  });
  if (!o.members.length) throw new Error('Student 1 details are required.');
  o.file = fileCheck_(p.file, 'research paper', EXT);
  o.plagFile = fileCheck_(p.plagFile, 'plagiarism report', EXT_PLAG);
  return o;
}
function fileCheck_(f, label, allowed) {
  f = f || {};
  const ext = String(f.name || '').split('.').pop().toLowerCase();
  if (!f.data) throw new Error('Please attach the ' + label + '.');
  if (!allowed[ext]) throw new Error('The ' + label + ' must be ' + Object.keys(allowed).join(' / ').toUpperCase() + '.');
  const bytes = Utilities.base64Decode(f.data);
  if (bytes.length > CONFIG.MAX_FILE_MB * 1048576) throw new Error('The ' + label + ' is larger than ' + CONFIG.MAX_FILE_MB + ' MB.');
  if (ext === 'pdf' && String.fromCharCode.apply(null, bytes.slice(0, 4)) !== '%PDF') throw new Error('The ' + label + ' does not look like a valid PDF file.');
  return { bytes, ext, mime: allowed[ext], original: clean_(f.name, 150) };
}

function submit_(payload) {
  const cache = CacheService.getScriptCache();
  const rate = Number(cache.get('rp_rate') || 0);
  if (rate > 60) throw new Error('Server is busy. Please try again in a minute.');
  cache.put('rp_rate', String(rate + 1), 60);

  const d = validate_(payload);
  const dl = deadline_();
  const late = !!(dl && new Date() > new Date(dl));
  if (late && !acceptLate_()) throw new Error('The submission deadline has passed. Please contact your lab faculty.');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const s = db_(), rs = rows_();
    if (CONFIG.BLOCK_DUPLICATE_URN) {
      for (const r of rs) for (const k of ['S1 URN', 'S2 URN', 'S3 URN', 'S4 URN']) {
        const u = String(r[COL[k]] || '').toUpperCase();
        if (u && d.members.some(m => m.urn === u)) throw new Error('URN ' + u + ' is already part of submission ' + r[0] + '. Each student can be in only one team. Contact your lab faculty for changes.');
      }
    }
    const queue = rs.filter(r => r[COL['Division']] === d.division && (r[COL['Status']] === 'Submitted' || r[COL['Status']] === 'Under Review')).length + 1;
    const seq = Number(s.p.getProperty('RP_SEQ_' + d.division) || 0) + 1;
    const id = 'RP-' + d.division + '-' + ('000' + seq).slice(-3);
    const it = s.folder.getFoldersByName('Division ' + d.division);
    const folder = it.hasNext() ? it.next() : s.folder.createFolder('Division ' + d.division);
    const fname = id + '_' + d.members[0].urn + '_' + d.title.replace(/[^A-Za-z0-9 ]+/g, '').trim().replace(/\s+/g, '_').slice(0, 40) + '.' + d.file.ext;
    const file = folder.createFile(Utilities.newBlob(d.file.bytes, d.file.mime, fname));
    const pname = id + '_PlagiarismReport.' + d.plagFile.ext;
    const pfile = folder.createFile(Utilities.newBlob(d.plagFile.bytes, d.plagFile.mime, pname));

    const row = new Array(HEADERS.length).fill('');
    const set = (k, v) => row[COL[k]] = v;
    set('Submission ID', id); set('Submitted On', now_()); set('Program', d.program); set('Division', d.division);
    set('Paper Title', safe_(d.title)); set('Research Area', safe_(d.domain)); set('Objective', safe_(d.objective)); set('Plagiarism %', d.plag);
    set('Team Size', String(d.members.length));
    d.members.forEach((m, i) => { const n = 'S' + (i + 1); set(n + (i ? ' Name' : ' Name (Leader)'), safe_(m.name)); set(n + ' URN', m.urn); set(n + ' Contact', m.contact); set(n + ' Email', m.email); });
    set('File Name', fname); set('File Link', file.getUrl()); set('File ID', file.getId());
    set('Plag Report Name', pname); set('Plag Report Link', pfile.getUrl()); set('Plag Report ID', pfile.getId()); set('Late', late ? 'YES' : 'NO'); set('Status', 'Submitted');
    s.sub.getRange(s.sub.getLastRow() + 1, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
    s.p.setProperty('RP_SEQ_' + d.division, String(seq));
    CacheService.getScriptCache().remove('rp_stats');
    const rowNum = s.sub.getLastRow();
    lock.releaseLock();
    const mail = sendSubmissionEmail_({ id, late, queue, submittedOn: row[COL['Submitted On']], program: d.program, division: d.division, title: d.title, members: d.members, fileName: fname, plagFileName: pname, plag: d.plag, objective: d.objective, domain: d.domain });
    s.sub.getRange(rowNum, COL['Email Log'] + 1).setValue(mail.log);
    return { id, late, queue, emailed: mail.sentTo, submittedOn: row[COL['Submitted On']], program: d.program, division: d.division, title: d.title,
             members: d.members.map(m => ({ name: m.name, urn: m.urn })), fileName: fname, plagFileName: pname };
  } finally { lock.releaseLock(); }
}

function track_(id, urn) {
  id = clean_(id, 20).toUpperCase(); urn = clean_(urn, 25).toUpperCase();
  if (!id || !urn) throw new Error('Enter both Submission ID and URN.');
  for (const r of rows_()) {
    if (String(r[0]).toUpperCase() !== id) continue;
    const o = rowObj_(r, false);
    if (!o.members.some(m => m.urn.toUpperCase() === urn)) break;
    return { id: o.id, title: o.title, division: o.division, program: o.program, submittedOn: o.submittedOn, status: o.status,
             remarks: o.remarks, reviewedOn: o.reviewedOn, late: o.late, members: o.members.map(m => m.name) };
  }
  throw new Error('No submission found for this ID and URN.');
}

/* ================================================================ FACULTY: auth */

function facRows_() { const sh = db_().ss.getSheetByName(SH_FAC); return { sh, v: sh.getDataRange().getValues() }; }

function login_(username, password) {
  username = clean_(username, 40).toLowerCase(); password = String(password || '');
  const cache = CacheService.getScriptCache(), lk = 'rp_f_' + username, fails = Number(cache.get(lk) || 0);
  if (fails >= CONFIG.MAX_LOGIN_FAILS) throw new Error('Account temporarily locked after repeated failed attempts. Try again in ' + CONFIG.LOCK_MINUTES + ' minutes.');
  Utilities.sleep(400);                                    // slow down guessing
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() !== username) continue;
    if (v[i][6] !== 'YES' || hash_(password, v[i][4]) !== v[i][5]) break;
    cache.remove(lk);
    const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    const sess = { u: v[i][0], name: v[i][1], role: v[i][2], divs: String(v[i][3] || 'ALL') };
    cache.put('rp_s_' + token, JSON.stringify(sess), CONFIG.SESSION_MINUTES * 60);
    sh.getRange(i + 1, 9).setValue(now_());
    log_(sess.u, 'LOGIN', 'success');
    return { token, user: sess, minutes: CONFIG.SESSION_MINUTES };
  }
  cache.put(lk, String(fails + 1), CONFIG.LOCK_MINUTES * 60);
  log_(username || '?', 'LOGIN_FAILED', 'attempt ' + (fails + 1));
  throw new Error('Invalid username or password.');
}

function auth_(r, adminOnly) {
  const cache = CacheService.getScriptCache(), raw = r && r.token && cache.get('rp_s_' + r.token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  cache.put('rp_s_' + r.token, raw, CONFIG.SESSION_MINUTES * 60);   // sliding expiry
  const s = JSON.parse(raw);
  if (adminOnly && s.role !== 'admin') throw new Error('Only the lab coordinator (admin) can do this.');
  return s;
}
const inScope_ = (s, div) => s.divs === 'ALL' || s.divs.split(',').map(x => x.trim()).indexOf(div) >= 0;
function findRow_(s, id) {
  const rs = rows_();
  for (let i = 0; i < rs.length; i++) if (String(rs[i][0]) === id) {
    if (!inScope_(s, rs[i][COL['Division']])) throw new Error('This division is outside your access.');
    return { r: rs[i], rowNum: i + 2 };
  }
  throw new Error('Submission not found.');
}

/* ================================================================ FACULTY: data */

function dashboard_(s) {
  let quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  return { user: s, config: publicConfig_(), emailQuota: CONFIG.EMAIL_ENABLED ? quota : null, rows: rows_().filter(r => inScope_(s, r[COL['Division']])).map(r => rowObj_(r, false)) };
}
function review_(s, id, status, remarks) {
  if (CONFIG.STATUSES.indexOf(status) < 0) throw new Error('Invalid status.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(s, String(id));
    const when = Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm');
    const sub = db_().sub;
    const changed = un_(f.r[COL['Status']]) !== status || un_(f.r[COL['Faculty Remarks']]) !== clean_(remarks, 500);
    sub.getRange(f.rowNum, COL['Status'] + 1, 1, 4).setValues([[status, safe_(clean_(remarks, 500)), when, s.name || s.u]]);
    log_(s.u, 'REVIEW', id + ' → ' + status);
    let emailed = '';
    if (changed && CONFIG.NOTIFY_ON_REVIEW) {
      const o = rowObj_(f.r, true);
      const mail = sendReviewEmail_(o, status, clean_(remarks, 500), s.name || s.u);
      emailed = mail.sentTo;
      sub.getRange(f.rowNum, COL['Email Log'] + 1).setValue((un_(f.r[COL['Email Log']]) + ' | ' + mail.log).replace(/^ \| /, '').slice(-480));
    }
    return { reviewedOn: when, reviewedBy: s.name || s.u, emailed };
  } finally { lock.releaseLock(); }
}
function reveal_(s, id) {
  const f = findRow_(s, String(id));
  log_(s.u, 'VIEW_CONTACTS', id);
  return rowObj_(f.r, true).members;
}
function file_(s, id, which) {
  const f = findRow_(s, String(id)), plag = which === 'plag';
  const blob = DriveApp.getFileById(f.r[COL[plag ? 'Plag Report ID' : 'File ID']]).getBlob();
  log_(s.u, plag ? 'DOWNLOAD_PLAG_REPORT' : 'DOWNLOAD_PAPER', id);
  return { name: f.r[COL[plag ? 'Plag Report Name' : 'File Name']], mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
function export_(s) {
  const q = v => '"' + un_(v).replace(/"/g, '""') + '"';
  const keep = HEADERS.filter(h => h !== 'File ID' && h !== 'Plag Report ID');
  const csv = [keep.map(q).join(',')].concat(rows_().map(r => keep.map(h => q(r[COL[h]])).join(','))).join('\n');
  log_(s.u, 'EXPORT', 'full CSV export');
  return { filename: 'TE_Research_Papers_' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm') + '.csv', csv };
}
function logs_() {
  const sh = db_().ss.getSheetByName(SH_LOG), n = sh.getLastRow();
  if (n < 2) return [];
  const start = Math.max(2, n - 299);
  return sh.getRange(start, 1, n - start + 1, 4).getValues().reverse().map(r => ({ time: r[0], user: r[1], action: r[2], details: un_(r[3]) }));
}
function changePassword_(s, oldPw, newPw) {
  newPw = String(newPw || '');
  if (newPw.length < 8 || !/[A-Za-z]/.test(newPw) || !/\d/.test(newPw)) throw new Error('New password needs 8+ characters with letters and numbers.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (v[i][0] === s.u) {
    if (hash_(String(oldPw || ''), v[i][4]) !== v[i][5]) throw new Error('Current password is incorrect.');
    const salt = Utilities.getUuid(); sh.getRange(i + 1, 5, 1, 2).setValues([[salt, hash_(newPw, salt)]]);
    log_(s.u, 'PASSWORD_CHANGED', ''); return true;
  }
  throw new Error('Account not found.');
}

/* ================================================================ ADMIN: faculty accounts */

function facultyList_() {
  return facRows_().v.slice(1).map(r => ({ username: r[0], name: r[1], role: r[2], divisions: r[3], active: r[6] === 'YES', created: r[7], lastLogin: r[8] }));
}
function facultySave_(s, a) {
  const username = clean_(a.username, 40).toLowerCase();
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('Username: 3–40 characters (letters, numbers, . _ -).');
  const role = a.role === 'admin' ? 'admin' : 'faculty';
  let divs = clean_(a.divisions, 60).toUpperCase().replace(/\s+/g, '');
  if (divs !== 'ALL') {
    divs = divs.split(',').filter(x => CONFIG.DIVISIONS.indexOf(x) >= 0).join(',');
    if (!divs) throw new Error('Choose at least one division (or ALL).');
  }
  const pw = String(a.password || '');
  if (pw && (pw.length < 8 || !/[A-Za-z]/.test(pw) || !/\d/.test(pw))) throw new Error('Password needs 8+ characters with letters and numbers.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === username) {
    if (username === s.u && (role !== 'admin' || a.active === false)) throw new Error('You cannot remove your own admin access.');
    sh.getRange(i + 1, 2, 1, 3).setValues([[safe_(clean_(a.name, 80)), role, divs]]);
    sh.getRange(i + 1, 7).setValue(a.active === false ? 'NO' : 'YES');
    if (pw) { const salt = Utilities.getUuid(); sh.getRange(i + 1, 5, 1, 2).setValues([[salt, hash_(pw, salt)]]); }
    log_(s.u, 'FACULTY_UPDATED', username + ' role=' + role + ' div=' + divs + (pw ? ' (password reset)' : ''));
    return true;
  }
  if (!pw) throw new Error('Set a password for the new account.');
  const salt = Utilities.getUuid();
  sh.appendRow([username, safe_(clean_(a.name, 80)), role, divs, salt, hash_(pw, salt), 'YES', now_(), '']);
  log_(s.u, 'FACULTY_CREATED', username + ' role=' + role + ' div=' + divs);
  return true;
}
function facultyDelete_(s, username) {
  username = clean_(username, 40).toLowerCase();
  if (username === s.u) throw new Error('You cannot delete your own account.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === username) { sh.deleteRow(i + 1); log_(s.u, 'FACULTY_DELETED', username); return true; }
  throw new Error('Account not found.');
}

/* ================================================================ EMAIL */

function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

function mailShell_(title, accent, inner) {
  return '<div style="background:#f7f2e8;padding:24px 12px;font-family:Georgia,\'Times New Roman\',serif;color:#1f1a14">' +
    '<div style="max-width:600px;margin:0 auto;background:#fffdf8;border:1px solid #e2d8c6">' +
    '<div style="padding:16px 24px;border-bottom:3px double #cfc2ab;text-align:center"><img src="' + CONFIG.LOGO_URL + '" alt="ADYPU" height="42" style="height:42px"></div>' +
    '<div style="padding:20px 24px 6px;text-align:center"><div style="font-family:Consolas,monospace;font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' + accent + '">' + esc_(CONFIG.SCHOOL) + ' · ' + esc_(CONFIG.LAB_NAME) + '</div>' +
    '<div style="font-size:26px;font-weight:600;margin-top:6px">' + title + '</div></div>' +
    '<div style="padding:14px 24px 20px;font-size:15px;line-height:1.6;font-family:Arial,Helvetica,sans-serif">' + inner + '</div>' +
    '<div style="padding:14px 24px;border-top:1px solid #e2d8c6;font-size:12px;color:#857a6a;font-family:Arial,Helvetica,sans-serif">Research Cell · ' + esc_(CONFIG.SCHOOL) + ', ' + esc_(CONFIG.UNIVERSITY) + ' · AY ' + esc_(CONFIG.ACADEMIC_YEAR) +
    '<br>This is an automated message. For queries, reply to this email.</div></div></div>';
}
function kvTable_(rows) {
  return '<table style="width:100%;border-collapse:collapse;font-size:14px;margin:12px 0">' + rows.map(r =>
    '<tr><td style="padding:7px 0;color:#6b7089;width:120px;vertical-align:top">' + r[0] + '</td><td style="padding:7px 0;font-weight:600">' + r[1] + '</td></tr>').join('') + '</table>';
}
function recipients_(members) {
  const seen = {}; return members.map(m => String(m.email || '').trim().toLowerCase()).filter(e => e && RX_EMAIL.test(e) && !seen[e] && (seen[e] = 1));
}
function send_(to, subject, html, plain) {
  if (!CONFIG.EMAIL_ENABLED) return { sentTo: '', log: 'email off' };
  if (!to.length) return { sentTo: '', log: 'no email address' };
  try {
    if (MailApp.getRemainingDailyQuota() < to.length) return { sentTo: '', log: now_() + ' SKIPPED (daily email quota used up)' };
    MailApp.sendEmail({ to: to.join(','), subject, htmlBody: html, body: plain, name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO });
    return { sentTo: to.join(', '), log: now_() + ' SENT ' + subject.split('·')[0].trim() + ' → ' + to.join(', ') };
  } catch (e) {
    return { sentTo: '', log: now_() + ' FAILED: ' + String(e.message || e).slice(0, 120) };
  }
}

function sendSubmissionEmail_(r) {
  const td = 'padding:7px 8px;border-bottom:1px solid #e2d8c6;font-size:13px;vertical-align:top';
  const th = 'padding:7px 8px;border-bottom:1px solid #1f1a14;font-size:11px;text-align:left;color:#857a6a;text-transform:uppercase;letter-spacing:.8px';
  const authors = '<table style="width:100%;border-collapse:collapse;margin:8px 0 14px"><tr><th style="' + th + '">#</th><th style="' + th + '">Name</th><th style="' + th + '">URN</th><th style="' + th + '">Contact</th><th style="' + th + '">Email</th></tr>' +
    r.members.map((m, i) => '<tr><td style="' + td + '">' + (i + 1) + '</td><td style="' + td + '"><b>' + esc_(m.name) + '</b>' + (i ? '' : '<br><span style="color:#7b1e2b;font-size:11px">first author</span>') + '</td><td style="' + td + ';font-family:Consolas,monospace">' + esc_(m.urn) + '</td><td style="' + td + '">' + esc_(m.contact || '-') + '</td><td style="' + td + '">' + esc_(m.email || '-') + '</td></tr>').join('') + '</table>';
  const inner =
    '<p>Dear ' + esc_(r.members[0].name.split(' ')[0]) + (r.members.length > 1 ? ' and co-authors' : '') + ',</p>' +
    '<p>We acknowledge receipt of your research paper. The details recorded are below. Please check them and keep this email for your records.</p>' +
    '<div style="border:1px solid #e2d8c6;border-left:4px solid #7b1e2b;padding:12px 16px;margin:14px 0;background:#fbf7ef"><div style="font-size:11px;color:#857a6a;letter-spacing:1.4px;text-transform:uppercase">Manuscript ID</div>' +
    '<div style="font-family:Consolas,monospace;font-size:26px;font-weight:700;color:#7b1e2b;letter-spacing:1px">' + esc_(r.id) + '</div>' +
    '<div style="font-size:13px;margin-top:2px">Position #' + r.queue + ' in the Division ' + esc_(r.division) + ' review queue' + (r.late ? ' · <b style="color:#b3261e">submitted after deadline</b>' : '') + '</div></div>' +
    kvTable_([['Submitted on', esc_(r.submittedOn)], ['Program', esc_(r.program)], ['Division', esc_(r.division)], ['Paper title', '<span style="font-family:Georgia,serif;font-size:15px">' + esc_(r.title) + '</span>'],
              ['Objective', esc_(r.objective)], ['Research area', esc_(r.domain || '-')], ['Similarity', r.plag !== '' ? esc_(r.plag) + '%' : 'as per attached report'],
              ['Manuscript', esc_(r.fileName)], ['Plag report', esc_(r.plagFileName)]]) +
    '<div style="font-size:11px;color:#857a6a;letter-spacing:1.4px;text-transform:uppercase;margin-top:6px">Authors</div>' + authors +
    '<p style="font-size:13px;color:#4a4237">If any detail is wrong, reply to this email before the deadline.</p>' +
    '<p><a href="' + CONFIG.SITE_URL + '#track" style="display:inline-block;background:#7b1e2b;color:#fff;text-decoration:none;padding:11px 18px;font-weight:700">Track review status →</a></p>' +
    '<p style="font-size:13px;color:#857a6a">To track, use your Manuscript ID and the URN of any author. You will also get an email when your paper is reviewed.</p>';
  const plain = 'Your TE research paper has been received.\nManuscript ID: ' + r.id + '\nTitle: ' + r.title + '\nObjective: ' + r.objective + '\nProgram: ' + r.program + ' | Division: ' + r.division +
    '\nAuthors: ' + r.members.map(m => m.name + ' (' + m.urn + ')').join(', ') + '\nTrack status: ' + CONFIG.SITE_URL + '#track';
  return send_(recipients_(r.members), r.id + ' · Research paper received ✅', mailShell_('Acknowledgement of Receipt', '#7b1e2b', inner), plain);
}

function sendReviewEmail_(o, status, remarks, by) {
  const col = { 'Approved': '#2f7d4f', 'Needs Revision': '#b3261e', 'Under Review': '#a86b00', 'Submitted': '#7b1e2b' }[status] || '#7b1e2b';
  const msg = {
    'Approved': 'Congratulations! Your research paper has been <b>approved</b> by your faculty.',
    'Needs Revision': 'Your faculty has reviewed your paper and <b>requested some changes</b>. Please go through the remarks below and revise the paper.',
    'Under Review': 'Your research paper is now <b>under review</b> by your faculty.',
    'Submitted': 'The status of your research paper has been updated.'
  }[status];
  const inner = '<p>Dear ' + esc_(o.members[0].name.split(' ')[0]) + ' and team,</p><p>' + msg + '</p>' +
    '<div style="display:inline-block;background:' + col + ';color:#fff;font-weight:700;padding:7px 14px;border-radius:999px;margin:6px 0">' + esc_(status) + '</div>' +
    kvTable_([['Submission ID', '<span style="font-family:Consolas,monospace">' + esc_(o.id) + '</span>'], ['Paper title', esc_(o.title)], ['Division', esc_(o.division)], ['Reviewed by', esc_(by)]]) +
    (remarks ? '<div style="border-left:4px solid ' + col + ';background:#fafaf7;padding:12px 14px;border-radius:8px;margin:10px 0"><div style="font-size:12px;color:#6b7089;margin-bottom:4px">FACULTY REMARKS</div>' + esc_(remarks) + '</div>' : '') +
    '<p><a href="' + CONFIG.SITE_URL + '#track" style="display:inline-block;background:#1f1a14;color:#fff;text-decoration:none;padding:11px 18px;font-weight:700">View on portal →</a></p>';
  const plain = 'TE research paper ' + o.id + ' status: ' + status + (remarks ? '\nRemarks: ' + remarks : '') + '\nTrack: ' + CONFIG.SITE_URL + '#track';
  return send_(recipients_(o.members), o.id + ' · Paper status: ' + status, mailShell_('Review update: ' + esc_(status), col, inner), plain);
}

function testEmail_(s) {
  const r = send_([CONFIG.REPLY_TO], 'TE Research Paper portal · test email', mailShell_('Email is working ✅', '#0f6b67',
    '<p>This test was sent from the TE Research Paper Submission Portal by <b>@' + esc_(s.u) + '</b>.</p><p>Student confirmation and review emails will look like this.</p>'), 'TE Research Paper portal test email');
  log_(s.u, 'TEST_EMAIL', r.log);
  if (!r.sentTo) throw new Error(r.log);
  return { sentTo: r.sentTo, quota: MailApp.getRemainingDailyQuota() };
}

/** Run once from the editor after updating the code, to grant the "send email" permission. */
function authorizeEmail() {
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}

function saveSettings_(s, deadline, acceptLate) {
  deadline = String(deadline || '').trim();
  if (deadline && isNaN(new Date(deadline).getTime())) throw new Error('Invalid deadline date.');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('RP_DEADLINE', deadline);
  p.setProperty('RP_ACCEPT_LATE', acceptLate ? 'true' : 'false');
  CacheService.getScriptCache().remove('rp_stats');
  log_(s.u, 'SETTINGS', 'deadline=' + (deadline || 'none') + ' acceptLate=' + !!acceptLate);
  return { DEADLINE: deadline, ACCEPT_LATE: !!acceptLate };
}
