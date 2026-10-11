/**
 * SG Tutor Match — Google Apps Script backend
 * Run setupSheets() once as the owner. In Project Settings > Script properties,
 * set ADMIN_PASSWORD to a strong unique password before using the admin functions.
 * Deploy as a Web App only after reviewing access and privacy settings.
 */
const SHEET_NAMES = { parents: 'Parents', tutors: 'Tutors', matches: 'Matches' };
const PARENT_HEADERS = ['id','created_at','status','contact_name','contact_phone','contact_wechat','contact_whatsapp','level','subject','budget','mode','region','frequency','duration','timeslot','goal','notes'];
const TUTOR_HEADERS = ['id','created_at','status','name','contact_phone','contact_email','whatsapp','level','subjects','rate','mode','regions','languages','experience','timeslots','bio'];
const MATCH_HEADERS = ['id','created_at','parent_id','tutor_id','score','reason','status'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SG Tutor Match')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('SG Tutor Match Data');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheet_(ss, SHEET_NAMES.parents, PARENT_HEADERS);
  ensureSheet_(ss, SHEET_NAMES.tutors, TUTOR_HEADERS);
  ensureSheet_(ss, SHEET_NAMES.matches, MATCH_HEADERS);
  return { ok: true, spreadsheetUrl: ss.getUrl(), message: 'Sheets created. Keep the spreadsheet private.' };
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    if (current.join('|') !== headers.join('|')) throw new Error('Header mismatch in ' + name + '. Do not overwrite existing data; inspect this sheet first.');
  }
  sh.setFrozenRows(1);
}

function submitPublicRequest(type, payload) {
  if (!['parent', 'tutor'].includes(type)) throw new Error('Invalid request type.');
  const clean = sanitize_(payload || {});
  const required = type === 'parent'
    ? ['contact_name','contact_phone','level','subject','budget','mode','region']
    : ['name','contact_phone','level','subjects','rate','mode'];
  required.forEach(k => { if (!String(clean[k] || '').trim()) throw new Error('Missing required field: ' + k); });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = openData_();
    const isParent = type === 'parent';
    const sheet = ss.getSheetByName(isParent ? SHEET_NAMES.parents : SHEET_NAMES.tutors);
    const headers = isParent ? PARENT_HEADERS : TUTOR_HEADERS;
    const rowObj = isParent ? {
      id: Utilities.getUuid(), created_at: new Date(), status: 'pending',
      contact_name: clean.contact_name, contact_phone: clean.contact_phone,
      contact_wechat: clean.contact_wechat || '', contact_whatsapp: clean.contact_whatsapp || '',
      level: clean.level, subject: normalizeSubject_(clean.subject), budget: clean.budget,
      mode: clean.mode, region: clean.region, frequency: clean.frequency || '',
      duration: clean.duration || '', timeslot: clean.timeslot || '', goal: clean.goal || '', notes: clean.notes || ''
    } : {
      id: Utilities.getUuid(), created_at: new Date(), status: 'active',
      name: clean.name, contact_phone: clean.contact_phone, contact_email: clean.contact_email || '',
      whatsapp: clean.whatsapp || '', level: clean.level,
      subjects: clean.subjects.split(',').map(normalizeSubject_).filter(Boolean).join(', '),
      rate: clean.rate, mode: clean.mode, regions: clean.regions || '',
      languages: clean.languages || '', experience: clean.experience || '',
      timeslots: clean.timeslots || '', bio: clean.bio || ''
    };
    sheet.appendRow(headers.map(h => safeCell_(rowObj[h] == null ? '' : rowObj[h])));
    return { ok: true, message: '提交成功。平台将审核资料并进行配对。' };
  } finally { lock.releaseLock(); }
}

function adminLogin(password) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected) throw new Error('管理员尚未设置。请在 Script Properties 设置 ADMIN_PASSWORD。');
  if (String(password || '') !== expected) throw new Error('密码不正确。');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('admin:' + token, '1', 21600);
  return { ok: true, token: token };
}

function adminGetData(token) {
  requireAdmin_(token);
  const ss = openData_();
  const parents = readObjects_(ss.getSheetByName(SHEET_NAMES.parents));
  const tutors = readObjects_(ss.getSheetByName(SHEET_NAMES.tutors));
  const matches = readObjects_(ss.getSheetByName(SHEET_NAMES.matches));
  return { parents: parents, tutors: tutors, matches: matches };
}

function adminRemove(token, type, id) {
  requireAdmin_(token);
  if (!['parent','tutor'].includes(type)) throw new Error('Invalid type.');
  const ss = openData_();
  const sh = ss.getSheetByName(type === 'parent' ? SHEET_NAMES.parents : SHEET_NAMES.tutors);
  const values = sh.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  const statusCol = values[0].indexOf('status');
  for (let i=1;i<values.length;i++) {
    if (String(values[i][idCol]) === String(id)) {
      sh.getRange(i+1,statusCol+1).setValue('removed');
      return { ok:true };
    }
  }
  throw new Error('Record not found.');
}

function runSmartMatch(token) {
  requireAdmin_(token);
  const ss = openData_();
  const parents = readObjects_(ss.getSheetByName(SHEET_NAMES.parents)).filter(x => x.status === 'pending');
  const tutors = readObjects_(ss.getSheetByName(SHEET_NAMES.tutors)).filter(x => x.status === 'active');
  const sh = ss.getSheetByName(SHEET_NAMES.matches);
  const existing = readObjects_(sh);
  let created = 0;
  parents.forEach(p => {
    tutors.forEach(t => {
      if (existing.some(m => String(m.parent_id) === String(p.id) && String(m.tutor_id) === String(t.id) && m.status !== 'removed')) return;
      const subjectMatch = splitSubjects_(t.subjects).some(s => normalizeSubject_(s) === normalizeSubject_(p.subject));
      if (!subjectMatch) return;
      const levelMatch = levelCompatible_(p.level, t.level);
      const modeMatch = modeCompatible_(p.mode, t.mode);
      const regionMatch = !p.region || !t.regions || String(t.regions).toLowerCase().includes(String(p.region).toLowerCase()) || String(p.region).toLowerCase().includes(String(t.regions).toLowerCase());
      const score = 60 + (levelMatch ? 20 : 0) + (modeMatch ? 10 : 0) + (regionMatch ? 10 : 0);
      sh.appendRow([Utilities.getUuid(), new Date(), p.id, t.id, score,
        '科目一致' + (levelMatch ? '；年级相符' : '；年级需人工确认') + (modeMatch ? '；授课模式相符' : '；授课模式需确认') + (regionMatch ? '；地区可能相符' : '；地区需确认'),
        'suggested']);
      created++;
    });
  });
  return { ok:true, created:created, message:'已生成 ' + created + ' 条候选配对。请管理员人工核对后联系双方。' };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('admin:' + token);
  return { ok:true };
}

function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('admin:' + token) !== '1') throw new Error('登录已失效，请重新登录。');
}
function openData_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('尚未初始化数据库。请管理员先运行 setupSheets()。');
  return SpreadsheetApp.openById(id);
}
function readObjects_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.map(row => headers.reduce((o,h,i) => { let v=row[i]; o[h]=v instanceof Date ? v.toISOString() : v; return o; }, {}));
}
function sanitize_(input) {
  const out = {};
  Object.keys(input).forEach(k => {
    if (!/^[a-z_]+$/.test(k)) return;
    let v = input[k];
    if (typeof v === 'object' || v == null) v = '';
    out[k] = String(v).trim().slice(0, 2000);
  });
  return out;
}
function safeCell_(v) {
  const s = v instanceof Date ? v : String(v);
  return typeof s === 'string' && /^[=+@]/.test(s) ? "'" + s : s;
}
function normalizeSubject_(s) {
  let x = String(s || '').toLowerCase().replace(/\s+/g,' ').trim();
  const aliases = {
    'h2 chemistry':'h2 chemistry', 'chemistry h2':'h2 chemistry',
    'h2 physics':'h2 physics', 'physics h2':'h2 physics',
    'h2 biology':'h2 biology', 'biology h2':'h2 biology',
    'e math':'e-math', 'emath':'e-math', 'elementary mathematics':'e-math',
    'a math':'a-math', 'amath':'a-math', 'additional mathematics':'a-math',
    'social studies':'social studies', 'social study':'social studies'
  };
  return aliases[x] || x;
}
function splitSubjects_(s) { return String(s || '').split(/[,;\n]/).map(x=>x.trim()).filter(Boolean); }
function levelCompatible_(parentLevel, tutorLevel) {
  const p=String(parentLevel||'').toLowerCase(), t=String(tutorLevel||'').toLowerCase();
  if (t.includes('multiple') || t.includes('all')) return true;
  if (p.startsWith('primary') && t.includes('primary')) return true;
  if (p.startsWith('secondary') && t.includes('secondary')) return true;
  if (p.startsWith('ip') && t.includes('ip')) return true;
  if (p.startsWith('jc') && t.includes('jc')) return true;
  return p===t;
}
function modeCompatible_(p,t) { return !p || !t || p===t || p==='均可' || t==='均可'; }
