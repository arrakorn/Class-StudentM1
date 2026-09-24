const SHEET_ID = '1gtDZIxFXdOC5u_Ry0YOFOQByJWkwXoPSQVtwKUGsA8M';
const TEMPLATE_SLIDE_ID = '1L4lX5VwihH_fqLbUpeIIRbVyTYfrLxwcdFiNrs49050';
const FOLDER_ID = '1PZow9X6lnc1f1E20FvLAIU4URJcvKOgM';

function sheet_(name, header) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName(name) || ss.insertSheet(name);
  if (s.getLastRow() === 0) s.appendRow(header);
  return s;
}
const scoreSheet = () => sheet_('คะแนน', ['เวลา','ชื่อ-นามสกุล','ก่อนเรียน','หลังเรียน','พัฒนาการ','ระดับคุณภาพ']);
const certSheet  = () => sheet_('เกียรติบัตร', ['IDเกียรติบัตร','ชื่อผู้รับ','ตำแหน่ง','ประทับเวลา','ลิงก์ PDF','File ID']);

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// เปิดลิงก์ /exec ตรง ๆ = แสดงหน้าเว็บ (ไฟล์ Index.html) / ถ้ามี action = เรียก API
function doGet(e) {
  if (e && e.parameter && e.parameter.action) return handle(e.parameter);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('บทเรียนออนไลน์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) { return handle(e.parameter); }

// ใช้สำหรับ fetch (เปิดเว็บจากที่อื่น)
function handle(p) { return out(route(p)); }

// ใช้สำหรับ google.script.run (เปิดเว็บจากลิงก์ /exec ของ Apps Script)
function apiCall(p) { return route(p); }

function route(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    switch (p.action) {
      case 'saveScore':   return saveScore(p);
      case 'leaderboard': return leaderboard();
      case 'nextId':      return { ok: true, id: nextId() };
      case 'createCert':  return createCert(p);
      case 'searchCert':  return searchCert(p.q);
      default:            return { ok: false, msg: 'unknown action' };
    }
  } catch (err) {
    return { ok: false, msg: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// รันฟังก์ชันนี้ 1 ครั้งในหน้าแก้ไข เพื่อกดอนุญาตสิทธิ์ Sheets / Drive / Slides
function authorize() {
  SpreadsheetApp.openById(SHEET_ID).getName();
  DriveApp.getFolderById(FOLDER_ID).getName();
  SlidesApp.openById(TEMPLATE_SLIDE_ID).getName();
  scoreSheet(); certSheet();
}

function levelOf(pct) {
  return pct >= 90 ? 'ดีเยี่ยม' : pct >= 80 ? 'ดีมาก' : pct >= 70 ? 'ดี' : pct >= 60 ? 'พอใช้' : 'ควรปรับปรุง';
}

function saveScore(p) {
  const s = scoreSheet();
  const data = s.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) if (data[i][1] === p.name) row = i + 1;
  const pre = p.pre === '' ? '' : Number(p.pre);
  const post = p.post === '' ? '' : Number(p.post);
  const dev = (pre !== '' && post !== '') ? post - pre : '';
  const lvl = post !== '' ? levelOf(post * 10) : '';
  const vals = [new Date(), p.name, pre, post, dev, lvl];
  if (row > 0) s.getRange(row, 1, 1, 6).setValues([vals]); else s.appendRow(vals);
  return { ok: true };
}

function leaderboard() {
  const d = scoreSheet().getDataRange().getValues().slice(1);
  const rows = d.filter(r => r[1] !== '' && r[3] !== '')
    .map(r => ({ name: r[1], pre: Number(r[2]) || 0, post: Number(r[3]) || 0 }))
    .sort((a, b) => b.post - a.post);
  return { ok: true, rows };
}

function nextId() {
  const s = certSheet();
  const n = s.getLastRow() > 1
    ? Math.max(...s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0))
    : 0;
  return String(n + 1).padStart(3, '0');
}

function createCert(p) {
  const id = nextId();
  const copy = DriveApp.getFileById(TEMPLATE_SLIDE_ID).makeCopy('cert-' + id);
  const pres = SlidesApp.openById(copy.getId());
  const map = {
    '{IDเกียรติบัตร}': id,
    '{ชื่อผู้รับ}': p.name,
    '{ตำแหน่ง}': p.position,
    '{ประทับเวลา}': p.ts
  };
  Object.keys(map).forEach(k => pres.replaceAllText(k, map[k]));
  pres.saveAndClose();
  const blob = copy.getAs('application/pdf').setName('เกียรติบัตร_' + id + '_' + p.name + '.pdf');
  const pdf = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  copy.setTrashed(true);
  const url = 'https://drive.google.com/file/d/' + pdf.getId() + '/view';
  certSheet().appendRow(["'" + id, p.name, p.position, p.ts, url, pdf.getId()]);
  return { ok: true, id, fileId: pdf.getId(), url };
}

function searchCert(q) {
  q = String(q || '').toLowerCase();
  const d = certSheet().getDataRange().getValues().slice(1);
  const rows = d.filter(r => String(r[0]).toLowerCase().includes(q) || String(r[1]).toLowerCase().includes(q))
    .map(r => ({ id: r[0], name: r[1], position: r[2], ts: r[3], url: r[4], fileId: r[5] }));
  return { ok: true, rows };
}
