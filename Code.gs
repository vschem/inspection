/**
 * Code.gs — Backend ของระบบตรวจสภาพรถและแผนซ่อมบำรุง (Google Apps Script + Google Sheets)
 *
 * API ที่เว็บแอปเรียกใช้
 *   GET  ?action=getAll                         -> { vehicle_inspections:[], vehicles:[], personnel:[], maintenance_tracking:[] }
 *   POST {action:'addDoc',    collection, data}
 *   POST {action:'updateDoc', collection, id, data}
 *   POST {action:'deleteDoc', collection, id}
 *
 * วิธีใช้: วางไฟล์นี้ใน Extensions > Apps Script ของ Google Sheet -> รันฟังก์ชัน setup() 1 ครั้ง
 *         -> Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone)
 *         -> นำ URL ที่ได้ไปใส่ที่ WEB_APP_URL ในไฟล์ HTML
 */

// คอลัมน์มาตรฐานของแต่ละชีต (ถ้าเว็บส่งฟิลด์ใหม่มา ระบบจะเพิ่มคอลัมน์ให้เอง)
const SHEETS = {
  vehicle_inspections: ['id', 'licensePlate', 'inspectionDate', 'driverName', 'inspectorName', 'currentMileage',
                        'itemsToFix', 'checklist', 'photoBase64', 'timestamp', 'overallResult'],
  vehicles: ['id', 'plate', 'interval', 'lastServiceKm', 'lastServiceDate'],
  personnel: ['id', 'name', 'role'],
  maintenance_tracking: ['id', 'inspectionId', 'plate', 'date', 'item', 'reportedBy', 'notes', 'status',
                         'priority', 'planDate', 'owner', 'plan', 'doneDate', 'count', 'lastSeen', 'updatedAt',
                         'approval', 'reqAt', 'reqNote', 'decideAt', 'decideNote']
};
const JSON_FIELDS = ['checklist'];      // ฟิลด์ที่เก็บเป็น JSON string และแปลงกลับตอนอ่าน
const PHOTO_FIELDS = ['photoBase64'];   // ฟิลด์รูปภาพ
// อีเมลผู้บริหารที่จะรับแจ้งเตือนเมื่อมีการส่งขออนุมัติ (คั่นหลายที่อยู่ด้วยเครื่องหมายจุลภาค) เว้นว่างไว้ = ไม่ส่งอีเมล
const EXEC_EMAILS = '';
const MAX_CELL = 49000;                 // Google Sheets จำกัด 50,000 ตัวอักษรต่อเซลล์

/* ---------- Entry points ---------- */

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'getAll') {
      const out = {};
      Object.keys(SHEETS).forEach(function (name) { out[name] = readAll_(name); });
      return json_(out);
    }
    return json_({ ok: true, message: 'Vehicle inspection API is running' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // กันเขียนชนกันเมื่อมีหลายคนบันทึกพร้อมกัน
    const req = JSON.parse(e.postData.contents);
    let result;
    switch (req.action) {
      case 'addDoc':    result = addDoc_(req.collection, req.data); break;
      case 'updateDoc': result = updateDoc_(req.collection, req.id, req.data); break;
      case 'deleteDoc': result = deleteDoc_(req.collection, req.id); break;
      case 'notifyExec': result = notifyExec_(req.subject, req.body); break;
      default: throw new Error('Unknown action: ' + req.action);
    }
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** รันครั้งเดียวเพื่อสร้างชีตและหัวคอลัมน์ทั้งหมด */
function setup() {
  Object.keys(SHEETS).forEach(function (name) { getSheet_(name); });
  SpreadsheetApp.getActive().toast('สร้างชีตเรียบร้อยแล้ว');
}

/* ---------- CRUD ---------- */

function addDoc_(name, data) {
  data = data || {};
  if (!data.id) throw new Error('missing id');
  const sheet = getSheet_(name);
  ensureColumns_(sheet, Object.keys(data));
  const headers = headers_(sheet);
  const row = headers.map(function (h) { return toCell_(h, data[h]); });
  const target = sheet.getLastRow() + 1;
  sheet.getRange(target, 1, 1, headers.length).setValues([row]);
  return { ok: true, id: data.id };
}

function updateDoc_(name, id, data) {
  const sheet = getSheet_(name);
  const rowNo = findRow_(sheet, id);
  if (!rowNo) return { ok: false, error: 'not found: ' + id };
  data = data || {};
  ensureColumns_(sheet, Object.keys(data));
  const headers = headers_(sheet);
  const range = sheet.getRange(rowNo, 1, 1, headers.length);
  const row = range.getValues()[0];
  headers.forEach(function (h, i) {
    if (h !== 'id' && Object.prototype.hasOwnProperty.call(data, h)) row[i] = toCell_(h, data[h]);
  });
  range.setValues([row]);
  return { ok: true, id: id };
}

function deleteDoc_(name, id) {
  const sheet = getSheet_(name);
  const rowNo = findRow_(sheet, id);
  if (rowNo) sheet.deleteRow(rowNo);
  return { ok: true, id: id };
}

function readAll_(name) {
  const sheet = getSheet_(name);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const headers = headers_(sheet);
  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(2, 1, last - 1, headers.length).getValues();
  const idIdx = headers.indexOf('id');
  const out = [];
  values.forEach(function (row) {
    if (!row[idIdx]) return;
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = fromCell_(h, row[i], tz); });
    out.push(obj);
  });
  return out;
}

/** ส่งอีเมลแจ้งผู้บริหารเมื่อมีการขออนุมัติซ่อม (ต้องกำหนด EXEC_EMAILS) */
function notifyExec_(subject, body) {
  const to = EXEC_EMAILS.split(',').map(function (x) { return x.trim(); }).filter(String);
  if (!to.length) return { ok: true, sent: 0 };
  MailApp.sendEmail(to.join(','), String(subject || 'ขออนุมัติซ่อมบำรุง'), String(body || ''));
  return { ok: true, sent: to.length };
}

/* ---------- Helpers ---------- */

function getSheet_(name) {
  if (!SHEETS[name]) throw new Error('Unknown collection: ' + name);
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastColumn() === 0) {
    const cols = SHEETS[name];
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8ece7');
    sheet.setFrozenRows(1);
    textFormat_(sheet, 1, cols.length);
  } else {
    ensureColumns_(sheet, SHEETS[name]); // เติมคอลัมน์ที่ขาดให้ชีตเก่า
  }
  return sheet;
}

function headers_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

/** เพิ่มคอลัมน์ที่ยังไม่มี (ต่อท้าย) */
function ensureColumns_(sheet, keys) {
  const have = headers_(sheet);
  const missing = keys.filter(function (k) { return k && have.indexOf(k) === -1; });
  if (!missing.length) return;
  const start = have.length + 1;
  sheet.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#e8ece7');
  textFormat_(sheet, start, missing.length);
}

/** ตั้งคอลัมน์เป็น Plain text เพื่อไม่ให้ Sheets แปลงวันที่/ตัวเลขเอง */
function textFormat_(sheet, col, count) {
  sheet.getRange(1, col, sheet.getMaxRows(), count).setNumberFormat('@');
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2 || !id) return 0;
  const idCol = headers_(sheet).indexOf('id') + 1;
  const ids = sheet.getRange(2, idCol, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function toCell_(key, v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') v = JSON.stringify(v);
  if (typeof v === 'string' && v.length > MAX_CELL) {
    // เกินขีดจำกัดเซลล์: ตัดรูปทิ้ง (ข้อมูลอื่นไม่ควรยาวขนาดนี้)
    return PHOTO_FIELDS.indexOf(key) !== -1 ? '' : v.substring(0, MAX_CELL);
  }
  return v;
}

function fromCell_(key, v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  if (v === '') return key === 'photoBase64' ? null : '';
  if (JSON_FIELDS.indexOf(key) !== -1 && typeof v === 'string') {
    try { return JSON.parse(v); } catch (e) { return {}; }
  }
  return v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
