/**
 * ระบบบันทึกการขอตัวอย่าง (Specimen Request Log)
 * Google Apps Script Web App — backend
 *
 * Spreadsheet: 1-s6pEGwCIo3baZ__GlMxDpMyJNFF8qUnA87rCk7ELY4
 *   Sheet "Data"          → ฐานข้อมูลหลัก (สร้างอัตโนมัติโดยฟังก์ชัน setupSheets)
 *   Sheet "dropdownlist"  → แหล่งข้อมูล dropdown (หัวคอลัมน์ = ชื่อฟิลด์, แถวถัดไป = ตัวเลือก)
 *   Sheet "ReportForm"    → เทมเพลตสำหรับพิมพ์ใบรายงาน (ใช้งานภายหลัง)
 */

// ---------- CONFIG ----------
var CONFIG = {
  SHEET_ID: '1-s6pEGwCIo3baZ__GlMxDpMyJNFF8qUnA87rCk7ELY4',
  DATA_SHEET: 'Data',
  DROPDOWN_SHEET: 'dropdownlist',
  REPORT_SHEET: 'ReportForm',
  TIMEZONE: 'Asia/Bangkok'
};

// คอลัมน์ของ Data sheet (1-based) — แก้ที่นี่ที่เดียวถ้าต้องเพิ่ม/ย้ายคอลัมน์
var COL = {
  RECORD_ID: 1,
  REQUEST_DATE: 2,
  REQUEST_DEPT: 3,
  PATH_LAB: 4,
  PATIENT_NAME: 5,
  HN: 6,
  SPECIMEN_LIST: 7,
  PURPOSE: 8,
  STATUS: 9,
  RECEIVE_DATE: 10,
  RECEIVER: 11,
  ACTION_TAKEN: 12,
  PICKUP_PERSON: 13,
  SIGNATURE: 14,
  DISPENSE_DATE: 15,
  NOTES: 16,
  LAST_UPDATE: 17,
  // เพิ่มต่อท้าย (ไม่แทรกกลาง) เพื่อไม่กระทบข้อมูลเดิมที่บันทึกไว้แล้ว
  RECORDER_STEP1: 18,
  RECORDER_STEP3: 19,
  RECEIVE_LOCATION: 20
};

var HEADERS = [
  'RecordID', 'วันที่บันทึกคำขอ', 'หน่วยงานที่ขอ', 'ห้องปฏิบัติการพยาธิฯ',
  'ชื่อ-สกุลผู้ป่วย', 'HN', 'รายการตัวอย่าง', 'เพื่อนำส่ง', 'สถานะ',
  'วันที่รับตัวอย่าง', 'ผู้รับตัวอย่าง', 'การดำเนินการ',
  'ผู้มารับตัวอย่าง', 'ลายเซ็น', 'วันที่จำหน่ายออก', 'หมายเหตุ', 'อัปเดตล่าสุด',
  'ผู้บันทึกคำขอ', 'ผู้บันทึกการจำหน่าย', 'สถานที่รับ'
];

var STATUS = {
  PENDING: 'รอรับตัวอย่าง',
  RECEIVED: 'รับตัวอย่างแล้ว',
  DISPENSED: 'จำหน่ายแล้ว'
};

// คอลัมน์ใน sheet "dropdownlist" พร้อมตัวอย่างค่าเริ่มต้น
// เพิ่มรายการที่นี่ถ้าต้องการ dropdown ใหม่ แล้ว Run setupSheets ซ้ำเพื่อสร้างคอลัมน์ให้
var DROPDOWN_COLUMNS = [
  { name: 'ห้องปฏิบัติการพยาธิฯ', samples: ['พยาธิวิทยากายวิภาค', 'พยาธิวิทยาคลินิก', 'นิติเวช'] },
  { name: 'หน่วยงานที่ขอ', samples: ['หอผู้ป่วยศัลยกรรม', 'ห้องฉุกเฉิน'] },
  { name: 'ผู้บันทึก', samples: ['เจ้าหน้าที่ A', 'เจ้าหน้าที่ B'] },
  { name: 'เพื่อนำส่ง', samples: ['ส่งตรวจเพิ่มเติม', 'ขอความเห็นที่สอง', 'ใช้เพื่อการวิจัย'] },
  { name: 'ผู้รับตัวอย่าง', samples: ['เจ้าหน้าที่ A', 'เจ้าหน้าที่ B'] },
  { name: 'สถานที่รับ', samples: ['ห้องปฏิบัติการพยาธิวิทยากายวิภาค', 'ห้องปฏิบัติการพยาธิวิทยาคลินิก'] }
];

// ---------- WEB APP ENTRY ----------
function doGet(e) {
  var tpl = HtmlService.createTemplateFromFile('Index');
  return tpl.evaluate()
    .setTitle('บันทึกการขอตัวอย่าง')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- ONE-TIME SETUP ----------
// เรียกฟังก์ชันนี้ครั้งเดียวจาก Apps Script editor (Run > setupSheets) เพื่อสร้างหัวตาราง
function setupSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  var data = ss.getSheetByName(CONFIG.DATA_SHEET) || ss.insertSheet(CONFIG.DATA_SHEET);
  if (data.getLastRow() === 0) {
    data.appendRow(HEADERS);
    data.setFrozenRows(1);
    data.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    // migration: เติมหัวคอลัมน์ที่ขาด (เช่น ผู้บันทึกคำขอ/รับ/จำหน่าย) ต่อท้ายโดยไม่แตะข้อมูลเดิม
    var existingHeaders = data.getRange(1, 1, 1, data.getLastColumn()).getValues()[0];
    if (existingHeaders.length < HEADERS.length) {
      var missing = HEADERS.slice(existingHeaders.length);
      data.getRange(1, existingHeaders.length + 1, 1, missing.length)
        .setValues([missing]).setFontWeight('bold');
    }
  }

  ensureDropdownColumns_(ss.getSheetByName(CONFIG.DROPDOWN_SHEET) || ss.insertSheet(CONFIG.DROPDOWN_SHEET));

  if (!ss.getSheetByName(CONFIG.REPORT_SHEET)) {
    ss.insertSheet(CONFIG.REPORT_SHEET);
  }
}

// เติมคอลัมน์ dropdown ที่ยังไม่มีต่อท้าย โดยไม่แตะคอลัมน์/ตัวเลือกที่ผู้ใช้แก้ไว้แล้ว
function ensureDropdownColumns_(dd) {
  var lastCol = dd.getLastColumn();
  var existing = lastCol > 0 ? dd.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  DROPDOWN_COLUMNS.forEach(function (col) {
    if (existing.indexOf(col.name) !== -1) return;
    var target = dd.getLastColumn() + 1;
    dd.getRange(1, target).setValue(col.name).setFontWeight('bold');
    dd.getRange(2, target, col.samples.length, 1).setValues(col.samples.map(function (v) {
      return [v];
    }));
  });
}

// ---------- DEBUG (ชั่วคราว) ----------
// รันฟังก์ชันนี้จาก Apps Script editor โดยตรง (เลือกจาก dropdown ด้านบน > Run)
// แล้วดูผลลัพธ์ที่ View > Logs (หรือ Ctrl+Enter) เพื่อตรวจว่าโค้ดอ่านชีตถูกต้องหรือไม่
function debugCheckData() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  Logger.log('Spreadsheet name: ' + ss.getName());
  Logger.log('Spreadsheet URL: ' + ss.getUrl());

  var sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!sheet) {
    Logger.log('!! ไม่พบชีตชื่อ "' + CONFIG.DATA_SHEET + '" ในสเปรดชีตนี้');
    return;
  }

  Logger.log('lastRow: ' + sheet.getLastRow() + ', lastColumn: ' + sheet.getLastColumn());
  Logger.log('แถวที่ 1 (ที่โค้ดถือว่าเป็นหัวตาราง): ' + JSON.stringify(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]));

  if (sheet.getLastRow() >= 2) {
    Logger.log('แถวที่ 2 (ที่โค้ดถือว่าเป็นข้อมูลจริงแถวแรก): ' + JSON.stringify(sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]));
    Logger.log('ค่าคอลัมน์สถานะ (COL.STATUS = ' + COL.STATUS + ') ของแถวที่ 2: "' + sheet.getRange(2, COL.STATUS).getValue() + '"');
  } else {
    Logger.log('!! lastRow < 2 แปลว่าโค้ดมองว่าไม่มีข้อมูลเลย (มีแค่แถวเดียวหรือว่างเปล่า)');
  }
}

// ---------- DROPDOWN LISTS ----------
// อ่านทุกคอลัมน์ใน sheet "dropdownlist" แล้วคืนค่าเป็น { หัวคอลัมน์: [ตัวเลือก, ...] }
// เพิ่ม/แก้ตัวเลือกได้จากในชีตโดยตรง ไม่ต้องแก้โค้ด
function getDropdownLists() {
  var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DROPDOWN_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var result = {};
  headers.forEach(function (header, colIdx) {
    if (!header) return;
    var options = [];
    for (var r = 1; r < values.length; r++) {
      var v = values[r][colIdx];
      if (v !== '' && v !== null && v !== undefined) options.push(String(v));
    }
    result[header] = options;
  });
  return result;
}

// ---------- HELPERS ----------
function getDataSheet_() {
  return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
}

// รองรับค่าจากชีตได้ทุกชนิด (Date / ข้อความ / ตัวเลข / ว่าง) โดยไม่ throw
// เซลล์วันที่ที่พิมพ์มือหรือผิดรูปแบบเพียงเซลล์เดียว เคยทำให้ Dashboard ทั้งหน้าพัง
function formatDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  }
  return String(value);
}

function todayStamp_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
}

// สร้างรหัสอ้างอิง เช่น SP-20260909-001 (รันตามลำดับของวันนั้น)
function generateRecordId_(sheet) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var prefix = 'SP-' + todayStamp_() + '-';
    var lastRow = sheet.getLastRow();
    var count = 0;
    if (lastRow > 1) {
      var ids = sheet.getRange(2, COL.RECORD_ID, lastRow - 1, 1).getValues();
      ids.forEach(function (row) {
        if (String(row[0]).indexOf(prefix) === 0) count++;
      });
    }
    var seq = ('000' + (count + 1)).slice(-3);
    return prefix + seq;
  } finally {
    lock.releaseLock();
  }
}

function findRowByRecordId_(sheet, recordId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, COL.RECORD_ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === recordId) return i + 2; // 1-based sheet row
  }
  return -1;
}

function rowToRecord_(row) {
  return {
    recordId: row[COL.RECORD_ID - 1],
    requestDate: formatDate_(row[COL.REQUEST_DATE - 1]),
    requestDept: row[COL.REQUEST_DEPT - 1],
    pathLab: row[COL.PATH_LAB - 1],
    patientName: row[COL.PATIENT_NAME - 1],
    hn: row[COL.HN - 1],
    specimenList: row[COL.SPECIMEN_LIST - 1],
    purpose: row[COL.PURPOSE - 1],
    status: row[COL.STATUS - 1],
    receiveDate: formatDate_(row[COL.RECEIVE_DATE - 1]),
    receiver: row[COL.RECEIVER - 1],
    actionTaken: row[COL.ACTION_TAKEN - 1],
    pickupPerson: row[COL.PICKUP_PERSON - 1],
    hasSignature: !!row[COL.SIGNATURE - 1],
    dispenseDate: formatDate_(row[COL.DISPENSE_DATE - 1]),
    notes: row[COL.NOTES - 1],
    lastUpdate: formatDate_(row[COL.LAST_UPDATE - 1]),
    recorderStep1: row[COL.RECORDER_STEP1 - 1],
    recorderStep3: row[COL.RECORDER_STEP3 - 1],
    receiveLocation: row[COL.RECEIVE_LOCATION - 1]
  };
}

// ---------- STEP 1: บันทึกคำขอ ----------
function submitStep1(payload) {
  if (!payload.patientName || !payload.hn || !payload.recorderStep1) {
    throw new Error('กรุณากรอกชื่อ-สกุล, HN และผู้บันทึก');
  }
  var sheet = getDataSheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var recordId;
  try {
    recordId = generateRecordId_(sheet);
    var now = new Date();
    var row = [];
    row[COL.RECORD_ID - 1] = recordId;
    row[COL.REQUEST_DATE - 1] = now;
    row[COL.REQUEST_DEPT - 1] = payload.requestDept || '';
    row[COL.PATH_LAB - 1] = payload.pathLab || '';
    row[COL.PATIENT_NAME - 1] = payload.patientName;
    row[COL.HN - 1] = payload.hn;
    row[COL.SPECIMEN_LIST - 1] = payload.specimenList || '';
    row[COL.PURPOSE - 1] = payload.purpose || '';
    row[COL.STATUS - 1] = STATUS.PENDING;
    row[COL.NOTES - 1] = payload.notes || '';
    row[COL.LAST_UPDATE - 1] = now;
    row[COL.RECORDER_STEP1 - 1] = payload.recorderStep1;
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { recordId: recordId };
}

// ---------- STEP 2 / STEP 3: ค้นหา ----------
// step: 'step2' -> เฉพาะสถานะ "รอรับตัวอย่าง", 'step3' -> เฉพาะ "รับตัวอย่างแล้ว"
function searchRecords(step, keyword) {
  var targetStatus = step === 'step2' ? STATUS.PENDING : STATUS.RECEIVED;
  var sheet = getDataSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var kw = (keyword || '').toString().trim().toLowerCase();

  var results = [];
  values.forEach(function (row) {
    if (row[COL.STATUS - 1] !== targetStatus) return;
    var hn = String(row[COL.HN - 1] || '').toLowerCase();
    var name = String(row[COL.PATIENT_NAME - 1] || '').toLowerCase();
    if (kw === '' || hn.indexOf(kw) !== -1 || name.indexOf(kw) !== -1) {
      results.push(rowToRecord_(row));
    }
  });
  return results;
}

// ---------- STEP 2: รับตัวอย่าง ----------
function submitStep2(recordId, payload) {
  var sheet = getDataSheet_();
  var rowIdx = findRowByRecordId_(sheet, recordId);
  if (rowIdx === -1) throw new Error('ไม่พบรายการที่เลือก (RecordID: ' + recordId + ')');

  var now = new Date();
  sheet.getRange(rowIdx, COL.RECEIVE_DATE).setValue(now);
  sheet.getRange(rowIdx, COL.RECEIVER).setValue(payload.receiver || '');
  sheet.getRange(rowIdx, COL.RECEIVE_LOCATION).setValue(payload.receiveLocation || '');
  sheet.getRange(rowIdx, COL.ACTION_TAKEN).setValue(payload.actionTaken || '');
  sheet.getRange(rowIdx, COL.STATUS).setValue(STATUS.RECEIVED);
  sheet.getRange(rowIdx, COL.LAST_UPDATE).setValue(now);
  if (payload.notes) {
    appendNote_(sheet, rowIdx, payload.notes);
  }
  return { ok: true };
}

// ---------- STEP 3: จำหน่ายออก ----------
function submitStep3(recordId, payload) {
  if (!payload.signature) {
    throw new Error('กรุณาลงลายเซ็นก่อนบันทึก');
  }
  var sheet = getDataSheet_();
  var rowIdx = findRowByRecordId_(sheet, recordId);
  if (rowIdx === -1) throw new Error('ไม่พบรายการที่เลือก (RecordID: ' + recordId + ')');

  var now = new Date();
  sheet.getRange(rowIdx, COL.PICKUP_PERSON).setValue(payload.pickupPerson || '');
  sheet.getRange(rowIdx, COL.RECORDER_STEP3).setValue(payload.recorderStep3 || '');
  sheet.getRange(rowIdx, COL.SIGNATURE).setValue(payload.signature); // dataURL (base64 PNG)
  sheet.getRange(rowIdx, COL.DISPENSE_DATE).setValue(now);
  sheet.getRange(rowIdx, COL.STATUS).setValue(STATUS.DISPENSED);
  sheet.getRange(rowIdx, COL.LAST_UPDATE).setValue(now);
  if (payload.notes) {
    appendNote_(sheet, rowIdx, payload.notes);
  }
  return { ok: true };
}

function appendNote_(sheet, rowIdx, newNote) {
  var cell = sheet.getRange(rowIdx, COL.NOTES);
  var existing = cell.getValue();
  var stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM HH:mm');
  var combined = (existing ? existing + '\n' : '') + '[' + stamp + '] ' + newNote;
  cell.setValue(combined);
}

// ดึงลายเซ็น (แยกออกจากรายการหลักเพื่อลดขนาดข้อมูลตอนโหลดตาราง/ค้นหา)
function getSignature(recordId) {
  var sheet = getDataSheet_();
  var rowIdx = findRowByRecordId_(sheet, recordId);
  if (rowIdx === -1) return '';
  return sheet.getRange(rowIdx, COL.SIGNATURE).getValue();
}

// ---------- DASHBOARD ----------
function getDashboardData(filter) {
  filter = filter || {};
  var sheet = getDataSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { summary: { total: 0, pending: 0, received: 0, dispensed: 0 }, records: [] };
  }

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var kw = (filter.keyword || '').toString().trim().toLowerCase();

  var summary = { total: 0, pending: 0, received: 0, dispensed: 0 };
  var records = [];

  values.forEach(function (row) {
    var status = row[COL.STATUS - 1];
    if (!status) return;

    summary.total++;
    if (status === STATUS.PENDING) summary.pending++;
    else if (status === STATUS.RECEIVED) summary.received++;
    else if (status === STATUS.DISPENSED) summary.dispensed++;

    if (filter.status && filter.status !== 'ALL' && status !== filter.status) return;

    if (kw) {
      var hn = String(row[COL.HN - 1] || '').toLowerCase();
      var name = String(row[COL.PATIENT_NAME - 1] || '').toLowerCase();
      var dept = String(row[COL.REQUEST_DEPT - 1] || '').toLowerCase();
      if (hn.indexOf(kw) === -1 && name.indexOf(kw) === -1 && dept.indexOf(kw) === -1) return;
    }

    if (filter.dateFrom) {
      var reqDate = row[COL.REQUEST_DATE - 1];
      if (reqDate && new Date(reqDate) < new Date(filter.dateFrom)) return;
    }
    if (filter.dateTo) {
      var reqDate2 = row[COL.REQUEST_DATE - 1];
      if (reqDate2 && new Date(reqDate2) > new Date(filter.dateTo + 'T23:59:59')) return;
    }

    records.push(rowToRecord_(row));
  });

  records.sort(function (a, b) {
    return String(b.recordId).localeCompare(String(a.recordId));
  });

  return { summary: summary, records: records };
}

// ---------- EDIT ----------
// แก้ไขข้อมูลคำขอ (ฟิลด์ของขั้นตอนที่ 1) จาก Dashboard ได้โดยไม่ต้องยุ่งกับสถานะ/ขั้นตอนถัดไป
function updateRecordFields(recordId, payload) {
  if (!payload.patientName || !payload.hn || !payload.pathLab || !payload.recorderStep1) {
    throw new Error('กรุณากรอกข้อมูลที่จำเป็น (*) ให้ครบ');
  }
  var sheet = getDataSheet_();
  var rowIdx = findRowByRecordId_(sheet, recordId);
  if (rowIdx === -1) throw new Error('ไม่พบรายการที่เลือก (RecordID: ' + recordId + ')');

  sheet.getRange(rowIdx, COL.REQUEST_DEPT).setValue(payload.requestDept || '');
  sheet.getRange(rowIdx, COL.PATH_LAB).setValue(payload.pathLab);
  sheet.getRange(rowIdx, COL.PATIENT_NAME).setValue(payload.patientName);
  sheet.getRange(rowIdx, COL.HN).setValue(payload.hn);
  sheet.getRange(rowIdx, COL.SPECIMEN_LIST).setValue(payload.specimenList || '');
  sheet.getRange(rowIdx, COL.PURPOSE).setValue(payload.purpose || '');
  sheet.getRange(rowIdx, COL.RECORDER_STEP1).setValue(payload.recorderStep1);
  sheet.getRange(rowIdx, COL.NOTES).setValue(payload.notes || '');
  sheet.getRange(rowIdx, COL.LAST_UPDATE).setValue(new Date());
  return { ok: true };
}

// ---------- REVERT ----------
// ย้อนสถานะจาก "จำหน่ายแล้ว" กลับเป็น "รับตัวอย่างแล้ว" เผื่อบันทึกขั้นตอน 3 ผิดพลาด
// ล้างข้อมูลเฉพาะของขั้นตอน 3 ทิ้ง เพื่อให้กลับไปทำขั้นตอน 3 ใหม่ได้สะอาด
function revertFromDispensed(recordId) {
  var sheet = getDataSheet_();
  var rowIdx = findRowByRecordId_(sheet, recordId);
  if (rowIdx === -1) throw new Error('ไม่พบรายการที่เลือก (RecordID: ' + recordId + ')');

  var currentStatus = sheet.getRange(rowIdx, COL.STATUS).getValue();
  if (currentStatus !== STATUS.DISPENSED) {
    throw new Error('รายการนี้ไม่ได้อยู่ในสถานะ "จำหน่ายแล้ว"');
  }

  sheet.getRange(rowIdx, COL.STATUS).setValue(STATUS.RECEIVED);
  sheet.getRange(rowIdx, COL.PICKUP_PERSON).setValue('');
  sheet.getRange(rowIdx, COL.SIGNATURE).setValue('');
  sheet.getRange(rowIdx, COL.DISPENSE_DATE).setValue('');
  sheet.getRange(rowIdx, COL.RECORDER_STEP3).setValue('');
  sheet.getRange(rowIdx, COL.LAST_UPDATE).setValue(new Date());
  return { ok: true };
}
