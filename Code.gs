/**
 * QSNCC — ระบบบันทึกการซ่อมเครื่องฆ่าเชื้อ (Backend)
 * Google Apps Script — bound to the existing "QSNCC Disinfection Machine Registry" spreadsheet.
 *
 * Endpoints
 *   doGet()                         → master data (zones/types/symptoms/models/statuses/points+machines)
 *   doGet(?action=history&machineId=X) → repair history for one machine
 *   doGet(?action=ping)             → health check
 *   doPost() {records:[...]}        → batch-write repair rows into "บันทึกการซ่อม"
 *
 * Design notes
 * - All reads/writes are HEADER-DRIVEN: columns are located by their Thai header text,
 *   so the script keeps working if columns are reordered, and future master symptoms
 *   automatically map to their per-symptom checkbox column.
 * - Existing report/summary sheets keep working because we fill the original columns.
 *   Two extra columns ("อาการที่ซ่อมเสร็จ (รหัส+ชื่อ)", "อาการที่ยังค้าง (รหัส+ชื่อ)")
 *   are appended once if missing, so repaired/pending detail is preserved without
 *   disturbing existing formulas.
 */

/* ------------------------------------------------------------------ config */
var SHEETS = {
  REGISTRY: 'ทะเบียนเครื่อง',
  MASTER:   'ข้อมูลหลัก',
  LOG:      'บันทึกการซ่อม'
};

// Statuses offered in the app (master-driven; edit here or expose from a sheet later).
var STATUS_LIST = ['เสร็จสิ้น', 'เสร็จบางส่วน', 'รออะไหล่', 'ส่งซ่อมภายนอก', 'ยกเลิก'];

// Columns the app guarantees exist in the repair log (created once if missing).
// "ผู้บันทึก" already exists in the source sheet — kept here as a safety net so the
// recorder is always trackable even on a sheet variant that lacks it. The two symptom
// columns capture app-computed detail the original schema had no place for.
var ENSURE_COLS = ['ผู้บันทึก', 'อาการที่ซ่อมเสร็จ (รหัส+ชื่อ)', 'อาการที่ยังค้าง (รหัส+ชื่อ)'];

/* ------------------------------------------------------------------ router */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'ping')    return json({ ok: true, ts: new Date().toISOString() });
    if (action === 'history') return json({ ok: true, history: getHistory(e.parameter.machineId) });
    if (action === 'master')  return json({ ok: true, data: getMasterData() });
    // Default: serve the app itself. Same-origin as the backend → no CORS, and the
    // client talks to the server via google.script.run (see index.html). This is the
    // only reliable pattern when the Workspace forbids anonymous ("Anyone") web apps.
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('QSNCC · บันทึกการซ่อมเครื่องฆ่าเชื้อ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var result = saveRecords(body);
    return json({ ok: true, result: result });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

/* -------------------------------------------------------------- master data */
function getMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  /* registry → points + machines */
  var reg = readSheet(ss, SHEETS.REGISTRY);
  var zones = {}, types = {}, pointsMap = {};
  var typeOrder = { S: 0, U: 1, F: 2, D: 3 };

  reg.rows.forEach(function (r) {
    var mid = r['Machine ID']; if (!mid) return;
    var z = r['โซน'], zname = r['ชื่อโซน'], pcode = r['รหัสจุด'];
    var tcode = r['รหัสประเภท'], tname = r['ประเภทสุขภัณฑ์'];
    if (z) zones[z] = zname;
    if (tcode) types[tcode] = tname;
    if (!pcode) return;
    var p = pointsMap[pcode] || (pointsMap[pcode] = {
      code: pcode, zone: z, zname: zname,
      sub: r['ชั้น/โซนย่อย'] || '', remark: r['ตำแหน่ง/หมายเหตุ'] || '', machines: []
    });
    p.machines.push({
      id: String(mid), type: tcode, num: cleanNum(r['เลขข้างเครื่อง']),
      model: r['รุ่น'] || '', status: r['สถานะ'] || ''
    });
  });

  var points = Object.keys(pointsMap).map(function (k) { return pointsMap[k]; });
  points.forEach(function (p) {
    p.machines.sort(function (a, b) {
      return (typeOrder[a.type] || 9) - (typeOrder[b.type] || 9) ||
             (parseInt(a.num, 10) || 999) - (parseInt(b.num, 10) || 999);
    });
  });
  points.sort(function (a, b) {
    return String(a.zone).localeCompare(String(b.zone)) || String(a.code).localeCompare(String(b.code));
  });

  /* master sheet → symptoms + models (header-agnostic block scan) */
  var symptoms = readSymptoms(ss);
  var models = readModels(ss);

  return {
    zones:    Object.keys(zones).map(function (k) { return { code: k, name: zones[k] }; }),
    types:    Object.keys(types).map(function (k) { return { code: k, name: types[k] }; }),
    symptoms: symptoms.length ? symptoms : defaultSymptoms(),
    models:   models.length ? models : ['รุ่น-1'],
    statuses: STATUS_LIST,
    points:   points,
    meta: { machineCount: reg.rows.length, pointCount: points.length, ts: new Date().toISOString() }
  };
}

function readSymptoms(ss) {
  // "ข้อมูลหลัก" holds a block: header row has "รหัส" | "อาการ". Scan for that pair.
  var sh = ss.getSheetByName(SHEETS.MASTER);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  var out = [], codeCol = -1, nameCol = -1;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === 'รหัส' &&
          String(values[r][c + 1] || '').trim() === 'อาการ') { codeCol = c; nameCol = c + 1; break; }
    }
    if (codeCol > -1) {
      for (var rr = r + 1; rr < values.length; rr++) {
        var code = values[rr][codeCol], name = values[rr][nameCol];
        if (code === '' || code == null || !name) break;
        out.push({ code: cleanNum(code), name: String(name).trim() });
      }
      break;
    }
  }
  return out;
}

function readModels(ss) {
  var sh = ss.getSheetByName(SHEETS.MASTER);
  if (!sh) return [];
  var values = sh.getDataRange().getValues(), out = [];
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === 'รุ่น' && String(values[r][1] || '').indexOf('รายละเอียด') === 0) {
      for (var rr = r + 1; rr < values.length; rr++) {
        var v = values[rr][0];
        if (!v || String(v).indexOf('รหัสโซน') === 0) break;
        out.push(String(v).trim());
      }
      break;
    }
  }
  return out;
}

function defaultSymptoms() {
  return [
    { code: '1', name: 'หัววงจรเสีย' }, { code: '2', name: 'รางแบตเตอรี่ชำรุด' },
    { code: '3', name: 'ฝาหน้า/ฝาหลัง' }, { code: '4', name: 'น้ำยารั่ว' },
    { code: '5', name: 'ตัวเครื่อง' }, { code: '6', name: 'ท่อหลุด' }
  ];
}

/* --------------------------------------------------------------- write path */
function saveRecords(body) {
  var records = body.records || [];
  if (!records.length) return { written: 0, skipped: 0, seq: [] };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.LOG);
  if (!sh) throw new Error('ไม่พบชีต "' + SHEETS.LOG + '"');

  ensureExtraColumns(sh);
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  header.forEach(function (h, i) { idx[h] = i; });

  var typeCounts = pointTypeCounts(ss);      // {pcode:{S,U,F,D}}
  var existingKeys = existingDupKeys(sh, idx); // Set of "date|machineId"

  var lastSeq = getLastSeq(sh, idx);
  var newRows = [], seqOut = [], skipped = 0;

  records.forEach(function (rec) {
    var date = rec.date;                                   // 'YYYY-MM-DD'
    var dupKey = date + '|' + rec.machineId;
    if (existingKeys[dupKey] && !rec.force) { skipped++; return; }
    existingKeys[dupKey] = true;

    lastSeq++;
    var row = new Array(header.length).fill('');
    var d = new Date(date + 'T00:00:00');
    var month = date.slice(0, 7);
    var quarter = date.slice(0, 4) + '-Q' + (Math.floor(d.getMonth() / 3) + 1);
    var round = rec.round || autoRound(d);
    var tc = typeCounts[rec.pointCode] || {};

    var found    = rec.found    || [];   // [{code,name}]
    var repaired = rec.repaired || [];
    var pending  = rec.pending  || [];

    set(row, idx, 'ลำดับ', lastSeq);
    set(row, idx, 'วันที่', date);
    set(row, idx, 'รอบ', round);
    set(row, idx, 'รหัสจุด', rec.pointCode);
    set(row, idx, 'เลขเครื่อง', rec.num);
    set(row, idx, 'ประเภท (เลือกเมื่อเลขซ้ำ)', rec.typeName);
    set(row, idx, 'ประเภทที่ใช้', rec.typeName);
    set(row, idx, 'Machine ID', rec.machineId);
    set(row, idx, 'ตรวจสอบ', 'OK');
    set(row, idx, 'โซน', rec.zone);
    set(row, idx, 'ชื่อโซน', rec.zoneName);
    set(row, idx, 'รหัสอาการ', found.map(function (s) { return s.code; }).join(','));
    // per-symptom checkbox columns (header text == symptom name)
    found.forEach(function (s) { if (idx[s.name] != null) row[idx[s.name]] = '✓'; });
    set(row, idx, 'อาการเสีย', found.map(function (s) { return s.name; }).join(', '));
    set(row, idx, 'สรุปอาการ (รหัส+ชื่อ)', symText(found));
    set(row, idx, 'รายละเอียดอาการ/การซ่อม', rec.detail || '');
    set(row, idx, 'สถานะ', rec.status);
    set(row, idx, 'ผู้บันทึก', rec.recorder);
    set(row, idx, 'หมายเหตุ', rec.note || '');
    set(row, idx, 'เดือน', month);
    set(row, idx, 'ไตรมาส', quarter);
    set(row, idx, 'ปี', date.slice(0, 4));
    set(row, idx, 'คีย์ช่วงเวลา', round);
    set(row, idx, 'cS', tc.S || ''); set(row, idx, 'cU', tc.U || '');
    set(row, idx, 'cF', tc.F || ''); set(row, idx, 'cD', tc.D || '');
    // appended columns
    set(row, idx, 'อาการที่ซ่อมเสร็จ (รหัส+ชื่อ)', symText(repaired));
    set(row, idx, 'อาการที่ยังค้าง (รหัส+ชื่อ)', symText(pending));

    newRows.push(row);
    seqOut.push({ machineId: rec.machineId, seq: lastSeq });
  });

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, header.length).setValues(newRows);
  }
  return { written: newRows.length, skipped: skipped, seq: seqOut };
}

function symText(list) {
  return (list || []).map(function (s) { return s.code + '  |  ' + s.name; }).join('\n');
}

function ensureExtraColumns(sh) {
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  ENSURE_COLS.forEach(function (name) {
    if (header.indexOf(name) === -1) {
      var col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(name);
      header.push(name);
    }
  });
}

/* ------------------------------------------------------------------ history */
function getHistory(machineId) {
  if (!machineId) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.LOG);
  var values = sh.getDataRange().getValues();
  var header = values[0].map(function (h) { return String(h).trim(); });
  var i = {}; header.forEach(function (h, k) { i[h] = k; });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][i['Machine ID']]) !== String(machineId)) continue;
    out.push({
      date:     fmtDate(values[r][i['วันที่']]),
      round:    values[r][i['รอบ']],
      found:    values[r][i['สรุปอาการ (รหัส+ชื่อ)']],
      repaired: i['อาการที่ซ่อมเสร็จ (รหัส+ชื่อ)'] != null ? values[r][i['อาการที่ซ่อมเสร็จ (รหัส+ชื่อ)']] : '',
      pending:  i['อาการที่ยังค้าง (รหัส+ชื่อ)']  != null ? values[r][i['อาการที่ยังค้าง (รหัส+ชื่อ)']]  : '',
      status:   values[r][i['สถานะ']],
      recorder: values[r][i['ผู้บันทึก']],
      note:     values[r][i['หมายเหตุ']]
    });
  }
  out.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  return out;
}

/* ------------------------------------------------------------------ helpers */
function readSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '"');
  var values = sh.getDataRange().getValues();
  var header = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r].every(function (c) { return c === '' || c == null; })) continue;
    var o = {};
    header.forEach(function (h, c) { o[h] = values[r][c]; });
    rows.push(o);
  }
  return { header: header, rows: rows };
}

function pointTypeCounts(ss) {
  var reg = readSheet(ss, SHEETS.REGISTRY);
  var m = {};
  reg.rows.forEach(function (r) {
    var p = r['รหัสจุด'], t = r['รหัสประเภท']; if (!p) return;
    (m[p] || (m[p] = {}))[t] = (m[p][t] || 0) + 1;
  });
  return m;
}

function existingDupKeys(sh, idx) {
  var keys = {};
  var last = sh.getLastRow();
  if (last < 2) return keys;
  var dcol = idx['วันที่'] + 1, mcol = idx['Machine ID'] + 1;
  var dates = sh.getRange(2, dcol, last - 1, 1).getValues();
  var mids  = sh.getRange(2, mcol, last - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    keys[fmtDate(dates[i][0]) + '|' + mids[i][0]] = true;
  }
  return keys;
}

function getLastSeq(sh, idx) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var col = idx['ลำดับ'] + 1;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var max = 0;
  vals.forEach(function (v) { var n = parseInt(v[0], 10); if (n > max) max = n; });
  return max;
}

function autoRound(d) {
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  return d.getFullYear() + '-' + m + '-R' + (d.getDate() <= 15 ? 1 : 2);
}

function set(row, idx, header, val) { if (idx[header] != null) row[idx[header]] = val; }

function cleanNum(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v % 1 === 0 ? String(v) : String(v);
  return String(v).replace(/\.0$/, '');
}

function fmtDate(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  return String(v).slice(0, 10);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
