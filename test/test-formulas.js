const { makeSheet, buildEnv, loadCode } = require('./gas-harness');

// Real header from the workbook
const LOG_HEADER = ['ลำดับ','วันที่','รอบ','รหัสจุด','เลขเครื่อง','ประเภท (เลือกเมื่อเลขซ้ำ)','ประเภทที่ใช้',
'Machine ID','ตรวจสอบ','โซน','ชื่อโซน','รหัสอาการ','หัววงจรเสีย','รางแบตเตอรี่ชำรุด','ฝาหน้า/ฝาหลัง','น้ำยารั่ว',
'ตัวเครื่อง','ท่อหลุด','(สำรอง 7)','(สำรอง 8)','(สำรอง 9)','(สำรอง 10)','อาการเสีย','สรุปอาการ (รหัส+ชื่อ)',
'รายละเอียดอาการ/การซ่อม','สถานะ','ผู้บันทึก','หมายเหตุ','เดือน','ไตรมาส','ปี','คีย์ช่วงเวลา','pidx','cS','cU','cF','cD'];

// Columns that are formulas in the real sheet (0-based), per the workbook analysis
const FORMULA_COLS = [0,2,6,7,8,9,10,12,13,14,15,16,17,18,19,20,21,22,23,28,29,30,31,32,33,34,35,36];
const INPUT_COLS   = [1,3,4,5,11,24,25,26,27];

function freshSheets() {
  // one existing data row (row 2), with formulas in the formula columns
  const existing = [ new Array(LOG_HEADER.length).fill('') ];
  existing[0][1]='2026-07-13'; existing[0][3]='T-B201'; existing[0][4]='7';
  existing[0][5]='ชาย (ยืน)'; existing[0][11]='1'; existing[0][25]='รออะไหล่';
  const log = makeSheet('บันทึกการซ่อม', LOG_HEADER, FORMULA_COLS, existing, 300);

  const regHeader = ['Machine ID','โซน','ชื่อโซน','ชั้น/โซนย่อย','รหัสจุด','รหัสประเภท','ประเภทสุขภัณฑ์','เลขข้างเครื่อง','รุ่น','ตำแหน่ง/หมายเหตุ','สถานะ'];
  const regRows = [
    ['R-B101-S1','R','Ratchada','F.B1','R-B101','S','ชาย (นั่ง)','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
    ['R-B101-S2','R','Ratchada','F.B1','R-B101','S','ชาย (นั่ง)','2','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
    ['T-B201-U7','T','Thaibev','F.B2','T-B201','U','ชาย (ยืน)','7','รุ่น-1','Locker NCC','ใช้งาน'],
  ];
  const registry = makeSheet('ทะเบียนเครื่อง', regHeader, [], regRows, 50);

  const masterHeader = ['รุ่น','รายละเอียด/หมายเหตุ','','รหัส','อาการ'];
  const masterRows = [
    ['รุ่น-1','ปัจจุบันมีรุ่นเดียว','','1','หัววงจรเสีย'],
    ['','','','2','รางแบตเตอรี่ชำรุด'],
    ['','','','3','ฝาหน้า/ฝาหลัง'],
    ['','','','4','น้ำยารั่ว'],
    ['','','','5','ตัวเครื่อง'],
    ['','','','6','ท่อหลุด'],
  ];
  const master = makeSheet('ข้อมูลหลัก', masterHeader, [], masterRows, 50);
  return { 'บันทึกการซ่อม': log, 'ทะเบียนเครื่อง': registry, 'ข้อมูลหลัก': master };
}

function rec(over) {
  return Object.assign({
    clientId:'ob_1', date:'2026-07-30', round:'2026-07-R2', recorder:'ช่างเอ',
    pointCode:'R-B101', zone:'R', zoneName:'Ratchada', machineId:'R-B101-S1', num:'1',
    typeName:'ชาย (นั่ง)', found:[{code:'1',name:'หัววงจรเสีย'}], repaired:[], pending:[{code:'1',name:'หัววงจรเสีย'}],
    status:'รออะไหล่', detail:'', note:''
  }, over||{});
}

function report(title, pass, extra) {
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${title}${extra ? '  — ' + extra : ''}`);
  if (!pass) process.exitCode = 1;
}

// ---------------------------------------------------------------- TEST 1
{
  const sheets = freshSheets();
  const env = buildEnv(sheets);
  const api = loadCode(env);
  const log = sheets['บันทึกการซ่อม'];

  const res = api.saveRecords({ records: [rec()] });
  const newRow = 3; // row 2 is existing data, so the append lands on row 3

  const formulasIntact = FORMULA_COLS.every(c => log._cells[newRow-1][c].f !== '');
  report('formula columns keep their formulas after an append', formulasIntact,
    formulasIntact ? `${FORMULA_COLS.length} formula cols intact on row ${newRow}`
                   : 'wiped: ' + FORMULA_COLS.filter(c=>!log._cells[newRow-1][c].f).map(c=>LOG_HEADER[c]).join(', '));

  const dateWritten = log._cells[newRow-1][1].v === '2026-07-30';
  const pointWritten = log._cells[newRow-1][3].v === 'R-B101';
  const codesWritten = String(log._cells[newRow-1][11].v) === '1';
  const recorderWritten = log._cells[newRow-1][26].v === 'ช่างเอ';
  report('input columns are written', dateWritten && pointWritten && codesWritten && recorderWritten,
    `วันที่=${log._cells[newRow-1][1].v} รหัสจุด=${log._cells[newRow-1][3].v} รหัสอาการ=${log._cells[newRow-1][11].v} ผู้บันทึก=${log._cells[newRow-1][26].v}`);

  report('append reported written=1', res.written === 1 && res.results[0].ok === true);

  // the ✓ marks must NOT be written as values — the sheet's formulas derive them from รหัสอาการ
  const checkboxUntouched = [12,13,14,15,16,17].every(c => log._cells[newRow-1][c].f !== '');
  report('per-symptom ✓ columns left to their formulas (not double-written)', checkboxUntouched);
}

// ---------------------------------------------------------------- TEST 2
{
  const sheets = freshSheets();
  const env = buildEnv(sheets);
  const api = loadCode(env);
  const log = sheets['บันทึกการซ่อม'];

  api.saveRecords({ records: [rec()] });
  const r = api.updateRecord(rec({ status:'เสร็จสิ้น', repaired:[{code:'1',name:'หัววงจรเสีย'}], pending:[] }));
  const row = r.row;
  const intact = FORMULA_COLS.every(c => log._cells[row-1][c].f !== '');
  report('formulas survive an edit-in-place too', intact && r.mode === 'updated',
    `mode=${r.mode} row=${row}`);
  report('edited status value landed', log._cells[row-1][25].v === 'เสร็จสิ้น');
}

// ---------------------------------------------------------------- TEST 3
{
  const sheets = freshSheets();
  const env = buildEnv(sheets);
  const api = loadCode(env);
  const log = sheets['บันทึกการซ่อม'];

  const before = log.getMaxRows();
  const res = api.fillFormulasDown({});
  // every formula column should now have a formula on the very last row of the sheet
  const lastRow = log.getMaxRows();
  const bottomFilled = FORMULA_COLS.every(c => log._cells[lastRow-1][c].f !== '');
  report('fillFormulasDown reaches the last row of the sheet', bottomFilled,
    `cols=${res.formulaColumns} rows ${res.fromRow}-${res.toRow} of ${before}`);

  // and it must not have touched the existing data row's values
  const existingIntact = log._cells[1][1].v === '2026-07-13' && log._cells[1][3].v === 'T-B201';
  report('fillFormulasDown leaves existing data values alone', existingIntact);

  // a save AFTER filling should still work and still not clobber formulas
  const res2 = api.saveRecords({ records: [rec({ clientId:'ob_2' })] });
  const newRow = 3;
  const stillIntact = FORMULA_COLS.every(c => log._cells[newRow-1][c].f !== '');
  report('append onto a formula-filled sheet keeps formulas', stillIntact && res2.written === 1);
}

// ---------------------------------------------------------------- TEST 4
{
  // A sheet with NO formulas at all: the app must fall back to writing computed values,
  // so it still works if someone strips the formulas out.
  const sheets = freshSheets();
  const log = makeSheet('บันทึกการซ่อม', LOG_HEADER, [], [], 100);
  sheets['บันทึกการซ่อม'] = log;
  const env = buildEnv(sheets);
  const api = loadCode(env);

  api.saveRecords({ records: [rec()] });
  const row = 2;
  const machineIdWritten = log._cells[row-1][7].v === 'R-B101-S1';
  const zoneWritten = log._cells[row-1][9].v === 'R';
  const checkWritten = log._cells[row-1][12].v === '✓';
  report('no-formula sheet: app fills computed columns itself', machineIdWritten && zoneWritten && checkWritten,
    `MachineID=${log._cells[row-1][7].v} โซน=${log._cells[row-1][9].v} ✓หัววงจร=${log._cells[row-1][12].v}`);
}

// ---------------------------------------------------------------- TEST 5
{
  const sheets = freshSheets();
  const env = buildEnv(sheets);
  const api = loadCode(env);
  const d = api.diagnose();
  report('diagnose reports formula coverage', d.formulaColumns === FORMULA_COLS.length,
    `formulaColumns=${d.formulaColumns} reachRow=${d.formulasReachRow} dataLastRow=${d.dataLastRow} maxRows=${d.maxRows}`);
}
