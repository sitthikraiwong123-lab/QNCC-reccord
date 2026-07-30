/* Can new symptoms / fixture types be added by editing the sheets alone? */
const { makeSheet, buildEnv, loadCode } = require('./gas-harness');

const LOG_HEADER = ['ลำดับ','วันที่','รอบ','รหัสจุด','เลขเครื่อง','ประเภท (เลือกเมื่อเลขซ้ำ)','ประเภทที่ใช้',
'Machine ID','ตรวจสอบ','โซน','ชื่อโซน','รหัสอาการ','หัววงจรเสีย','รางแบตเตอรี่ชำรุด','ฝาหน้า/ฝาหลัง','น้ำยารั่ว',
'ตัวเครื่อง','ท่อหลุด','ฝารองนั่งหาย','(สำรอง 8)','(สำรอง 9)','(สำรอง 10)','อาการเสีย','สรุปอาการ (รหัส+ชื่อ)',
'รายละเอียดอาการ/การซ่อม','สถานะ','ผู้บันทึก','หมายเหตุ','เดือน','ไตรมาส','ปี','คีย์ช่วงเวลา','pidx','cS','cU','cF','cD'];
const FORMULA_COLS = [0,2,6,7,8,9,10,12,13,14,15,16,17,18,19,20,21,22,23,28,29,30,31,32,33,34,35,36];

/* getMasterData() isn't in the harness's export list, so evaluate it directly */
function masterData(sheets) {
  const env = buildEnv(sheets);
  const src = require('fs').readFileSync('/home/user/QNCC-reccord/Code.gs', 'utf8');
  const names = Object.keys(env);
  const fn = new Function(...names, src + '\n; return getMasterData();');
  return fn(...names.map(n => env[n]));
}

function report(title, pass, extra) {
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${title}${extra ? '  — ' + extra : ''}`);
  if (!pass) process.exitCode = 1;
}

function sheetsWith(extraSymptomRows, extraRegistryRows) {
  const log = makeSheet('บันทึกการซ่อม', LOG_HEADER, FORMULA_COLS, [], 200);

  const regHeader = ['Machine ID','โซน','ชื่อโซน','ชั้น/โซนย่อย','รหัสจุด','รหัสประเภท','ประเภทสุขภัณฑ์','เลขข้างเครื่อง','รุ่น','ตำแหน่ง/หมายเหตุ','สถานะ'];
  const regRows = [
    ['R-B101-S1','R','Ratchada','F.B1','R-B101','S','ชาย (นั่ง)','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
    ['R-B101-U1','R','Ratchada','F.B1','R-B101','U','ชาย (ยืน)','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
    ['R-B101-F1','R','Ratchada','F.B1','R-B101','F','หญิง','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
    ['R-B101-D1','R','Ratchada','F.B1','R-B101','D','คนพิการ','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน'],
  ].concat(extraRegistryRows || []);
  const registry = makeSheet('ทะเบียนเครื่อง', regHeader, [], regRows, 60);

  const masterHeader = ['รุ่น','รายละเอียด/หมายเหตุ','','รหัส','อาการ'];
  const masterRows = [
    ['รุ่น-1','ปัจจุบันมีรุ่นเดียว','','1','หัววงจรเสีย'],
    ['','','','2','รางแบตเตอรี่ชำรุด'],
    ['','','','3','ฝาหน้า/ฝาหลัง'],
    ['','','','4','น้ำยารั่ว'],
    ['','','','5','ตัวเครื่อง'],
    ['','','','6','ท่อหลุด'],
  ].concat(extraSymptomRows || []);
  const master = makeSheet('ข้อมูลหลัก', masterHeader, [], masterRows, 60);
  return { 'บันทึกการซ่อม': log, 'ทะเบียนเครื่อง': registry, 'ข้อมูลหลัก': master };
}

// ---------------------------------------------------------------- symptoms
{
  // add symptom #7 the way a user would: type code + name into ข้อมูลหลัก
  const sheets = sheetsWith([['','','','7','ฝารองนั่งหาย']]);
  const api = loadCode(buildEnv(sheets));
  const md = masterData(sheets);

  const has7 = md.symptoms.length === 7 && md.symptoms[6].code === '7' && md.symptoms[6].name === 'ฝารองนั่งหาย';
  report('symptom added in ข้อมูลหลัก is picked up by the app', has7,
    `symptoms=${md.symptoms.length} last=${md.symptoms[6] ? md.symptoms[6].code + ':' + md.symptoms[6].name : '-'}`);

  // and it can actually be recorded — the code lands in รหัสอาการ, which drives the
  // sheet's pre-wired ✓ column for code 7
  const log = sheets['บันทึกการซ่อม'];
  api.saveRecords({ records: [{
    clientId:'ob_x', date:'2026-07-30', round:'2026-07-R2', recorder:'ช่างเอ',
    pointCode:'R-B101', zone:'R', zoneName:'Ratchada', machineId:'R-B101-S1', num:'1',
    typeName:'ชาย (นั่ง)', found:[{code:'7',name:'ฝารองนั่งหาย'}], repaired:[], pending:[{code:'7',name:'ฝารองนั่งหาย'}],
    status:'รออะไหล่', detail:'', note:''
  }]});
  const codeCell = log._cells[1][11].v;   // รหัสอาการ on the written row
  report('a 7th symptom records its code into รหัสอาการ', String(codeCell) === '7', `รหัสอาการ=${codeCell}`);
}

// ---------------------------------------------------------------- types
{
  // add a 5th fixture type by registering machines of that type
  const sheets = sheetsWith([], [
    ['R-B101-C1','R','Ratchada','F.B1','R-B101','C','ห้องอาบน้ำ','1','รุ่น-1','ฟู้ด คอร์ท','ใช้งาน']
  ]);
  const md = masterData(sheets);
  const t = md.types.find(x => x.code === 'C');
  report('fixture type added in ทะเบียนเครื่อง is picked up by the app', !!t && t.name === 'ห้องอาบน้ำ',
    `types=${md.types.map(x => x.code).join(',')}`);
  const pt = md.points.find(p => p.code === 'R-B101');
  const m = pt && pt.machines.find(x => x.id === 'R-B101-C1');
  report('its machines appear in the point picker', !!m, m ? `found ${m.id} (type ${m.type})` : 'missing');
}
