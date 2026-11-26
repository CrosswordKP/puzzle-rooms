/**
 * กระดานภารกิจ ห้องปริศนา (ว21105 แผนที่ 1)
 * รับความคืบหน้าจากมือถือนักเรียน (ชุดรหัส ทีม บทบาท ด่าน เท่านั้น ไม่มีชื่อนักเรียน)
 * และส่งสถานะล่าสุดให้หน้ากระดานบนเครื่องครู
 *
 * ติดตั้ง : สร้าง Google Sheet > ส่วนขยาย > Apps Script > วางโค้ดนี้ > ทำให้ใช้งานได้ > เว็บแอป
 *          ดำเนินการในฐานะ : ฉัน · ผู้ที่มีสิทธิ์เข้าถึง : ทุกคน
 */
var LOG = 'log';
var HEAD = ['เวลา', 'ts', 'ชุดรหัส', 'ทีม', 'บทบาท', 'ด่าน', 'ข้อ', 'งานที่เสร็จ', 'งานทั้งหมด', 'ตอบผิด', 'เหตุการณ์', 'เครื่อง'];
var EVENTS = { start: 1, task: 1, finish: 1, capsule: 1 };

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LOG);
  if (!sh) {
    sh = ss.insertSheet(LOG);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function int_(v, lo, hi) {
  var n = Math.floor(Number(v));
  if (!isFinite(n) || n < lo || n > hi) throw new Error('out of range');
  return n;
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ev = String(d.ev || '');
    if (!EVENTS[ev]) throw new Error('event');
    var row = [
      new Date(), Date.now(),
      int_(d.set, 1, 10), int_(d.team, 1, 5), int_(d.role, 0, 5),
      int_(d.stage, 0, 4), int_(d.task, 0, 20), int_(d.done, 0, 50), int_(d.total, 0, 50), int_(d.att, 0, 99),
      ev, String(d.dev || '').replace(/[^a-z0-9]/gi, '').slice(0, 12)
    ];
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { sheet_().appendRow(row); } finally { lock.releaseLock(); }
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('ignored');
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var set = Number(p.set) || 1, since = Number(p.since) || 0;
  var values = sheet_().getDataRange().getValues();
  var latest = {}, caps = {};
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (Number(r[2]) !== set || Number(r[1]) < since) continue;
    var team = Number(r[3]), role = Number(r[4]), ts = Number(r[1]), ev = String(r[10]);
    if (ev === 'capsule') {
      if (!caps[team] || ts < caps[team]) caps[team] = ts;
      continue;
    }
    var key = team + '-' + role;
    if (!latest[key] || ts >= latest[key].ts) {
      latest[key] = { team: team, role: role, stage: Number(r[5]), task: Number(r[6]), done: Number(r[7]), total: Number(r[8]), ev: ev, ts: ts };
    }
  }
  var out = {
    ok: true,
    now: Date.now(),
    rows: Object.keys(latest).map(function (k) { return latest[k]; }),
    caps: Object.keys(caps).map(function (t) { return { team: Number(t), ts: caps[t] }; })
  };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** กดรันครั้งแรกเพื่อสร้างแผ่นงาน log และให้สิทธิ์ */
function setup() {
  sheet_();
}
