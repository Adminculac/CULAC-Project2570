/**
 * ระบบบริหารโครงการ CULAC — Backend (Google Apps Script)
 * ------------------------------------------------------
 * สคริปต์นี้ทำหน้าที่เป็น "ฐานข้อมูลกลาง" ให้กับไฟล์ HTML ของระบบ
 * โดยเก็บข้อมูลทั้งหมด (โครงการ, บุคลากร, สิทธิ์, ประวัติ ฯลฯ) เป็น JSON
 * ก้อนเดียวไว้ใน Google Sheet (แผ่นงานชื่อ "DB") และให้บริการ 2 คำสั่ง:
 *
 *   - โหลดข้อมูล : GET  ?action=load
 *   - บันทึกข้อมูล: POST body = {"data":{...}}
 *
 * รุ่นนี้ไม่ใช้ Token — ความปลอดภัยอาศัย "ความไม่เปิดเผยของ URL" เพียงอย่างเดียว
 * (URL ที่ได้จากการ Deploy มีรหัสสุ่มยาวคาดเดายาก) เหมาะกับการใช้งานภายในหน่วยงาน
 * ที่แจกลิงก์ให้เฉพาะบุคลากรที่เกี่ยวข้องเท่านั้น
 *
 * วิธีติดตั้ง
 * 1) เปิด Google Sheet ใหม่ 1 ไฟล์ (ใช้เป็นฐานข้อมูล)
 * 2) เมนู "ส่วนขยาย" → "Apps Script" แล้ววางโค้ดนี้ทับไฟล์ Code.gs เดิมทั้งหมด
 * 3) กด "ปรับใช้" (Deploy) → "ปรับใช้เป็นเว็บแอป" (New deployment → Web app)
 *      - Execute as: Me
 *      - Who has access: Anyone (จำเป็น เพราะไฟล์ HTML เรียกจากเบราว์เซอร์ของผู้ใช้โดยตรง
 *        โดยไม่ผ่านการล็อกอิน Google ใดๆ)
 * 4) คัดลอก URL ของเว็บแอปที่ได้ (ลงท้ายด้วย /exec)
 * 5) เปิดไฟล์ระบบ (culac.html) → เข้าสู่ระบบด้วยบัญชีเลขานุการ → เมนู "ตั้งค่าระบบ"
 *    → การ์ด "การเชื่อมต่อฐานข้อมูลกลาง" → วาง URL จากข้อ 4 ลงในช่อง แล้วกด "บันทึกและทดสอบ"
 *    ระบบจะเปลี่ยนจาก "โหมดทดลอง" (เก็บในเครื่องผู้ใช้เท่านั้น) เป็นโหมดซิงก์ข้อมูลกลางทันที
 * 6) หลังเชื่อมต่อสำเร็จ ระบบจะสร้าง "ลิงก์สำหรับแจกให้ทุกคน" ให้อัตโนมัติ (มี ?gas=... ต่อท้าย)
 *    คัดลอกลิงก์นั้นส่งให้ทุกคนใช้เปิดแทนไฟล์เดิม — เปิดครั้งแรกครั้งเดียว ระบบจะจำการเชื่อมต่อ
 *    ไว้ให้อัตโนมัติในเครื่องนั้น ไม่ต้องตั้งค่าเองอีก
 *
 * หมายเหตุความปลอดภัย: อย่าเผยแพร่ URL เว็บแอปนี้แบบสาธารณะ (เช่น โพสต์ในเว็บ/โซเชียลที่เปิดเผย)
 * เพราะใครก็ตามที่มี URL นี้จะสามารถอ่าน/เขียนทับข้อมูลทั้งระบบได้ทันที โดยไม่ต้องมีบัญชีหรือรหัสผ่านใดๆ
 * แจกให้เฉพาะบุคลากรที่เกี่ยวข้องเท่านั้น (เช่น ส่งทางไลน์กลุ่มหน่วยงาน ไม่ใช่โพสต์สาธารณะ)
 */

const SHEET_NAME = 'DB';
const CELL = 'A1';   // เซลล์ที่ใช้เก็บ JSON ก้อนเดียวของทั้งระบบ

/** เตรียมแผ่นงาน DB ให้พร้อมใช้งาน (สร้างให้อัตโนมัติถ้ายังไม่มี) */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A1').setValue('{}');
    sh.getRange('B1').setValue(new Date().toISOString());
    sh.hideSheet();
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET: ใช้สำหรับโหลดข้อมูลล่าสุด (?action=load) */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.action !== 'load') {
      return jsonOut_({ ok: false, error: 'unknown action' });
    }
    const sh = getSheet_();
    const raw = sh.getRange(CELL).getValue();
    const savedAt = sh.getRange('B1').getValue();
    let data = null;
    if (raw) {
      try { data = JSON.parse(raw); } catch (err) { data = null; }
    }
    return jsonOut_({ ok: true, data: data, savedAt: savedAt ? new Date(savedAt).toISOString() : null });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** POST: ใช้สำหรับบันทึกข้อมูลทั้งก้อน (body: {data}) */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.data === undefined) {
      return jsonOut_({ ok: false, error: 'missing data' });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // กันข้อมูลชนกันเมื่อมีหลายคนบันทึกพร้อมกัน
    try {
      const sh = getSheet_();
      sh.getRange(CELL).setValue(JSON.stringify(body.data));
      sh.getRange('B1').setValue(new Date().toISOString());
    } finally {
      lock.releaseLock();
    }
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
