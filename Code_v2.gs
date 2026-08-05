/**
 * ระบบบริหารโครงการ CULAC — Backend (Google Apps Script)  v2
 * -----------------------------------------------------------
 * ปรับปรุงจากรุ่นเดิมเพื่อแก้ 3 ปัญหาหลัก:
 *
 *   1) ช้าและมีเพดานพัง  — รุ่นเดิมเก็บ JSON ทั้งระบบไว้ในเซลล์ A1 เซลล์เดียว
 *      ซึ่ง Google Sheet จำกัด 50,000 ตัวอักษร/เซลล์ พอข้อมูลโตจะบันทึกไม่ครบ
 *      รุ่นนี้ "แบ่งเก็บเป็นชิ้น (chunk)" ลงหลายเซลล์ในคอลัมน์ A จึงเก็บได้มาก
 *      และไม่พังเมื่อข้อมูลโต
 *
 *   2) โหลดช้าทุกครั้ง  — เพิ่มเลขรุ่นข้อมูล (rev) ให้ฝั่งหน้าเว็บถามสั้น ๆ ว่า
 *      "ข้อมูลเปลี่ยนไหม" (?action=rev) ถ้าไม่เปลี่ยนก็ไม่ต้องดึงทั้งก้อน
 *
 *   3) หลายคนใช้พร้อมกันแล้วข้อมูลหาย  — รุ่นเดิมเป็น "ใครบันทึกทีหลังทับหมด"
 *      รุ่นนี้ให้ฝั่งหน้าเว็บส่ง baseRev มาด้วย ถ้า rev บนเซิร์ฟเวอร์ใหม่กว่า
 *      แปลว่ามีคนอื่นบันทึกแทรก เซิร์ฟเวอร์จะไม่ยอมทับ แต่ส่งข้อมูลล่าสุดกลับไป
 *      ให้หน้าเว็บรวม (merge) แล้วค่อยบันทึกใหม่ — งานจึงไม่หาย
 *      (ถ้าฝั่งหน้าเว็บไม่ส่ง baseRev มา ระบบยังทำงานได้แบบเดิม — เข้ากันได้ย้อนหลัง)
 *
 * วิธีติดตั้ง (เหมือนเดิม)
 * 1) เปิด Google Sheet ที่ใช้เป็นฐานข้อมูล → เมนู "ส่วนขยาย" → "Apps Script"
 * 2) วางโค้ดนี้ทับ Code.gs เดิมทั้งหมด แล้วบันทึก
 * 3) กด "ปรับใช้" → "การปรับใช้ใหม่" (New deployment) → ประเภท "เว็บแอป"
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    (ถ้าเคย Deploy แล้ว ให้ "จัดการการปรับใช้" → แก้ไข → เวอร์ชันใหม่ เพื่อให้ URL เดิมใช้โค้ดใหม่)
 * 4) URL เดิม (/exec) ใช้ต่อได้เลย ไม่ต้องเปลี่ยนในหน้าเว็บ
 *
 * หมายเหตุ: ครั้งแรกที่รันโค้ดใหม่บนชีตที่มีข้อมูลเดิมอยู่ใน A1 ระบบจะอ่านข้อมูลเดิม
 * ได้ตามปกติ และจะเปลี่ยนมาเก็บแบบแบ่งชิ้นให้เองในการบันทึกครั้งถัดไป
 */

const SHEET_NAME = 'DB';
const CHUNK_SIZE = 40000;   // ตัวอักษรต่อเซลล์ (เผื่อจากเพดาน 50,000)
const META_SAVED = 'B1';    // เวลาบันทึกล่าสุด (ISO string)
const META_REV   = 'C1';    // เลขรุ่นข้อมูล (นับขึ้นทุกครั้งที่บันทึกสำเร็จ)
const META_COUNT = 'D1';    // จำนวนชิ้น (chunk) ที่ใช้เก็บ JSON

/** เตรียมแผ่นงาน DB (สร้างให้อัตโนมัติถ้ายังไม่มี) */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A1').setValue('{}');
    sh.getRange(META_SAVED).setValue(new Date().toISOString());
    sh.getRange(META_REV).setValue(0);
    sh.getRange(META_COUNT).setValue(1);
    sh.hideSheet();
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRev_(sh) {
  const v = sh.getRange(META_REV).getValue();
  return (typeof v === 'number' && !isNaN(v)) ? v : 0;
}

function getSavedAt_(sh) {
  const v = sh.getRange(META_SAVED).getValue();
  if (!v) return null;
  try { return new Date(v).toISOString(); } catch (e) { return null; }
}

/**
 * อ่าน JSON ที่แบ่งเก็บเป็นชิ้น (chunk) กลับมาต่อกัน
 * - รองรับรูปแบบเดิม: ถ้า D1 ว่างหรือ =1 ก็อ่านจาก A1 เซลล์เดียว (ข้อมูลเก่าใช้ได้ทันที)
 */
function readData_(sh) {
  const count = Number(sh.getRange(META_COUNT).getValue()) || 1;
  let raw;
  if (count <= 1) {
    raw = sh.getRange('A1').getValue();
  } else {
    const values = sh.getRange(1, 1, count, 1).getValues(); // A1:A{count}
    raw = values.map(r => r[0] == null ? '' : String(r[0])).join('');
  }
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

/** เขียน JSON โดยแบ่งเป็นชิ้นลงคอลัมน์ A (A1, A2, A3, ...) */
function writeData_(sh, dataObj) {
  const json = JSON.stringify(dataObj);
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push([json.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push(['{}']);

  const prevCount = Number(sh.getRange(META_COUNT).getValue()) || 1;

  // ล้างชิ้นเก่าที่เกินความจำเป็น (กันเศษข้อมูลเดิมค้าง)
  if (prevCount > chunks.length) {
    sh.getRange(chunks.length + 1, 1, prevCount - chunks.length, 1).clearContent();
  }
  sh.getRange(1, 1, chunks.length, 1).setValues(chunks); // เขียน A1:A{n}
  sh.getRange(META_COUNT).setValue(chunks.length);
}

/**
 * GET
 *   ?action=load  → ส่งข้อมูลทั้งก้อน + rev + savedAt
 *   ?action=rev   → ส่งเฉพาะ rev + savedAt (เบามาก ใช้เช็คว่าข้อมูลเปลี่ยนหรือยัง)
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const sh = getSheet_();

    if (params.action === 'rev') {
      return jsonOut_({ ok: true, rev: getRev_(sh), savedAt: getSavedAt_(sh) });
    }
    if (params.action === 'load') {
      return jsonOut_({ ok: true, data: readData_(sh), rev: getRev_(sh), savedAt: getSavedAt_(sh) });
    }
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/**
 * POST — บันทึกข้อมูลทั้งก้อน
 *   body: { data: {...}, baseRev?: number }
 *
 *   - ถ้าไม่ส่ง baseRev มา → บันทึกทันที (พฤติกรรมเดิม) แล้วคืน rev ใหม่
 *   - ถ้าส่ง baseRev มา และไม่ตรงกับ rev ปัจจุบัน → มีคนอื่นบันทึกแทรก
 *     คืน { ok:false, conflict:true, rev, data } ให้หน้าเว็บนำไปรวม (merge) แล้วส่งใหม่
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.data === undefined) {
      return jsonOut_({ ok: false, error: 'missing data' });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = getSheet_();
      const currentRev = getRev_(sh);

      // ตรวจการชนกันของหลายผู้ใช้ (เฉพาะเมื่อหน้าเว็บส่ง baseRev มา)
      if (body.baseRev !== undefined && body.baseRev !== null && Number(body.baseRev) !== currentRev) {
        return jsonOut_({
          ok: false,
          conflict: true,
          rev: currentRev,
          savedAt: getSavedAt_(sh),
          data: readData_(sh)
        });
      }

      writeData_(sh, body.data);
      const newRev = currentRev + 1;
      sh.getRange(META_REV).setValue(newRev);
      sh.getRange(META_SAVED).setValue(new Date().toISOString());
      return jsonOut_({ ok: true, rev: newRev, savedAt: getSavedAt_(sh) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
