/**
 * ระบบสแกนใบหน้าเข้า-ออกงาน
 * ------------------------------------------------------
 * วิธีตั้งค่า:
 * 1. เปลี่ยนค่า DRIVE_FOLDER_ID ด้านล่างเป็น Folder ID ของโฟลเดอร์รูปพนักงานใน Drive
 * 2. ชื่อไฟล์รูปพนักงานควรตั้งเป็น "รหัสพนักงาน_ชื่อพนักงาน.jpg" เช่น "EMP001_สมชาย ใจดี.jpg"
 *    (ถ้าไม่มี "_" ระบบจะใช้ชื่อไฟล์ทั้งหมดเป็นทั้งรหัสและชื่อ)
 * ------------------------------------------------------
 */

const DRIVE_FOLDER_ID = 'ใส่_FOLDER_ID_ของคุณตรงนี้'; // <-- แก้ตรงนี้
const SHEET_NAME = 'Attendance';

/**
 * เสิร์ฟหน้าเว็บสแกนหน้า (เปิดผ่านลิงก์ Web App)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบสแกนใบหน้าเข้า-ออกงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * คืนค่าข้อมูลเบา ๆ (จำนวนไฟล์ + เวลาแก้ไขล่าสุด) ใช้เช็คว่าต้องโหลดรูปใหม่ทั้งหมดหรือไม่
 * เพื่อให้หน้าเว็บ cache ข้อมูลใบหน้าไว้ใน localStorage แล้วเปิดครั้งถัดไปเร็วขึ้น
 */
function getEmployeeMeta() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.JPEG);
  const files2 = folder.getFilesByType(MimeType.PNG);
  let count = 0;
  let lastUpdatedSum = 0;
  const collect = (iterator) => {
    while (iterator.hasNext()) {
      const f = iterator.next();
      count++;
      lastUpdatedSum += f.getLastUpdated().getTime();
    }
  };
  collect(files);
  collect(files2);
  return { cacheKey: count + '-' + lastUpdatedSum };
}

/**
 * ดึงรูปพนักงานทั้งหมดจาก Drive แปลงเป็น base64 ส่งให้ frontend ไปตรวจจับใบหน้าเอง
 */
function getEmployeePhotos() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const employees = [];

  const addFromIterator = (iterator) => {
    while (iterator.hasNext()) {
      const file = iterator.next();
      const rawName = file.getName().replace(/\.[^/.]+$/, ''); // ตัดนามสกุลไฟล์ออก
      const parts = rawName.split('_');
      const empId = parts[0].trim();
      const empName = parts.length > 1 ? parts.slice(1).join('_').trim() : empId;

      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      const mimeType = blob.getContentType();

      employees.push({
        id: empId,
        name: empName,
        dataUrl: 'data:' + mimeType + ';base64,' + base64
      });
    }
  };

  addFromIterator(folder.getFilesByType(MimeType.JPEG));
  addFromIterator(folder.getFilesByType(MimeType.PNG));

  return employees;
}

/**
 * บันทึกเวลาเข้า/ออกงานลง Google Sheet
 * จะสลับ IN/OUT ให้อัตโนมัติตามประวัติล่าสุดของพนักงานคนนั้นในวันเดียวกัน
 */
function logAttendance(empId, empName, confidence) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['เวลา (Timestamp)', 'รหัสพนักงาน', 'ชื่อ', 'ประเภท', 'ความมั่นใจ (Confidence)']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  // หาประเภทล่าสุดของพนักงานคนนี้ในวันนี้ เพื่อสลับ IN/OUT
  const data = sheet.getDataRange().getValues();
  let lastType = null;
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = Utilities.formatDate(new Date(data[i][0]), tz, 'yyyy-MM-dd');
    if (String(data[i][1]) === String(empId) && rowDate === today) {
      lastType = data[i][3];
      break;
    }
  }

  const type = (lastType === 'IN') ? 'OUT' : 'IN';
  sheet.appendRow([now, empId, empName, type, confidence]);

  return {
    success: true,
    type: type,
    timestamp: now.toISOString(),
    displayTime: Utilities.formatDate(now, tz, 'HH:mm:ss')
  };
}
