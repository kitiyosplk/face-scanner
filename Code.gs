const SPREADSHEET_ID = '1VkTCXAkbc0yc_my7ZeZG9utUdLpfCDPc4g1P9RvVcT4';
const REG_FOLDER_ID = '1M-8AIn5zzKqTreTghYgogg-2S8Wfb5tq'; // โฟลเดอร์ ลงทะเบียน
const SCAN_FOLDER_ID = '1dyzkvfwflEILGJ4shzXn_bypCk2jOynO'; // โฟลเดอร์ สแกนเข้างาน

/**
 * รองรับการดึงข้อมูลพนักงานผ่าน GET Request จาก Vercel
 */
function doGet(e) {
  try {
    const action = e.parameter.action || 'getEmployees';
    
    if (action === 'getEmployees') {
      const data = getEmployeePhotos();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({error: 'Invalid action'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * รองรับการบันทึกเวลาและอัปโหลดรูปภาพผ่าน POST Request จาก Vercel
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const empId = data.empId;
    const empName = data.name;
    const department = data.department;
    const actionType = data.actionType;
    const confidence = data.confidence;
    const snapshotBase64 = data.snapshot;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getActiveSheet();
    
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    
    let otHours = "";
    let otRate = "";
    let otTotal = "";
    let remarks = "สแกนหน้าสำเร็จ";

    // บันทึกภาพถ่ายจริงลงในโฟลเดอร์ "สแกนเข้างาน"
    if (snapshotBase64) {
      try {
        const scanFolder = DriveApp.getFolderById(SCAN_FOLDER_ID);
        const timeStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HHmmss');
        const fileName = empId + "_" + empName + "_" + actionType + "_" + timeStr + ".jpg";
        
        const base64Data = snapshotBase64.split(",")[1];
        const decodedBytes = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
        scanFolder.createFile(blob);
        remarks = "บันทึกรูปหลักฐานสำเร็จ";
      } catch (imgErr) {
        remarks = "บันทึกรูปไม่สำเร็จ: " + imgErr.toString();
      }
    }

    // ระบบคำนวณ OT อัตโนมัติ (70 บาท/ชั่วโมง)
    if (actionType === "ออก OT") {
      otRate = 70;
      const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
      const dataRange = sheet.getDataRange().getValues();
      let startTime = null;

      for (let i = dataRange.length - 1; i >= 1; i--) {
        const rowDate = new Date(dataRange[i][0]);
        const rowDateStr = Utilities.formatDate(rowDate, tz, "yyyy-MM-dd");
        const rowEmpId = String(dataRange[i][1]);
        const rowAction = dataRange[i][4];

        if (rowEmpId === String(empId) && rowDateStr === todayStr && rowAction === "เข้าทำ OT") {
          startTime = rowDate;
          break;
        }
      }

      if (startTime) {
        const diffMs = now - startTime;
        otHours = Number((diffMs / (1000 * 60 * 60)).toFixed(2));
        otTotal = otHours * otRate;
        remarks += " | คำนวณ OT อัตโนมัติสำเร็จ";
      } else {
        otHours = 0;
        otTotal = 0;
        remarks += " | ไม่พบรายการ 'เข้าทำ OT' ในวันนี้";
      }
    } else if (actionType === "เข้าทำ OT") {
      otRate = 70;
    }

    // บันทึกลง Google Sheets ตามคอลัมน์ A ถึง K
    sheet.appendRow([
      now,
      empId,
      empName,
      department,
      actionType,
      otHours,
      otRate,
      otTotal,
      (parseFloat(confidence) * 100).toFixed(2) + "%",
      "ปกติ",
      remarks
    ]);

    const displayTime = Utilities.formatDate(now, tz, 'HH:mm:ss');
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      actionType: actionType,
      displayTime: displayTime,
      otTotal: otTotal
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันดึงรูปพนักงานจากโฟลเดอร์ "ลงทะเบียน"
 */
function getEmployeePhotos() {
  const folder = DriveApp.getFolderById(REG_FOLDER_ID);
  const employees = [];
  
  const addFromIterator = (iterator) => {
    while (iterator.hasNext()) {
      const file = iterator.next();
      const rawName = file.getName().replace(/\.[^/.]+$/, '');
      const parts = rawName.split('_');
      const empId = parts[0].trim();
      const empName = parts.length > 1 ? parts[1].trim() : empId;
      const department = parts.length > 2 ? parts[2].trim() : 'ทั่วไป';

      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      const mimeType = blob.getContentType();

      employees.push({
        id: empId,
        name: empName,
        department: department,
        dataUrl: 'data:' + mimeType + ';base64,' + base64
      });
    }
  };

  addFromIterator(folder.getFilesByType(MimeType.JPEG));
  addFromIterator(folder.getFilesByType(MimeType.PNG));
  return employees;
}