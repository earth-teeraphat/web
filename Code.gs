// ================================================================
// ระบบรับสมัครผู้เข้าแข่งขันโปรแกรมคอมพิวเตอร์ โรงเรียนหล่มสักวิทยาคม
// ไฟล์ Backend: Code.gs (สำหรับวางใน Google Apps Script Editor)
// ================================================================

const SPREADSHEET_ID = "1M0-mUV462rwAvCOKxaXFRUGXuKNQR3Bj0FzKXddDqTs";

// กำหนดช่วงเวลาเปิดรับสมัคร (4 - 17 สิงหาคม พ.ศ. 2569)
const REG_START_DATE = new Date('2026-08-04T00:00:00+07:00');
const REG_END_DATE = new Date('2026-08-17T23:59:59+07:00');

function isRegistrationOpen() {
  const now = new Date();
  return now >= REG_START_DATE && now <= REG_END_DATE;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('กิจกรรมสัปดาห์วิทยาศาสตร์ (การแข่งขันด้านเทคโนโลยี) โรงเรียนหล่มสักวิทยาคม')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initSheet() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  let sheet = ss.getSheetByName("Registrations");
  if (!sheet) {
    sheet = ss.insertSheet("Registrations");
    const headers = [
      "รหัสอ้างอิง", "วันที่สมัคร", "รายการแข่งขัน", "รหัสหมวดย่อย",
      "ผู้สมัคร 1", "ชั้น ม.1", "เบอร์ ม.1",
      "ผู้สมัคร 2", "ชั้น ม.2", "เบอร์ ม.2", "สถานะ"
    ];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#4f46e5")
         .setFontColor("#ffffff");
  }
  return sheet;
}

function getQuotaStatus() {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  const counts = {};
  
  for (let i = 1; i < data.length; i++) {
    const subCatKey = data[i][3];
    if (subCatKey) {
      counts[subCatKey] = (counts[subCatKey] || 0) + 1;
    }
  }
  return counts;
}

function saveRegistration(payload) {
  try {
    // 1. ตรวจสอบช่วงเวลาเปิดรับสมัคร
    if (!isRegistrationOpen()) {
      return { 
        success: false, 
        message: "ไม่อยู่ในช่วงเวลาเปิดรับสมัคร (เปิดรับสมัครวันที่ 4 - 17 สิงหาคม 2569)" 
      };
    }

    // 2. ตรวจสอบโควตาคงเหลือ
    const sheet = initSheet();
    const currentCounts = getQuotaStatus();
    const currentRegistered = currentCounts[payload.subCategoryId] || 0;
    
    if (currentRegistered >= payload.maxLimit) {
      return { 
        success: false, 
        message: "ขออภัย รายการแข่งขันนี้มีผู้สมัครเต็มจำนวนโควตาแล้ว!" 
      };
    }

    // 3. สร้างรหัสอ้างอิงและบันทึกข้อมูล
    const regId = "LSK-COMP-" + Math.floor(100000 + Math.random() * 900000);
    const timestamp = new Date();

    const rowData = [
      regId,
      timestamp,
      payload.subCategoryName,
      payload.subCategoryId,
      payload.m1Prefix + payload.m1Name,
      payload.m1Class,
      payload.m1Phone,
      payload.m2Prefix + payload.m2Name,
      payload.m2Class,
      payload.m2Phone,
      "อนุมัติแล้ว"
    ];

    sheet.appendRow(rowData);

    return {
      success: true,
      regId: regId,
      message: "ลงทะเบียนเรียบร้อยแล้ว!"
    };

  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() };
  }
}

function getAllRegistrations() {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // ข้ามแถวว่าง

    let formattedDate = "";
    try {
      const d = new Date(row[1]);
      if (!isNaN(d.getTime())) {
        formattedDate = Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm");
      } else {
        formattedDate = String(row[1] || "");
      }
    } catch(e) {
      formattedDate = String(row[1] || "");
    }

    results.push({
      regId: String(row[0]),
      timestamp: formattedDate,
      subCategoryName: String(row[2] || ""),
      subCategoryId: String(row[3] || ""),
      m1Name: String(row[4] || ""),
      m1Class: String(row[5] || ""),
      m1Phone: String(row[6] || ""),
      m2Name: String(row[7] || ""),
      m2Class: String(row[8] || ""),
      m2Phone: String(row[9] || ""),
      status: String(row[10] || "อนุมัติแล้ว")
    });
  }

  // เรียงลำดับรายการล่าสุดขึ้นก่อน
  return results.reverse();
}
