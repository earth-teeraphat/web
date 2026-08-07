
/**
 * ==============================================================================
 * ระบบรายงานการทำเวรประจำวัน ม.4/3 โรงเรียนหล่มสักวิทยาคม (Google Apps Script Backend)
 * ==============================================================================
 * คำแนะนำการใช้งาน:
 * 1. นำโค้ดนี้ไปวางในไฟล์ Code.gs บน Google Apps Script
 * 2. ระบุ FOLDER_ID ของ Google Drive สำหรับเก็บรูปภาพ
 * 3. ทำการ Deploy เป็น Web App (Execute as: Me, Who has access: Anyone)
 */

// Global Configuration Variables
const CONFIG = {
  SHEET_ID: "17dS7KQME0t0KLdw1RUKsKwfFm3sov9lz0vFyo4rswws",       // ชื่อชีตสำหรับบันทึกข้อมูล
  FOLDER_ID: "1o9i_Vc1EnD_i3YAUq6H30DEKznJ7vYeA",                     // หากต้องการกำหนด Google Drive Folder ID เจาะจง ให้ใส่ตรงนี้ (ถ้าว่างไว้ระบบจะสร้างโฟลเดอร์ให้อัตโนมัติ)
  FOLDER_NAME: "" // ชื่อโฟลเดอร์สำรองหากไม่ได้ระบุ ID
};

/**
 * ฟังก์ชันหลักในการแสดงผลหน้าเว็บ HTML (Web App Entry Point)
 */
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ระบบรายงานการทำเวรประจำวัน ม.4/3 โรงเรียนหล่มสักวิทยาคม')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("เกิดข้อผิดพลาดในการโหลดระบบ: " + err.toString());
  }
}

/**
 * รองรับการส่งข้อมูลผ่าน Web API POST (สำหรับกรณีการเชื่อมต่อจากภายนอก)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = saveDutyReport(data);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * บันทึกรายงานการทำเวรลง Google Sheet และอัปโหลดรูปภาพลง Google Drive
 * @param {Object} data ข้อมูลรายงานที่ส่งมาจาก Frontend
 * @returns {Object} ผลลัพธ์สถานะการทำงาน
 */
function saveDutyReport(data) {
  try {
    const sheet = getOrCreateSheet();
    let imageUrl = "";

    // 1. จัดการอัปโหลดรูปภาพไปยัง Google Drive
    if (data.imageData && data.imageData.includes("base64,")) {
      imageUrl = uploadImageToDrive(data.imageData, data.date, data.dayOfWeek, data.shift);
    }

    // 2. แปลงรายการสถานะนักเรียนเป็นข้อความ JSON
    const studentStatusJson = JSON.stringify(data.studentStatuses || []);
    const areasString = Array.isArray(data.areas) ? data.areas.join(", ") : (data.areas || "");

    // 3. เพิ่มข้อมูลลงในแถวใหม่ของ Google Sheet
    const recordId = "REP-" + Date.now();
    sheet.appendRow([
      recordId,
      new Date(),                  // วันเวลาที่บันทึก
      data.date,                   // วันที่ปฏิบัติหน้าที่
      data.dayOfWeek,              // วันประจำสัปดาห์
      data.shift,                  // ช่วงเวลา (เช้า/เย็น)
      data.completedCount || 0,    // จำนวนคนที่ทำเรียบร้อย
      data.totalCount || 0,        // จำนวนนักเรียนเวรทั้งหมด
      studentStatusJson,          // รายละเอียดการเช็คชื่อรายบุคคล (JSON)
      areasString,                // บริเวณที่ทำความสะอาด
      imageUrl,                   // ลิงก์รูปภาพใน Google Drive
      data.remarks || ""          // หมายเหตุ
    ]);

    return {
      status: "success",
      message: "บันทึกรายงานการทำเวรสำเร็จเรียบร้อยแล้ว",
      id: recordId,
      imageUrl: imageUrl
    };

  } catch (err) {
    Logger.log("Error in saveDutyReport: " + err.toString());
    return {
      status: "error",
      message: "เกิดข้อผิดพลาดฝั่ง Server: " + err.toString()
    };
  }
}

/**
 * ดึงรายการประวัติการทำเวรทั้งหมดจาก Google Sheet
 * @returns {Array} รายการบันทึกเวรทั้งหมด
 */
function getDutyReports() {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    // หากมีแค่ Header
    if (data.length <= 1) {
      return [];
    }

    const reports = [];
    // ข้ามแถวที่ 0 (Header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // ข้ามแถวว่าง

      let studentStatuses = [];
      try {
        studentStatuses = JSON.parse(row[7]);
      } catch (e) {
        studentStatuses = [];
      }

      reports.push({
        id: row[0],
        timestamp: row[1] ? new Date(row[1]).toLocaleString('th-TH') : '',
        date: row[2] ? formatDateString(row[2]) : '',
        dayOfWeek: row[3] || '',
        shift: row[4] || '',
        completedCount: Number(row[5]) || 0,
        totalCount: Number(row[6]) || 0,
        studentStatuses: studentStatuses,
        areas: row[8] ? row[8].toString().split(', ') : [],
        imageUrl: row[9] || '',
        remarks: row[10] || ''
      });
    }

    // เรียงลำดับจากล่าสุดไปเก่าสุด
    return reports.reverse();

  } catch (err) {
    Logger.log("Error in getDutyReports: " + err.toString());
    return [];
  }
}

/**
 * ฟังก์ชันช่วยแปลงภาพ Base64 และบันทึกลงใน Google Drive
 */
function uploadImageToDrive(base64Data, date, day, shift) {
  try {
    const folder = getDriveFolder();
    
    // ดึงประเภทไฟล์และข้อมูล Base64
    const splitData = base64Data.split("base64,");
    const contentType = splitData[0].split(":")[1].split(";")[0];
    const decodedData = Utilities.base64Decode(splitData[1]);
    
    const fileName = `เวร_${day}_${date}_${shift.includes("เช้า") ? "เช้า" : "เย็น"}_${Date.now()}.jpg`;
    const blob = Utilities.newBlob(decodedData, contentType, fileName);
    
    // สร้างไฟล์ใน Drive
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // คืนค่า Direct CDN Image URL เพื่อให้ <img> แสดงผลได้โดยตรง
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    Logger.log("Image upload error: " + e.toString());
    return "";
  }
}

/**
 * ตรวจสอบและดึงชีตสำหรับเก็บข้อมูล หากยังไม่มีจะสร้างขึ้นใหม่อัตโนมัติพร้อม Header
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // สร้าง Header Column
    const headers = [
      "ID",
      "วันเวลาที่บันทึก",
      "วันที่ทำเวร",
      "วันประจำสัปดาห์",
      "ช่วงเวลา",
      "จำนวนที่ทำเรียบร้อย",
      "จำนวนเวรทั้งหมด",
      "สถานะการเช็คชื่อ (JSON)",
      "บริเวณที่ทำความสะอาด",
      "ลิงก์รูปภาพ Drive",
      "หมายเหตุ"
    ];
    sheet.appendRow(headers);
    
    // ตกแต่ง Header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1e1b4b")
               .setFontColor("#ffffff")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * ดึงโฟลเดอร์ Google Drive หรือสร้างใหม่หากไม่มี
 */
function getDriveFolder() {
  if (CONFIG.FOLDER_ID && CONFIG.FOLDER_ID.trim() !== "") {
    try {
      return DriveApp.getFolderById(CONFIG.FOLDER_ID);
    } catch (e) {
      Logger.log("Folder ID ไม่ถูกต้อง กำลังใช้โฟลเดอร์สำรอง...");
    }
  }

  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(CONFIG.FOLDER_NAME);
  }
}

/**
 * แปลงวันที่เป็น YYYY-MM-DD
 */
function formatDateString(d) {
  if (d instanceof Date) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return d.toString().split('T')[0];
}
