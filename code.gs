// --- Google Apps Script (Backend) ---

// --- ค่าคงที่ (กรุณาตรวจสอบให้แน่ใจว่า ID ถูกต้อง) ---
const SPREADSHEET_ID = '18gvbLCW2TZ90zC7r4WdZxhu-BVxKxBCDleEsXUq1bMI';
const SHEET_NAME = 'Form';
const FOLDER_ID = '1tXofTdn3qaT2o6oGfOUd9qmmlmE-OVZU';

/**
 * ฟังก์ชันหลักในการแสดง Web App
 * @param {object} e - พารามิเตอร์จาก URL
 */
function doGet(e) {
  // ตรวจสอบพารามิเตอร์ 'page' ใน URL เพื่อเลือกว่าจะแสดงหน้าไหน
   if (e && e.parameter && e.parameter.page === 'stats') {
    var template = HtmlService.createTemplateFromFile('stats'); // แสดงหน้าสถิติ
    return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle("สถิติ OTP Virtual Run 2025")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    // ถ้าไม่มี parameter 'page=stats' ให้แสดงหน้าฟอร์มปกติ
    var template = HtmlService.createTemplateFromFile('index'); 
    return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle("OTP Virtual Run 2025")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
/**
 * ฟังก์ชันที่ถูกเรียกจากฝั่ง Client-side เพื่อบันทึกข้อมูล
 * @param {object} data - ข้อมูลทั้งหมดจากฟอร์ม รวมถึงข้อมูลไฟล์
 * @returns {string} - ข้อความสถานะเพื่อส่งกลับไปแสดงผล
 */
function saveData(data) {
  try {
    // 1. อัปโหลดรูปภาพไปยัง Google Drive
    const imageBlob = Utilities.newBlob(
      Utilities.base64Decode(data.fileData),
      data.mimeType,
      data.fileName
    );
    
    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const uploadedFile = targetFolder.createFile(imageBlob);
    const imageUrl = uploadedFile.getUrl();

    // 2. บันทึกข้อมูลลงใน Google Sheet
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const timestamp = new Date();

    // เพิ่มแถวใหม่ (ตรวจสอบลำดับคอลัมน์ให้ตรงกับ Sheet ของคุณ)
    sheet.appendRow([
      timestamp,
      data.fullName,
      data.department,
      data.activityDate,
      data.activityTime,
      data.quantity, // ระยะทาง
      data.duration,
      imageUrl
    ]);
    
    Logger.log('บันทึกข้อมูลสำเร็จ: ' + data.fullName);
    return '✅ บันทึกข้อมูลของคุณเรียบร้อยแล้ว!';

  } catch (error) {
    Logger.log('เกิดข้อผิดพลาด: ' + error.toString());
    throw new Error('ไม่สามารถบันทึกข้อมูลได้ กรุณาลองใหม่อีกครั้ง (' + error.message + ')');
  }
}

/**
 * ฟังก์ชันสำหรับดึงและประมวลผลข้อมูลสถิติจาก Google Sheet
 * @returns {Array} - ข้อมูลหน่วยงานและระยะทางรวม เรียงจากมากไปน้อย
 */
function getStatsData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // ลบ header ออก (แถวแรก)
    values.shift(); 
    
    const departmentStats = {};

    // วนลูปเพื่อรวมระยะทาง (คอลัมน์ F หรือ index 5) ของแต่ละหน่วยงาน (คอลัมน์ C หรือ index 2)
    values.forEach(row => {
      const department = row[2];
      const distance = parseFloat(row[5]);
      
      if (department && !isNaN(distance)) {
        if (departmentStats[department]) {
          departmentStats[department] += distance;
        } else {
          departmentStats[department] = distance;
        }
      }
    });

    // แปลง object เป็น array และเรียงลำดับข้อมูล
    const sortedStats = Object.entries(departmentStats).map(([department, totalDistance]) => {
      return { department: department, distance: totalDistance };
    });

    sortedStats.sort((a, b) => b.distance - a.distance);
    
    Logger.log(sortedStats);
    return sortedStats;

  } catch (error) {
    Logger.log('เกิดข้อผิดพลาดในการดึงข้อมูลสถิติ: ' + error.toString());
    throw new Error('ไม่สามารถดึงข้อมูลสถิติได้: ' + error.message);
  }
}
