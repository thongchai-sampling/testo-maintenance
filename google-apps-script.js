/**
 * Google Apps Script สำหรับระบบรายงานการตรวจสอบและบำรุงรักษา Testo
 * รองรับเซ็นเซอร์แยก 5 ชนิด: O2, CO, NO, NO2, SO2 พร้อมสถานะสรุปรวม
 */

// กำหนดชื่อชีตที่ต้องการบันทึกข้อมูล (ถ้าใช้ชีตแรกให้ปล่อยเป็น null หรือระบุชื่อแท็บ เช่น 'ชีต1' หรือ 'Sheet1')
const SHEET_NAME = null; 

/**
 * ฟังก์ชันรับคำขอ POST จากหน้าเว็บ Form (index.html)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // รอล็อคสูงสุด 10 วินาที เพื่อป้องกันปัญหาข้อมูลชนกันเมื่อมีการส่งพร้อมกัน
  if (!lock.tryLock(10000)) {
    return createJsonResponse({ status: 'error', message: 'ระบบกำลังประมวลผลข้อมูลอื่นอยู่ กรุณาลองใหม่อีกครั้ง' });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

    // ถ้ายังไม่มีหัวตาราง ให้สร้างอัตโนมัติ
    if (sheet.getLastRow() === 0) {
      setupHeaders(sheet);
    }

    // แปลงข้อมูล JSON ที่ส่งมาจาก Form
    const data = JSON.parse(e.postData.contents);

    // วันที่และเวลาปัจจุบันที่บันทึกลงระบบ
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

    // จัดเรียงข้อมูลให้ตรงกับหัวตาราง 16 คอลัมน์ (A - P)
    const rowData = [
      timestamp,                               // A: วันเวลาที่บันทึกระบบ (Timestamp)
      data.date || '',                         // B: วันที่ตรวจสอบและบำรุงรักษา (Date)
      data.technician || '',                   // C: ชื่อผู้ตรวจสอบ (Technician)
      data.model || '',                        // D: รุ่นเครื่อง (Model)
      data.serial || '',                       // E: หมายเลขเครื่อง (Serial No.)
      data.chk_filter || '',                   // F: 1. ฟิลเตอร์กรองฝุ่น (Filter)
      data.chk_pump || '',                     // G: 2. ปั๊มดูดอากาศ (Pump)
      data.chk_battery || '',                  // H: 3. ความพร้อมของแบตเตอรี่ (Battery)
      data.chk_clean || '',                    // I: 4. ความสะอาดของเครื่องมือ (Cleanliness)
      data.chk_sensor_o2 || '',                // J: 5. เซ็นเซอร์ O2 (Oxygen)
      data.chk_sensor_co || '',                // K: 6. เซ็นเซอร์ CO (Carbon Monoxide)
      data.chk_sensor_no || '',                // L: 7. เซ็นเซอร์ NO (Nitric Oxide)
      data.chk_sensor_no2 || '',               // M: 8. เซ็นเซอร์ NO2 (Nitrogen Dioxide)
      data.chk_sensor_so2 || '',               // N: 9. เซ็นเซอร์ SO2 (Sulfur Dioxide)
      data.chk_sensor || '',                   // O: สถานะเซ็นเซอร์รวม (Sensor Summary)
      data.notes || ''                         // P: ภาพรวมหรือรายละเอียดเพิ่มเติม (Notes)
    ];

    // เพิ่มแถวข้อมูลใหม่
    sheet.appendRow(rowData);

    return createJsonResponse({ 
      status: 'success', 
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว' 
    });

  } catch (err) {
    return createJsonResponse({ 
      status: 'error', 
      message: err.toString() 
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * ฟังก์ชันสำหรับทดสอบการทำงานของ Web App ผ่าน Browser (GET)
 */
function doGet(e) {
  return ContentService.createTextOutput("Testo Maintenance Web App is running successfully!");
}

/**
 * ฟังก์ชันช่วยสร้าง JSON Response ส่งกลับไปยังหน้าเว็บ
 */
function createJsonResponse(outputObj) {
  return ContentService.createTextOutput(JSON.stringify(outputObj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ฟังก์ชันสำหรับสร้างหัวตาราง (Headers) อัตโนมัติ พร้อมจัดสไตล์ให้สวยงาม
 * สามารถกดปุ่ม "Run" ที่ฟังก์ชันนี้ใน Apps Script Editor ได้ทันที
 */
function setupHeaders(targetSheet) {
  const sheet = targetSheet || (SHEET_NAME ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) : SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]);
  
  const headers = [
    "วันเวลาที่บันทึกระบบ (Timestamp)",
    "วันที่ตรวจสอบ (Date)",
    "ชื่อผู้ตรวจสอบ (Technician)",
    "รุ่นเครื่อง (Model)",
    "หมายเลขเครื่อง (Serial No.)",
    "1. ฟิลเตอร์กรองฝุ่น",
    "2. ปั๊มดูดอากาศ",
    "3. แบตเตอรี่",
    "4. ความสะอาดเครื่องมือ",
    "5. เซ็นเซอร์ O2",
    "6. เซ็นเซอร์ CO",
    "7. เซ็นเซอร์ NO",
    "8. เซ็นเซอร์ NO2",
    "9. เซ็นเซอร์ SO2",
    "สถานะเซ็นเซอร์รวม",
    "รายละเอียดเพิ่มเติม (Notes)"
  ];

  // ใส่หัวตารางที่แถว 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // จัดรูปแบบหัวตาราง
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#2563eb");     // พื้นหลังสีน้ำเงิน (เข้ากับธีมของเว็บ)
  headerRange.setFontColor("#ffffff");       // ตัวอักษรสีขาว
  headerRange.setFontWeight("bold");         // ตัวหนา
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  // ตรึงแถวแรก (Freeze Row)
  sheet.setFrozenRows(1);

  // ปรับความกว้างคอลัมน์อัตโนมัติ
  for (let col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
  }
}
