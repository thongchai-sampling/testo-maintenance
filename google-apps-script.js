/**
 * ==============================================================================
 * ระบบบันทึกการสอบเทียบ บำรุงรักษา และทวนสอบเครื่องมือวัดประสิทธิภาพการเผาไหม้ Testo
 * ห้องปฏิบัติการตรวจวิเคราะห์ / ฝ่ายตรวจวิเคราะห์
 * บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด
 * Backend: Google Apps Script (GAS) - รองรับการบันทึกและดึงข้อมูลย้อนหลังครบ 3 แบบฟอร์ม
 * ==============================================================================
 */

// ⚙️ การตั้งค่าระบบ (Configuration)
const CONFIG = {
  // ชื่อ Sheet ฐานข้อมูล (Database Sheets)
  DB_FORM1_SHEET: "DB_Maintenance_Cal",
  DB_FORM2_SHEET: "DB_Repair_History",
  DB_FORM3_SHEET: "DB_Gas_Verification",
  LEGACY_SHEET: "Maintenance_Logs",

  // ชื่อ Sheet แบบฟอร์มมาตรฐานพร้อมพิมพ์ (Form Template Sheets)
  FORM1_TEMPLATE_SHEET: "Form1_สอบเทียบและบำรุงรักษา",
  FORM2_TEMPLATE_SHEET: "Form2_ประวัติเครื่องมือ",
  FORM3_TEMPLATE_SHEET: "Form3_Gas_Verification",

  // อีเมลผู้รับรายงานสรุปประจำสัปดาห์ (คั่นด้วย , หากมีหลายอีเมล)
  EMAIL_RECIPIENTS: "your-email@example.com",
  REPORT_LOOKBACK_DAYS: 7,
  TRIGGER_DAY_OF_WEEK: ScriptApp.WeekDay.MONDAY,
  TRIGGER_HOUR: 8
};

// ทะเบียนเครื่องมือเริ่มต้นของบริษัท (Master Data 6 เครื่อง)
const EQUIPMENT_MASTER = {
  "01563443": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 327-1", code: "" },
  "01807527": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 350XL", code: "" },
  "01859560": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 350XL", code: "" },
  "60378478": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 350New", code: "" },
  "63455616": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 350New", code: "" },
  "63455658": { name: "Flue Gas Analyzer", brand: "Testo", model: "Testo 350New", code: "" }
};

/**
 * สร้างเมนูพิเศษบนแถบเมนูของ Google Sheets เมื่อเปิดไฟล์
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 ระบบแบบฟอร์ม Testo (ETC 1992)")
    .addItem("🔄 สร้าง/รีเฟรชหน้าแบบฟอร์มทั้ง 3 ไฟล์", "initializeAllSheetsAndForms")
    .addItem("📄 ดึงข้อมูลตาม S/N และช่วงวันที่ที่ระบุในแบบฟอร์ม", "refreshAllFormViews")
    .addItem("📥 ดาวน์โหลดหน้าแบบฟอร์มปัจจุบัน (PDF / Excel)", "showDownloadDialogForActiveSheet")
    .addSeparator()
    .addItem("📧 ส่งอีเมลรายงานสรุปประจำสัปดาห์ทันที", "sendWeeklyDashboardReport")
    .addItem("⏰ ตั้งเวลาส่งอีเมลอัตโนมัติทุกวันจันทร์", "setupWeeklyTrigger")
    .addToUi();
}

/**
 * เมื่อผู้ใช้เปลี่ยนค่า S/N หรือช่วงวันที่ในหน้าแบบฟอร์ม ให้ดึงข้อมูลย้อนหลังมาแสดงทันที
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var a1 = e.range.getA1Notation();

  // Form 1: เปลี่ยน S/N (I4) หรือ ตั้งแต่วันที่ (I5) หรือ ถึงวันที่ (K5)
  if (sheetName === CONFIG.FORM1_TEMPLATE_SHEET && (a1 === "I4" || a1 === "I5" || a1 === "K5")) {
    var sn1 = String(sheet.getRange("I4").getValue()).trim();
    var start1 = sheet.getRange("I5").getValue();
    var end1 = sheet.getRange("K5").getValue();
    populateForm1BySerial(sn1, start1, end1);
  }
  // Form 2: เปลี่ยน S/N (E5) หรือ ตั้งแต่วันที่ (B6) หรือ ถึงวันที่ (E6)
  else if (sheetName === CONFIG.FORM2_TEMPLATE_SHEET && (a1 === "E5" || a1 === "B6" || a1 === "E6")) {
    var sn2 = String(sheet.getRange("E5").getValue()).trim();
    var start2 = sheet.getRange("B6").getValue();
    var end2 = sheet.getRange("E6").getValue();
    populateForm2BySerial(sn2, start2, end2);
  }
  // Form 3: เปลี่ยน S/N (F3) หรือ เปลี่ยนรายการย้อนหลัง (F4)
  else if (sheetName === CONFIG.FORM3_TEMPLATE_SHEET && (a1 === "F3" || a1 === "F4")) {
    var sn3 = String(sheet.getRange("F3").getValue()).trim();
    var selectedRec = a1 === "F4" ? String(sheet.getRange("F4").getValue()).trim() : "";
    populateForm3BySerialAndRecord(sn3, selectedRec);
  }
}

/**
 * ฟังก์ชันรับ GET Request (รองรับทั้งการเช็คสถานะ และการดึงข้อมูลย้อนหลัง)
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "query") {
      var res = queryHistoricalData({
        targetForm: e.parameter.targetForm || "FORM_1",
        serial: e.parameter.serial || "",
        startDate: e.parameter.startDate || "",
        endDate: e.parameter.endDate || ""
      });
      return createJsonResponse(res);
    }

    return createJsonResponse({
      status: "online",
      message: "Testo Maintenance & Calibration API (ETC 1992) is active.",
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * ฟังก์ชันหลักในการรับข้อมูลจากหน้าเว็บแอป (POST Request)
 * รองรับทั้งการบันทึกข้อมูลใหม่ และการค้นหาข้อมูลย้อนหลัง (action: QUERY_HISTORY)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "ไม่พบข้อมูลที่ส่งมา (Empty payload)" });
    }

    var data = JSON.parse(e.postData.contents);

    // =========================================================================
    // กรณีพิเศษ: ค้นหาและดึงข้อมูลย้อนหลังสำหรับแท็บที่ 4 (QUERY_HISTORY)
    // =========================================================================
    if (data.action === "QUERY_HISTORY") {
      var queryResult = queryHistoricalData(data);
      return createJsonResponse(queryResult);
    }

    var eq = data.equipment || {
      name: data.eqName || "Flue Gas Analyzer",
      model: data.model || "",
      serial: data.serial || "",
      code: data.eqCode || ""
    };

    var summaryMessages = [];

    // =========================================================================
    // กรณีที่ 1: บันทึกจากแท็บที่ 1 (แบบฟอร์มที่ 1: สอบเทียบและบำรุงรักษา)
    // =========================================================================
    if (data.formType === "FORM_1_MAINTENANCE_CAL" || !data.formType) {
      recordForm1Data(data, eq);
      recordLegacyChecklist(data, eq);
      populateForm1BySerial(eq.serial, "", "");
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มที่ 1 (สอบเทียบและบำรุงรักษา) เรียบร้อย");

      if (data.syncToForm2) {
        recordForm2Data({
          date: data.date,
          damageDetail: data.repairDetail || data.notes || "พบรายการต้องแก้ไขจากการตรวจสภาพเครื่องมือ",
          operator: data.operator || data.technician || "",
          supervisor: data.supervisor || "",
          notes: "เชื่อมโยงอัตโนมัติจากแบบฟอร์มที่ 1 (" + (data.result || "ไม่ผ่าน") + ")"
        }, eq);
        populateForm2BySerial(eq.serial, "", "");
        summaryMessages.push("✓ เชื่อมโยงข้อมูลเข้าแบบฟอร์มที่ 2 (ประวัติความเสียหาย/ซ่อมแซม) อัตโนมัติ");
      }
    }

    // =========================================================================
    // กรณีที่ 2: บันทึกจากแท็บที่ 2 (แบบฟอร์มที่ 2: ประวัติความเสียหาย/ซ่อมแซม)
    // =========================================================================
    else if (data.formType === "FORM_2_REPAIR_HISTORY") {
      recordForm2Data(data, eq);
      populateForm2BySerial(eq.serial, "", "");
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มที่ 2 (ประวัติเครื่องมือ/การซ่อมแซม) เรียบร้อย");
    }

    // =========================================================================
    // กรณีที่ 3: บันทึกจากแท็บที่ 3 (แบบฟอร์มที่ 3: Flue Gas Analyzer Verification)
    // =========================================================================
    else if (data.formType === "FORM_3_GAS_VERIFICATION") {
      recordForm3Data(data, eq);
      populateForm3BySerialAndRecord(eq.serial, "");
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มที่ 3 (Flue Gas Analyzer Verification) เรียบร้อย");

      if (data.syncToForm1) {
        recordForm1Data({
          subType: "verification",
          date: data.date,
          nextDate: data.nextDate || "",
          description: "ทวนสอบค่าก๊าซมาตรฐาน (Flue Gas Verification: NO, NO2, SO2)",
          correction: "Max Error: " + (data.maxErrorPct || "-"),
          uncertainty: "-",
          tolerance: data.tolerancePct || "±5%",
          result: data.result || "ผ่าน",
          agency: "ภายในห้องปฏิบัติการ (ETC)",
          operator: data.calibratedBy || "",
          supervisor: data.approvedBy || "",
          notes: "Temp: " + (data.temp || "-") + "°C, P: " + (data.pressure || "-") + " Mbar, RH: " + (data.humidity || "-") + "%"
        }, eq);
        populateForm1BySerial(eq.serial, "", "");
        summaryMessages.push("✓ ส่งผลสรุปเข้าแบบฟอร์มที่ 1 (บันทึกการสอบเทียบและบำรุงรักษา) อัตโนมัติ");
      }
    }

    return createJsonResponse({
      status: "success",
      message: "Data recorded successfully",
      summary: summaryMessages.join("\n")
    });

  } catch (error) {
    Logger.log("doPost Error: " + error.toString());
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// ส่วนค้นหาและดึงข้อมูลย้อนหลังสำหรับหน้าเว็บแอป (Query Historical Data API)
// =============================================================================

function queryHistoricalData(params) {
  var targetForm = params.targetForm || "FORM_1";
  var cleanSerial = String(params.serial || "").replace(/^'/, "").trim();
  var startDt = parseDateValue(params.startDate);
  var endDt = parseDateValue(params.endDate);
  if (endDt) {
    endDt = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate(), 23, 59, 59);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var records = [];
  var dbSheetName = CONFIG.DB_FORM1_SHEET;
  if (targetForm === "FORM_2") dbSheetName = CONFIG.DB_FORM2_SHEET;
  if (targetForm === "FORM_3") dbSheetName = CONFIG.DB_FORM3_SHEET;

  var eqInfo = getEquipmentInfoBySerial(cleanSerial || "01563443", dbSheetName);
  var sheet = ss.getSheetByName(dbSheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      status: "success",
      targetForm: targetForm,
      equipment: eqInfo,
      records: []
    };
  }

  var allData = sheet.getDataRange().getValues();

  // ดึงข้อมูลแบบฟอร์มที่ 1
  if (targetForm === "FORM_1") {
    var seq1 = 1;
    for (var i = 1; i < allData.length; i++) {
      var r = allData[i];
      var rowSn = String(r[4] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSn !== cleanSerial) continue;

      var recDate = parseDateValue(r[5]);
      if (startDt && recDate && recDate < startDt) continue;
      if (endDt && recDate && recDate > endDt) continue;

      records.push({
        seq: seq1++,
        timestamp: r[0] ? formatDateDisplay(new Date(r[0])) : "",
        eqName: r[1] || "Flue Gas Analyzer",
        eqModel: r[2] || "",
        eqCode: r[3] || "",
        serial: rowSn,
        date: formatCellDate(r[5]),
        rawDate: formatIsoDate(r[5]),
        nextDate: formatCellDate(r[6]),
        description: r[7] || "",
        correction: r[8] || "-",
        uncertainty: r[9] || "-",
        tolerance: r[10] || "-",
        result: r[11] || "ผ่าน",
        agency: r[12] || "",
        operator: r[13] || "",
        supervisor: r[14] || "",
        notes: r[15] || ""
      });
    }
  }

  // ดึงข้อมูลแบบฟอร์มที่ 2
  else if (targetForm === "FORM_2") {
    var seq2 = 1;
    for (var j = 1; j < allData.length; j++) {
      var r2 = allData[j];
      var rowSn2 = String(r2[4] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSn2 !== cleanSerial) continue;

      var recDate2 = parseDateValue(r2[5]);
      if (startDt && recDate2 && recDate2 < startDt) continue;
      if (endDt && recDate2 && recDate2 > endDt) continue;

      records.push({
        seq: seq2++,
        timestamp: r2[0] ? formatDateDisplay(new Date(r2[0])) : "",
        eqName: r2[1] || "Flue Gas Analyzer",
        eqModel: r2[2] || "",
        eqCode: r2[3] || "",
        serial: rowSn2,
        date: formatCellDate(r2[5]),
        rawDate: formatIsoDate(r2[5]),
        damageDetail: r2[6] || "",
        operator: r2[7] || "",
        supervisor: r2[8] || "",
        notes: r2[9] || ""
      });
    }
  }

  // ดึงข้อมูลแบบฟอร์มที่ 3
  else if (targetForm === "FORM_3") {
    var seq3 = 1;
    for (var k = 1; k < allData.length; k++) {
      var r3 = allData[k];
      var rowSn3 = String(r3[4] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSn3 !== cleanSerial) continue;

      var recDate3 = parseDateValue(r3[1]);
      if (startDt && recDate3 && recDate3 < startDt) continue;
      if (endDt && recDate3 && recDate3 > endDt) continue;

      var parsedGas = { NO: [], NO2: [], SO2: [] };
      try {
        parsedGas = JSON.parse(r3[13] || "{}");
      } catch (err) {
        parsedGas = { NO: [], NO2: [], SO2: [] };
      }

      records.push({
        seq: seq3++,
        rowIndex: k + 1,
        date: formatCellDate(r3[1]),
        rawDate: formatIsoDate(r3[1]),
        time: formatCellTime(r3[2]),
        eqModel: r3[3] || eqInfo.model,
        serial: rowSn3,
        eqCode: r3[5] || eqInfo.code,
        temp: r3[6] !== undefined && r3[6] !== "" ? String(r3[6]) : "-",
        pressure: r3[7] !== undefined && r3[7] !== "" ? String(r3[7]) : "-",
        humidity: r3[8] !== undefined && r3[8] !== "" ? String(r3[8]) : "-",
        calibratedBy: r3[9] || "",
        approvedBy: r3[10] || "",
        result: r3[11] || "ผ่าน",
        maxErrorPct: r3[12] || "-",
        gasData: parsedGas
      });
    }
  }

  return {
    status: "success",
    targetForm: targetForm,
    equipment: eqInfo,
    records: records
  };
}

// =============================================================================
// ส่วนจัดการฐานข้อมูล (Database Sheets)
// =============================================================================

function recordForm1Data(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DB_FORM1_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DB_FORM1_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp", "ชื่อเครื่องมือ", "ยี่ห้อ/รุ่น", "หมายเลขเครื่องมือ (Code)", "หมายเลขเครื่อง (S/N)",
      "ว/ด/ป (ครั้งนี้)", "ว/ด/ป (ครั้งต่อไป)", "บันทึกการสอบเทียบ / การตรวจสภาพเครื่องมือ",
      "Correction / Error", "Uncertainty", "เกณฑ์การยอมรับ (Tolerance)", "ผลการสอบเทียบ (ผ่าน/ไม่ผ่าน)",
      "หน่วยงานให้บริการ", "ผู้ดำเนินการ", "ผู้รับผิดชอบ", "หมายเหตุ"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#1e40af");
  }

  sheet.appendRow([
    new Date(),
    eq.name || "Flue Gas Analyzer",
    eq.model || "",
    eq.code || "",
    "'" + String(eq.serial || "").trim(),
    data.date || "",
    data.nextDate || "",
    data.description || "",
    data.correction || "-",
    data.uncertainty || "-",
    data.tolerance || "-",
    data.result || "ผ่าน",
    data.agency || "ภายในห้องปฏิบัติการ (ETC)",
    data.operator || "",
    data.supervisor || "",
    data.notes || ""
  ]);
}

function recordForm2Data(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DB_FORM2_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DB_FORM2_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp", "Equipment Name", "Brand/Model", "Equipment Code", "Serial No. (S/N)",
      "ว/ด/ป", "รายการเสียหาย / ซ่อมแซม", "ผู้ดำเนินการ", "ผู้รับผิดชอบ", "หมายเหตุ"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#b45309");
  }

  sheet.appendRow([
    new Date(),
    eq.name || "Flue Gas Analyzer",
    eq.model || "",
    eq.code || "",
    "'" + String(eq.serial || "").trim(),
    data.date || "",
    data.damageDetail || "",
    data.operator || "",
    data.supervisor || "",
    data.notes || ""
  ]);
}

function recordForm3Data(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DB_FORM3_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DB_FORM3_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp", "Calibrated Date", "Calibrated Time", "Model", "Serial No.",
      "Equipment Code", "Temp (°C)", "Pressure (Mbar)", "Humidity (%RH)",
      "Calibrated By", "Approve By", "Result", "Max Error (%)", "GasData_JSON"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#3730a3");
  }

  sheet.appendRow([
    new Date(),
    data.date || "",
    data.time || "",
    eq.model || "",
    "'" + String(eq.serial || "").trim(),
    eq.code || "",
    data.temp || "-",
    data.pressure || "-",
    data.humidity || "-",
    data.calibratedBy || "",
    data.approvedBy || "",
    data.result || "ผ่าน",
    data.maxErrorPct || "-",
    JSON.stringify(data.gasData || {})
  ]);
}

function recordLegacyChecklist(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LEGACY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.LEGACY_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "วันเวลาที่บันทึก (Timestamp)", "วันที่ซ่อมบำรุง", "ชื่อช่าง", "รุ่นเครื่อง", "Serial No.",
      "กรองฝุ่น", "ปั๊มอากาศ", "แบตเตอรี่", "ความสะอาด", "เซ็นเซอร์", "หมายเหตุ / รายละเอียด"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#1e40af");
  }

  sheet.appendRow([
    new Date(),
    data.date || "",
    data.operator || data.technician || "",
    eq.model || "",
    "'" + String(eq.serial || "").trim(),
    data.chk_filter || "ปกติ",
    data.chk_pump || "ปกติ",
    data.chk_battery || "ปกติ",
    data.chk_clean || "สะอาด/ปกติ",
    data.chk_sensor || "ปกติ",
    data.notes || data.description || ""
  ]);
}

function styleDbHeader(sheet, colCount, bgColor) {
  var range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground(bgColor)
       .setFontColor("#ffffff")
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
}

// =============================================================================
// ส่วนจัดการและกรอกข้อมูลลงในแบบฟอร์มทั้ง 3 ไฟล์ (รองรับการกรองย้อนหลัง)
// =============================================================================

function initializeAllSheetsAndForms() {
  populateForm1BySerial("01563443", "", "");
  populateForm2BySerial("01563443", "", "");
  populateForm3BySerialAndRecord("01563443", "");
}

function refreshAllFormViews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s1 = ss.getSheetByName(CONFIG.FORM1_TEMPLATE_SHEET);
  if (s1) {
    populateForm1BySerial(
      String(s1.getRange("I4").getValue()).trim() || "01563443",
      s1.getRange("I5").getValue(),
      s1.getRange("K5").getValue()
    );
  } else {
    populateForm1BySerial("01563443", "", "");
  }

  var s2 = ss.getSheetByName(CONFIG.FORM2_TEMPLATE_SHEET);
  if (s2) {
    populateForm2BySerial(
      String(s2.getRange("E5").getValue()).trim() || "01563443",
      s2.getRange("B6").getValue(),
      s2.getRange("E6").getValue()
    );
  } else {
    populateForm2BySerial("01563443", "", "");
  }

  var s3 = ss.getSheetByName(CONFIG.FORM3_TEMPLATE_SHEET);
  if (s3) {
    populateForm3BySerialAndRecord(
      String(s3.getRange("F3").getValue()).trim() || "01563443",
      String(s3.getRange("F4").getValue()).trim()
    );
  } else {
    populateForm3BySerialAndRecord("01563443", "");
  }
}

/**
 * 1. แบบฟอร์มที่ 1: บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ (รองรับกรอง S/N และช่วงวันที่)
 */
function populateForm1BySerial(targetSerial, startDateVal, endDateVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FORM1_TEMPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.FORM1_TEMPLATE_SHEET);

  var cleanSerial = String(targetSerial || "01563443").replace(/^'/, "").trim();
  var query = queryHistoricalData({
    targetForm: "FORM_1",
    serial: cleanSerial,
    startDate: startDateVal,
    endDate: endDateVal
  });
  var eqInfo = query.equipment;

  sheet.clear();

  // หัวกระดาษ
  sheet.getRange("A1:L1").merge().setValue("บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ")
       .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:L2").merge().setValue("ห้องปฏิบัติการตรวจวิเคราะห์ บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด")
       .setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A3:B3").merge().setValue("ชื่อเครื่องมือ").setFontWeight("bold");
  sheet.getRange("C3:F3").merge().setValue(eqInfo.name);
  sheet.getRange("G3:H3").merge().setValue("ยี่ห้อ/รุ่น").setFontWeight("bold");
  sheet.getRange("I3:L3").merge().setValue(eqInfo.model);

  sheet.getRange("A4:B4").merge().setValue("หมายเลขเครื่องมือ").setFontWeight("bold");
  sheet.getRange("C4:F4").merge().setValue(eqInfo.code || "-");
  sheet.getRange("G4:H4").merge().setValue("หมายเลขเครื่อง (S/N)").setFontWeight("bold");
  sheet.getRange("I4").setValue("'" + cleanSerial).setBackground("#fef9c3").setFontWeight("bold");

  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["01563443", "01807527", "01859560", "60378478", "63455616", "63455658"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("I4").setDataValidation(snRule);
  sheet.getRange("J4:L4").merge().setValue("◀ เลือก S/N (สีเหลือง)").setFontColor("#64748b").setFontSize(9);

  // แถวที่ 5: ตัวกรองช่วงวันที่ย้อนหลัง
  sheet.getRange("G5:H5").merge().setValue("กรองตั้งแต่วันที่:").setFontColor("#475569").setFontSize(9).setHorizontalAlignment("right");
  sheet.getRange("I5").setValue(startDateVal || "").setBackground("#f0f9ff").setFontSize(9);
  sheet.getRange("J5").setValue("ถึงวันที่:").setFontColor("#475569").setFontSize(9).setHorizontalAlignment("right");
  sheet.getRange("K5").setValue(endDateVal || "").setBackground("#f0f9ff").setFontSize(9);
  sheet.getRange("L5").setValue("(ว่าง = ทั้งหมด)").setFontColor("#94a3b8").setFontSize(8);

  // หัวตาราง (แถว 6-7)
  sheet.getRange("A6:A7").merge().setValue("ลำดับที่");
  sheet.getRange("B6:C6").merge().setValue("ว/ด/ป");
  sheet.getRange("B7").setValue("ครั้งนี้");
  sheet.getRange("C7").setValue("ครั้งต่อไป");
  sheet.getRange("D6:D7").merge().setValue("บันทึกการสอบเทียบ /\nการตรวจสภาพเครื่องมือ");
  sheet.getRange("E6:E7").merge().setValue("Correction/\nError");
  sheet.getRange("F6:F7").merge().setValue("Uncertainty");
  sheet.getRange("G6:G7").merge().setValue("เกณฑ์การยอมรับ\n(Tolerance)");
  sheet.getRange("H6:I6").merge().setValue("ผลการสอบเทียบ");
  sheet.getRange("H7").setValue("ผ่าน");
  sheet.getRange("I7").setValue("ไม่ผ่าน");
  sheet.getRange("J6:J7").merge().setValue("หน่วยงาน\nให้บริการ");
  sheet.getRange("K6:K7").merge().setValue("ผู้ดำเนินการ/\nผู้รับผิดชอบ");
  sheet.getRange("L6:L7").merge().setValue("หมายเหตุ");

  sheet.getRange("A6:L7")
       .setBackground("#e2e8f0")
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setWrap(true)
       .setBorder(true, true, true, true, true, true);

  var rowsToInsert = [];
  for (var i = 0; i < query.records.length; i++) {
    var rec = query.records[i];
    var isPass = String(rec.result).indexOf("ไม่ผ่าน") === -1;
    var opAndSup = rec.operator || "";
    if (rec.supervisor) opAndSup = opAndSup ? (opAndSup + " / " + rec.supervisor) : rec.supervisor;

    rowsToInsert.push([
      rec.seq,
      rec.date,
      rec.nextDate,
      rec.description,
      rec.correction,
      rec.uncertainty,
      rec.tolerance,
      isPass ? "✓" : "",
      !isPass ? "✓" : "",
      rec.agency,
      opAndSup,
      rec.notes
    ]);
  }

  var minRows = Math.max(rowsToInsert.length, 15);
  while (rowsToInsert.length < minRows) {
    rowsToInsert.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
  }

  sheet.getRange(8, 1, rowsToInsert.length, 12)
       .setValues(rowsToInsert)
       .setVerticalAlignment("middle")
       .setWrap(true)
       .setBorder(true, true, true, true, true, true);

  sheet.getRange(8, 1, rowsToInsert.length, 3).setHorizontalAlignment("center");
  sheet.getRange(8, 5, rowsToInsert.length, 5).setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 95);
  sheet.setColumnWidth(3, 95);
  sheet.setColumnWidth(4, 240);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 85);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 55);
  sheet.setColumnWidth(9, 55);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 140);
  sheet.setColumnWidth(12, 160);
}

/**
 * 2. แบบฟอร์มที่ 2: บันทึกประวัติเครื่องมือ (รองรับกรอง S/N และช่วงวันที่)
 */
function populateForm2BySerial(targetSerial, startDateVal, endDateVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FORM2_TEMPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.FORM2_TEMPLATE_SHEET);

  var cleanSerial = String(targetSerial || "01563443").replace(/^'/, "").trim();
  var query = queryHistoricalData({
    targetForm: "FORM_2",
    serial: cleanSerial,
    startDate: startDateVal,
    endDate: endDateVal
  });
  var eqInfo = query.equipment;

  sheet.clear();

  sheet.getRange("A1:F1").merge().setValue("บันทึกประวัติเครื่องมือ")
       .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:F2").merge().setValue("ฝ่ายตรวจวิเคราะห์  บริษัท  อีสเทิร์น ไทย คอนซัลติ้ง  1992  จำกัด")
       .setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A3:F3").merge().setValue("ประวัติความเสียหาย ความผิดปกติ หรือการซ่อมแซมเครื่องมือ")
       .setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A4").setValue("Equipment Name").setFontWeight("bold");
  sheet.getRange("B4:C4").merge().setValue(eqInfo.name);
  sheet.getRange("D4").setValue("Brand/Model").setFontWeight("bold");
  sheet.getRange("E4:F4").merge().setValue(eqInfo.model);

  sheet.getRange("A5").setValue("Equipment Code").setFontWeight("bold");
  sheet.getRange("B5:C5").merge().setValue(eqInfo.code || "-");
  sheet.getRange("D5").setValue("Serial No.(S/N)").setFontWeight("bold");
  sheet.getRange("E5").setValue("'" + cleanSerial).setBackground("#fef9c3").setFontWeight("bold");

  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["01563443", "01807527", "01859560", "60378478", "63455616", "63455658"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("E5").setDataValidation(snRule);
  sheet.getRange("F5").setValue("◀ เลือก S/N").setFontColor("#64748b").setFontSize(9);

  // แถวที่ 6: ตัวกรองช่วงวันที่ย้อนหลัง
  sheet.getRange("A6").setValue("กรองตั้งแต่วันที่:").setFontColor("#475569").setFontSize(9);
  sheet.getRange("B6").setValue(startDateVal || "").setBackground("#f0f9ff").setFontSize(9);
  sheet.getRange("D6").setValue("ถึงวันที่:").setFontColor("#475569").setFontSize(9);
  sheet.getRange("E6").setValue(endDateVal || "").setBackground("#f0f9ff").setFontSize(9);
  sheet.getRange("F6").setValue("(ว่าง = ทั้งหมด)").setFontColor("#94a3b8").setFontSize(8);

  var headers = ["ลำดับที่", "ว/ด/ป", "รายการเสียหาย / ซ่อมแซม", "ผู้ดำเนินการ", "ผู้รับผิดชอบ", "หมายเหตุ"];
  sheet.getRange(7, 1, 1, 6).setValues([headers])
       .setBackground("#e2e8f0")
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true);

  var rowsToInsert = [];
  for (var i = 0; i < query.records.length; i++) {
    var rec = query.records[i];
    rowsToInsert.push([
      rec.seq,
      rec.date,
      rec.damageDetail,
      rec.operator,
      rec.supervisor,
      rec.notes
    ]);
  }

  var minRows = Math.max(rowsToInsert.length, 15);
  while (rowsToInsert.length < minRows) {
    rowsToInsert.push(["", "", "", "", "", ""]);
  }

  sheet.getRange(8, 1, rowsToInsert.length, 6)
       .setValues(rowsToInsert)
       .setVerticalAlignment("middle")
       .setWrap(true)
       .setBorder(true, true, true, true, true, true);

  sheet.getRange(8, 1, rowsToInsert.length, 2).setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 65);
  sheet.setColumnWidth(2, 105);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 180);
}

/**
 * 3. แบบฟอร์มที่ 3: Flue Gas Analyzer Verification (รองรับเลือก S/N และเลือกรายการวันที่ย้อนหลัง)
 */
function populateForm3BySerialAndRecord(targetSerial, selectedRecordLabel) {
  var cleanSerial = String(targetSerial || "01563443").replace(/^'/, "").trim();
  var query = queryHistoricalData({
    targetForm: "FORM_3",
    serial: cleanSerial
  });
  var eqInfo = query.equipment;
  var records = query.records;

  var recordOptions = [];
  var chosenRecord = null;

  for (var i = records.length - 1; i >= 0; i--) {
    var r = records[i];
    var label = "#" + r.seq + " | " + r.date + " " + r.time + " (" + r.result + ")";
    recordOptions.push(label);
    if (selectedRecordLabel && label === selectedRecordLabel) {
      chosenRecord = r;
    }
  }

  if (!chosenRecord && records.length > 0) {
    chosenRecord = records[records.length - 1];
    selectedRecordLabel = recordOptions[0];
  }

  var payload = chosenRecord || {
    date: "-",
    time: "-",
    temp: "-",
    pressure: "-",
    humidity: "-",
    calibratedBy: "........................................",
    approvedBy: "........................................",
    gasData: { NO: [], NO2: [], SO2: [] }
  };

  renderForm3Sheet(payload, eqInfo, recordOptions, selectedRecordLabel || "ล่าสุด");
}

function renderForm3Sheet(data, eq, recordOptions, currentLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FORM3_TEMPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.FORM3_TEMPLATE_SHEET);

  var cleanSerial = String(eq.serial || "01563443").replace(/^'/, "").trim();
  sheet.clear();

  sheet.getRange("A1:D1").merge().setValue("Flue Gas Analyzer Verification")
       .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A2").setValue("Calibrated Date :").setFontWeight("bold");
  sheet.getRange("B2").setValue(data.date || "-");
  sheet.getRange("C2").setValue("Calibrated Time :").setFontWeight("bold");
  sheet.getRange("D2").setValue(data.time || "-");

  sheet.getRange("A3").setValue("Flue Gas Analyzer Model :").setFontWeight("bold");
  sheet.getRange("B3").setValue(eq.model || "");
  sheet.getRange("C3").setValue("Serial No. :").setFontWeight("bold");
  sheet.getRange("D3").setValue("'" + cleanSerial).setFontWeight("bold");

  // กล่องเลือก S/N และเลือกรายการย้อนหลังด้านข้าง (คอลัมน์ E-F)
  sheet.getRange("E2:F2").merge().setValue("🔍 ค้นหาใบ Verification ย้อนหลัง")
       .setFontWeight("bold").setBackground("#dbeafe").setHorizontalAlignment("center");
  sheet.getRange("E3").setValue("1. เลือก S/N:").setFontWeight("bold").setFontSize(9);
  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["01563443", "01807527", "01859560", "60378478", "63455616", "63455658"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("F3").setValue("'" + cleanSerial).setDataValidation(snRule).setBackground("#fef9c3").setFontWeight("bold");

  sheet.getRange("E4").setValue("2. เลือกครั้งที่ทวนสอบ:").setFontWeight("bold").setFontSize(9);
  sheet.getRange("F4").setValue(currentLabel).setBackground("#fef9c3").setFontSize(9);
  if (recordOptions && recordOptions.length > 0) {
    var recRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(recordOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange("F4").setDataValidation(recRule);
  }

  var envText = "Environment Status   Temperature : " + (data.temp || "........") + " °C     " +
                "Pressure : " + (data.pressure || "........") + " Mbar     " +
                "Humidity : " + (data.humidity || "........") + " %RH";
  sheet.getRange("A4:D4").merge().setValue(envText).setFontWeight("bold");

  var gasList = [
    { key: "NO", title: "NO Calibrated Data", startRow: 6 },
    { key: "NO2", title: "NO2 Calibrated Data", startRow: 12 },
    { key: "SO2", title: "SO2 Calibrated Data", startRow: 18 }
  ];

  var gasObj = data.gasData || {};

  for (var g = 0; g < gasList.length; g++) {
    var item = gasList[g];
    var rIdx = item.startRow;

    sheet.getRange(rIdx, 1, 1, 4).merge().setValue(item.title)
         .setFontWeight("bold").setBackground("#f1f5f9");

    var colHeaders = [
      "Standard Gas Value (PPM)",
      "Flue Gas Analyzer Response (PPM)",
      "Difference (PPM)",
      "Calibration Error (%)"
    ];
    sheet.getRange(rIdx + 1, 1, 1, 4).setValues([colHeaders])
         .setBackground("#e2e8f0")
         .setFontWeight("bold")
         .setHorizontalAlignment("center")
         .setBorder(true, true, true, true, true, true);

    var rows = gasObj[item.key] || [];
    var tableRows = [];
    for (var k = 0; k < 3; k++) {
      var rowData = rows[k] || {};
      tableRows.push([
        rowData.std !== undefined && rowData.std !== "" ? rowData.std : "",
        rowData.res !== undefined && rowData.res !== "" ? rowData.res : "",
        rowData.diff !== undefined && rowData.diff !== "" ? rowData.diff : "",
        rowData.err !== undefined && rowData.err !== "" ? rowData.err : ""
      ]);
    }

    sheet.getRange(rIdx + 2, 1, 3, 4).setValues(tableRows)
         .setHorizontalAlignment("center")
         .setBorder(true, true, true, true, true, true);
  }

  sheet.getRange("A24:B24").merge().setValue("Calibrated By .......................................................")
       .setHorizontalAlignment("center");
  sheet.getRange("C24:D24").merge().setValue("Approve By .......................................................")
       .setHorizontalAlignment("center");

  sheet.getRange("A25:B25").merge().setValue("( " + (data.calibratedBy || "........................................") + " )")
       .setHorizontalAlignment("center");
  sheet.getRange("C25:D25").merge().setValue("( " + (data.approvedBy || "........................................") + " )")
       .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 210);
}

/**
 * แสดงหน้าต่างป๊อปอัปใน Google Sheets เพื่อกดดาวน์โหลดแบบฟอร์มปัจจุบันเป็น PDF หรือ Excel (.xlsx)
 */
function showDownloadDialogForActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ssId = ss.getId();
  var sheetId = sheet.getSheetId();
  var sheetName = sheet.getName();

  var isLandscape = (sheetName === CONFIG.FORM1_TEMPLATE_SHEET);
  var exportRange = "";
  if (sheetName === CONFIG.FORM3_TEMPLATE_SHEET) {
    exportRange = "&r1=0&c1=0&r2=26&c2=4"; // พิมพ์เฉพาะคอลัมน์ A-D ของใบ Verification
  }

  var pdfUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf" +
               "&gid=" + sheetId +
               "&size=A4" +
               "&portrait=" + (!isLandscape) +
               "&fitw=true" +
               "&gridlines=false" +
               exportRange;

  var xlsxUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx&gid=" + sheetId;

  var html = HtmlService.createHtmlOutput(
    "<div style='font-family:sans-serif; padding:12px; text-align:center;'>" +
      "<h3 style='margin-top:0; color:#1e3a8a;'>ดาวน์โหลดเอกสาร: " + sheetName + "</h3>" +
      "<p style='font-size:13px; color:#475569;'>เลือกประเภทไฟล์ที่ต้องการดาวน์โหลดลงเครื่องคอมพิวเตอร์:</p>" +
      "<div style='margin:18px 0;'>" +
        "<a href='" + pdfUrl + "' target='_blank' style='display:block; background:#dc2626; color:#fff; text-decoration:none; padding:10px 16px; border-radius:8px; font-weight:bold; margin-bottom:10px;'>📄 ดาวน์โหลดเป็นไฟล์ PDF (ขนาด A4)</a>" +
        "<a href='" + xlsxUrl + "' target='_blank' style='display:block; background:#16a34a; color:#fff; text-decoration:none; padding:10px 16px; border-radius:8px; font-weight:bold;'>📊 ดาวน์โหลดเป็นไฟล์ Excel (.xlsx)</a>" +
      "</div>" +
    "</div>"
  ).setWidth(380).setHeight(240);

  SpreadsheetApp.getUi().showModalDialog(html, "📥 ดาวน์โหลดแบบฟอร์ม");
}

function getEquipmentInfoBySerial(serial, dbSheetName) {
  var clean = String(serial || "").replace(/^'/, "").trim();
  var info = EQUIPMENT_MASTER[clean]
    ? {
        name: EQUIPMENT_MASTER[clean].name,
        model: EQUIPMENT_MASTER[clean].model,
        code: EQUIPMENT_MASTER[clean].code,
        serial: clean
      }
    : {
        name: "Flue Gas Analyzer",
        model: "Testo",
        code: "",
        serial: clean
      };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(dbSheetName);
  if (dbSheet && dbSheet.getLastRow() > 1) {
    var data = dbSheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var rowSn = String(data[i][4] || "").replace(/^'/, "").trim();
      if (rowSn === clean) {
        if (dbSheetName === CONFIG.DB_FORM3_SHEET) {
          if (data[i][3]) info.model = data[i][3];
          if (data[i][5]) info.code = data[i][5];
        } else {
          if (data[i][1]) info.name = data[i][1];
          if (data[i][2]) info.model = data[i][2];
          if (data[i][3]) info.code = data[i][3];
        }
        break;
      }
    }
  }
  return info;
}

function formatCellDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split("-");
    return p[2] + "/" + p[1] + "/" + p[0];
  }
  return s;
}

function formatIsoDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val).trim();
}

function formatCellTime(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(val).trim();
}

// =============================================================================
// ส่วนรายงาน PDF สรุปประจำสัปดาห์ทางอีเมล (Weekly Dashboard Report)
// =============================================================================

function sendWeeklyDashboardReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LEGACY_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var cutoffDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - CONFIG.REPORT_LOOKBACK_DAYS, 0, 0, 0);

  var records = [];
  var issueCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var recordDate = parseDateValue(row[1]);
    if (!recordDate) continue;

    if (recordDate >= cutoffDate) {
      var hasIssue = [row[5], row[6], row[7], row[8], row[9]].some(function(val) {
        var str = String(val);
        return str.indexOf("ต้องแก้ไข") !== -1 || str.indexOf("เสื่อม") !== -1 || str.indexOf("ผิดปกติ") !== -1;
      });
      if (hasIssue) issueCount++;

      records.push({
        date: formatDateDisplay(recordDate),
        technician: row[2] || "-",
        model: row[3] || "-",
        serial: String(row[4] || "-").replace(/^'/, ""),
        filter: row[5] || "-",
        pump: row[6] || "-",
        battery: row[7] || "-",
        clean: row[8] || "-",
        sensor: row[9] || "-",
        notes: row[10] || "-",
        hasIssue: hasIssue
      });
    }
  }

  var htmlContent = buildWeeklyHtmlReport(records, issueCount, cutoffDate, today);
  var blob = Utilities.newBlob(htmlContent, MimeType.HTML)
    .setName("Testo_Weekly_Maintenance_Report_" + Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd") + ".pdf");
  var pdfBlob = blob.getAs(MimeType.PDF);

  var emailSubject = (issueCount > 0 ? "⚠️ [ต้องติดตาม] " : "✅ [ปกติ] ") +
    "รายงานสรุปการซ่อมบำรุงเครื่อง Testo ประจำสัปดาห์ (" +
    formatDateDisplay(cutoffDate) + " - " + formatDateDisplay(today) + ")";

  var emailBody = "<div style='font-family: sans-serif; font-size: 14px; line-height: 1.6; color: #333;'>" +
    "<p>เรียน ทีมงานและผู้เกี่ยวข้อง (บจก. อีสเทิร์น ไทย คอนซัลติ้ง 1992),</p>" +
    "<p>ระบบได้สรุปผลการตรวจสอบและบำรุงรักษาเครื่องวัดประสิทธิภาพการเผาไหม้ Testo ในรอบ 7 วันที่ผ่านมา ดังนี้:</p>" +
    "<ul>" +
      "<li><b>จำนวนครั้งที่ตรวจสอบ:</b> " + records.length + " รายการ</li>" +
      "<li><b>พบรายการที่ต้องแก้ไข:</b> <span style='color:" + (issueCount > 0 ? "#dc2626" : "#16a34a") + "; font-weight:bold;'>" + issueCount + " รายการ</span></li>" +
    "</ul>" +
    "<p>รายละเอียดครบถ้วนระบุอยู่ในเอกสาร PDF ที่แนบมาพร้อมอีเมลฉบับนี้</p>" +
    "</div>";

  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: emailSubject,
    htmlBody: emailBody,
    attachments: [pdfBlob]
  });
}

function buildWeeklyHtmlReport(records, issueCount, startDate, endDate) {
  function renderBadge(val) {
    if (!val || val === "-") return "<span style='color:#666;'>-</span>";
    var str = String(val);
    if (str.indexOf("ต้องแก้ไข") !== -1 || str.indexOf("เสื่อม") !== -1 || str.indexOf("ผิดปกติ") !== -1) {
      return "<span style='background:#fee2e2; color:#b91c1c; padding:3px 6px; border-radius:4px; font-weight:bold; font-size:11px;'>ต้องแก้ไข</span>";
    }
    return "<span style='background:#dcfce7; color:#15803d; padding:3px 6px; border-radius:4px; font-size:11px;'>ปกติ</span>";
  }

  var rowsHtml = "";
  if (records.length === 0) {
    rowsHtml = "<tr><td colspan='9' style='text-align:center; padding:20px; color:#6b7280;'>ไม่มีรายการซ่อมบำรุงในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var rowBg = r.hasIssue ? "#fff7ed" : (i % 2 === 0 ? "#ffffff" : "#f9fafb");
      rowsHtml += "<tr style='background-color:" + rowBg + ";'>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center; font-size:10px;'>" + r.date + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; font-size:10px;'>" + escapeXml(r.technician) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; font-size:10px;'>" + escapeXml(r.model) + "<br><small style='color:#6b7280;'>S/N: " + escapeXml(r.serial) + "</small></td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center;'>" + renderBadge(r.filter) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center;'>" + renderBadge(r.pump) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center;'>" + renderBadge(r.battery) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center;'>" + renderBadge(r.clean) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; text-align:center;'>" + renderBadge(r.sensor) + "</td>" +
        "<td style='padding:6px; border:1px solid #d1d5db; font-size:10px; color:#374151;'>" + escapeXml(r.notes) + "</td>" +
      "</tr>";
    }
  }

  return "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<style>" +
      "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');" +
      "body { font-family: 'Sarabun', Arial, sans-serif; color: #1f2937; margin: 0; padding: 15px; font-size: 11px; }" +
      ".header { text-align: center; border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 15px; }" +
      ".header h1 { font-size: 18px; margin: 0 0 5px 0; color: #1e40af; }" +
      "table { width: 100%; border-collapse: collapse; margin-top: 10px; }" +
      "th { background-color: #2563eb; color: #ffffff; padding: 7px 4px; font-size: 10px; border: 1px solid #1d4ed8; text-align: center; }" +
    "</style></head><body>" +
    "<div class='header'>" +
      "<h1>รายงานสรุปการซ่อมบำรุงเครื่องวัดประสิทธิภาพการเผาไหม้ Testo</h1>" +
      "<p>บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด | รอบวันที่: " + formatDateDisplay(startDate) + " ถึง " + formatDateDisplay(endDate) + "</p>" +
    "</div>" +
    "<table><thead><tr>" +
      "<th>วันที่</th><th>ผู้ตรวจเช็ค</th><th>รุ่น / S/N</th><th>กรอง</th><th>ปั๊ม</th><th>แบต</th><th>ความสะอาด</th><th>เซ็นเซอร์</th><th>หมายเหตุ</th>" +
    "</tr></thead><tbody>" + rowsHtml + "</tbody></table></body></html>";
}

function setupWeeklyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendWeeklyDashboardReport") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("sendWeeklyDashboardReport")
    .timeBased()
    .onWeekDay(CONFIG.TRIGGER_DAY_OF_WEEK)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();
}

function parseDateValue(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var str = String(val).trim();
  if (!str) return null;

  // รองรับ YYYY-MM-DD
  var dashParts = str.split("-");
  if (dashParts.length === 3 && dashParts[0].length === 4) {
    return new Date(parseInt(dashParts[0], 10), parseInt(dashParts[1], 10) - 1, parseInt(dashParts[2], 10));
  }
  // รองรับ DD/MM/YYYY
  var slashParts = str.split("/");
  if (slashParts.length === 3) {
    var y = parseInt(slashParts[2], 10);
    if (y > 2400) y -= 543; // แปลง พ.ศ. เป็น ค.ศ. หากกรอกเป็น พ.ศ.
    return new Date(y, parseInt(slashParts[1], 10) - 1, parseInt(slashParts[0], 10));
  }

  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateDisplay(date) {
  if (!date || isNaN(date.getTime())) return "-";
  var d = String(date.getDate()).padStart(2, "0");
  var m = String(date.getMonth() + 1).padStart(2, "0");
  var y = date.getFullYear();
  return d + "/" + m + "/" + y;
}

function escapeXml(unsafe) {
  if (unsafe === null || unsafe === undefined) return "";
  return String(unsafe)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
