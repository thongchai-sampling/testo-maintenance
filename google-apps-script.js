/**
 * ==============================================================================
 * ระบบบันทึกการสอบเทียบ บำรุงรักษา และทวนสอบเครื่องมือวัดประสิทธิภาพการเผาไหม้ Testo
 * แผนกปฏิบัติการภาคสนาม
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

  // ชื่อ Sheet ฐานข้อมูลเครื่องวัดระดับเสียง (Sound Level Meter Database Sheets)
  DB_SOUND_FORM1_SHEET: "DB_Sound_Maintenance_Cal",
  DB_SOUND_FORM2_SHEET: "DB_Sound_Repair_History",
  DB_SOUND_FORM3_SHEET: "DB_Sound_Verification",

  // ชื่อ Sheet ฐานข้อมูลเครื่องวัดความร้อน (Heat Stress / WBGT Database Sheets)
  DB_HEAT_FORM1_SHEET: "DB_Heat_Maintenance_Cal",
  DB_HEAT_FORM2_SHEET: "DB_Heat_Repair_History",
  DB_HEAT_FORM3_SHEET: "DB_Heat_Verification",

  // ชื่อ Sheet แบบฟอร์มมาตรฐานพร้อมพิมพ์ (Form Template Sheets)
  FORM1_TEMPLATE_SHEET: "Form1_สอบเทียบและบำรุงรักษา",
  FORM2_TEMPLATE_SHEET: "Form2_ประวัติเครื่องมือ",
  FORM3_TEMPLATE_SHEET: "Form3_Gas_Verification",
  SOUND_FORM1_TEMPLATE_SHEET: "Form1_Sound_สอบเทียบและบำรุงรักษา",
  SOUND_FORM2_TEMPLATE_SHEET: "Form2_Sound_ประวัติเครื่องมือ",
  SOUND_FORM3_TEMPLATE_SHEET: "Form3_Sound_Verification",
  HEAT_FORM1_TEMPLATE_SHEET: "Form1_Heat_สอบเทียบและบำรุงรักษา",
  HEAT_FORM2_TEMPLATE_SHEET: "Form2_Heat_ประวัติเครื่องมือ",
  HEAT_FORM3_TEMPLATE_SHEET: "Form3_Heat_Verification",

  // อีเมลผู้รับรายงานสรุปประจำสัปดาห์ (คั่นด้วย , หากมีหลายอีเมล)
  EMAIL_RECIPIENTS: "thongchai_b@etc1992.com",
  REPORT_LOOKBACK_DAYS: 7,
  TRIGGER_DAY_OF_WEEK: ScriptApp.WeekDay.MONDAY,
  TRIGGER_HOUR: 8,
  WEBAPP_URL: "https://thongchai-sampling.github.io/testo-maintenance/"
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

// ทะเบียนเครื่องวัดระดับเสียง (Sound Level Meter - ข้อมูลจริง 80 เครื่องจากทะเบียน Sampling)
const SOUND_EQUIPMENT_MASTER = {
  "090085": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020466", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ไม่ได้ขึ้นทะเบียนกรมสวัสดิฯ **ใช้ตรวจวัดLeq24hrเท่านั้น"},
  "090086": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020467", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ไม่ได้ขึ้นทะเบียนกรมสวัสดิฯ **ใช้ตรวจวัดLeq24hrเท่านั้น"},
  "090087": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020468", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-198/2567) รอซื้อทดแทน"},
  "090088": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020469", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ไม่ได้ขึ้นทะเบียนกรมสวัสดิฯ **ใช้ตรวจวัดLeq24hrเท่านั้น"},
  "090089": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020470", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-143/2567) ซื้อทดแทนด้วย S/N 00641700"},
  "090090": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020471", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-198/2567) รอซื้อทดแทน"},
  "090091": {"name": "Sound Level Meter", "brand": "ACO", "model": "ACO 6226", "code": "S52020471/1", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-143/2567) ซื้อทดแทนด้วย S/N 00641701"},
  "10185674": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-04", "code": "B4205052", "standard": "IEC 651 (Type 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-319/2567) ซื้อทดแทนด้วย S/N 00641702"},
  "00310455": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B40050001", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00310456": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B40050002", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00310458": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B40050003", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "00443357": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B47050194", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00443358": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B47050195", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "00443359": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "B47050196", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00209071": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S53020567", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "00209072": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S53020568", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "ไม่ได้ขึ้นทะเบียนกรมสวัสดิฯ **ใช้ตรวจวัดLeq24hrเท่านั้น"},
  "00209079": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S53020569", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01209912": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S54020591", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01209914": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S54020592", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01209915": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S54020593", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "01209916": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S54020594", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01209917": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-21", "code": "S54020595", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "00157456": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42", "code": "B58020030", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-143/2567) ซื้อทดแทนด้วย S/N 00641703"},
  "00646443": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42", "code": "B58020031", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "01147298": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42", "code": "B58020032", "standard": "IEC 61672 (Class 2)", "status": "รอซ่อม / ส่งซ่อม", "notes": ""},
  "01147299": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42", "code": "B58020033", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01147300": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42", "code": "B58020034", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "G300957": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S62021311", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301013": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S62021312", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301039": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S62021313", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301607": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S63021399", "standard": "IEC 61672 (Class 2)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "G301635": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S63021400", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301638": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S63021401", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301660": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S63021402", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "G301661": {"name": "Sound Level Meter", "brand": "Cirrus", "model": "Cirrus CR:172A", "code": "S63021403", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00222592": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021654", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00222593": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021655", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00222594": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021656", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322744": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021657", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322745": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021658", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322746": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021659", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322747": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021660", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322748": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021661", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322749": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021662", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322750": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021663", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322751": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021664", "standard": "IEC 61672 (Class 2)", "status": "รอซ่อม / ส่งซ่อม", "notes": ""},
  "00322752": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021665", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322753": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021666", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322754": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021667", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322755": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021668", "standard": "IEC 61672 (Class 2)", "status": "รอซ่อม / ส่งซ่อม", "notes": ""},
  "00322756": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021669", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00322757": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-42A", "code": "S66021670", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "01120943": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021678", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120944": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021679", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120945": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021680", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120946": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021681", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "ไม่ได้ขึ้นทะเบียนกรมสวัสดิฯ **ใช้ตรวจวัดLeq24hrเท่านั้น ** บันทึกเสียงได้"},
  "01120947": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021682", "standard": "IEC 61672 (Class 1)", "status": "รอซ่อม / ส่งซ่อม", "notes": "** บันทึกเสียงได้"},
  "01120948": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021683", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120949": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021684", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120950": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021685", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120952": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021686", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "01120953": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021687", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230985": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021688", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230986": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021689", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230987": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021690", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230988": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021691", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230989": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021692", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230991": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021693", "standard": "IEC 61672 (Class 1)", "status": "ชำรุด / ยกเลิกใช้งาน", "notes": "ยกเลิกทรัพย์สินแล้ว(MEMO-181/2568) รอซื้อทดแทน"},
  "00230992": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021694", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230993": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021695", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230994": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021696", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00230995": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-52A", "code": "S66021697", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00641700": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-43", "code": "S67021829", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00641701": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-43", "code": "S67021830", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00641702": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-43", "code": "S67021831", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00641703": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-43", "code": "S67021832", "standard": "IEC 61672 (Class 2)", "status": "พร้อมใช้งาน", "notes": ""},
  "00741217": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-53", "code": "S67021833", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00741218": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-53", "code": "S67021834", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00741219": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-53", "code": "S67021835", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"},
  "00741254": {"name": "Sound Level Meter", "brand": "RION", "model": "RION NL-53", "code": "S67021836", "standard": "IEC 61672 (Class 1)", "status": "พร้อมใช้งาน", "notes": "** บันทึกเสียงได้"}
};

// ทะเบียนเครื่องวัดความร้อน (Heat Stress / WBGT Meter - ข้อมูลจริง 18 เครื่องจากทะเบียน Sampling)
const HEAT_EQUIPMENT_MASTER = {
  "TPH 060001": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPI 050069": {"name": "Heat Stress Monitor", "brand": "Quest technologies", "model": "Quest technologies QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPI 050070": {"name": "Heat Stress Monitor", "brand": "Quest technologies", "model": "Quest technologies QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPL 060039": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPL 060040": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPL 090016": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPL 090017": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPQ 030023": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TPQ 030024": {"name": "Heat Stress Monitor", "brand": "3 M", "model": "3 M QUESTemp 32", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TEU 080011": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 34", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TEU 080012": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 34", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TEU 080013": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 34", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TEU 080014": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 34", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "TEU 080015": {"name": "Heat Stress Monitor", "brand": "TSI", "model": "TSI QUESTemp 34", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "22004316": {"name": "Heat Stress Monitor", "brand": "DeltaOHM", "model": "DeltaOHM HD32.2", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "22004318": {"name": "Heat Stress Monitor", "brand": "DeltaOHM", "model": "DeltaOHM HD32.2", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "22004319": {"name": "Heat Stress Monitor", "brand": "DeltaOHM", "model": "DeltaOHM HD32.2", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""},
  "22004320": {"name": "Heat Stress Monitor", "brand": "DeltaOHM", "model": "DeltaOHM HD32.2", "code": "", "standard": "ISO 7243", "status": "พร้อมใช้งาน", "notes": ""}
};

function getHeatSerialList() {
  return Object.keys(HEAT_EQUIPMENT_MASTER);
}

function resolveCategoryConfig(category) {
  var cat = String(category || "").toUpperCase();
  if (cat === "SOUND") {
    return {
      category: "SOUND",
      db1: CONFIG.DB_SOUND_FORM1_SHEET,
      db2: CONFIG.DB_SOUND_FORM2_SHEET,
      db3: CONFIG.DB_SOUND_FORM3_SHEET,
      tpl1: CONFIG.SOUND_FORM1_TEMPLATE_SHEET,
      tpl2: CONFIG.SOUND_FORM2_TEMPLATE_SHEET,
      tpl3: CONFIG.SOUND_FORM3_TEMPLATE_SHEET,
      defaultName: "Sound Level Meter",
      defaultSerial: "01120943"
    };
  }
  if (cat === "HEAT") {
    return {
      category: "HEAT",
      db1: CONFIG.DB_HEAT_FORM1_SHEET,
      db2: CONFIG.DB_HEAT_FORM2_SHEET,
      db3: CONFIG.DB_HEAT_FORM3_SHEET,
      tpl1: CONFIG.HEAT_FORM1_TEMPLATE_SHEET,
      tpl2: CONFIG.HEAT_FORM2_TEMPLATE_SHEET,
      tpl3: CONFIG.HEAT_FORM3_TEMPLATE_SHEET,
      defaultName: "Heat Stress Monitor",
      defaultSerial: "TPH 060001"
    };
  }
  return {
    category: "FLUE_GAS",
    db1: CONFIG.DB_FORM1_SHEET,
    db2: CONFIG.DB_FORM2_SHEET,
    db3: CONFIG.DB_FORM3_SHEET,
    tpl1: CONFIG.FORM1_TEMPLATE_SHEET,
    tpl2: CONFIG.FORM2_TEMPLATE_SHEET,
    tpl3: CONFIG.FORM3_TEMPLATE_SHEET,
    defaultName: "Flue Gas Analyzer",
    defaultSerial: "01563443"
  };
}

function getSoundSerialList() {
  var active = [];
  var others = [];
  var keys = Object.keys(SOUND_EQUIPMENT_MASTER);
  for (var i = 0; i < keys.length; i++) {
    var sn = keys[i];
    if (SOUND_EQUIPMENT_MASTER[sn].status === "พร้อมใช้งาน") {
      active.push(sn);
    } else {
      others.push(sn);
    }
  }
  return active.concat(others);
}

/**
 * สร้างเมนูพิเศษบนแถบเมนูของ Google Sheets เมื่อเปิดไฟล์
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 ระบบบำรุงรักษาเครื่องมือภาคสนาม (ETC 1992)")
    .addItem("🔄 สร้าง/รีเฟรชหน้าแบบฟอร์ม Flue Gas (Testo)", "initializeAllSheetsAndForms")
    .addItem("🔊 สร้าง/รีเฟรชหน้าแบบฟอร์มเครื่องวัดเสียง (Sound)", "initializeSoundSheetsAndForms")
    .addItem("🌡️ สร้าง/รีเฟรชหน้าแบบฟอร์มเครื่องวัดความร้อน (Heat Stress)", "initializeHeatSheetsAndForms")
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
  // Sound Form 1
  else if (sheetName === CONFIG.SOUND_FORM1_TEMPLATE_SHEET && (a1 === "I4" || a1 === "I5" || a1 === "K5")) {
    populateForm1BySerial(
      String(sheet.getRange("I4").getValue()).trim(),
      sheet.getRange("I5").getValue(),
      sheet.getRange("K5").getValue(),
      "SOUND"
    );
  }
  // Sound Form 2
  else if (sheetName === CONFIG.SOUND_FORM2_TEMPLATE_SHEET && (a1 === "E5" || a1 === "B6" || a1 === "E6")) {
    populateForm2BySerial(
      String(sheet.getRange("E5").getValue()).trim(),
      sheet.getRange("B6").getValue(),
      sheet.getRange("E6").getValue(),
      "SOUND"
    );
  }
  // Sound Form 3
  else if (sheetName === CONFIG.SOUND_FORM3_TEMPLATE_SHEET && (a1 === "F3" || a1 === "F4")) {
    var sSn3 = String(sheet.getRange("F3").getValue()).trim();
    var sSelectedRec = a1 === "F4" ? String(sheet.getRange("F4").getValue()).trim() : "";
    populateSoundForm3BySerialAndRecord(sSn3, sSelectedRec);
  }
  // Heat Form 1
  else if (sheetName === CONFIG.HEAT_FORM1_TEMPLATE_SHEET && (a1 === "I4" || a1 === "I5" || a1 === "K5")) {
    populateForm1BySerial(String(sheet.getRange("I4").getValue()).trim(), sheet.getRange("I5").getValue(), sheet.getRange("K5").getValue(), "HEAT");
  }
  // Heat Form 2
  else if (sheetName === CONFIG.HEAT_FORM2_TEMPLATE_SHEET && (a1 === "E5" || a1 === "B6" || a1 === "E6")) {
    populateForm2BySerial(String(sheet.getRange("E5").getValue()).trim(), sheet.getRange("B6").getValue(), sheet.getRange("E6").getValue(), "HEAT");
  }
  // Heat Form 3
  else if (sheetName === CONFIG.HEAT_FORM3_TEMPLATE_SHEET && (a1 === "F3" || a1 === "F4")) {
    var hSn3 = String(sheet.getRange("F3").getValue()).trim();
    var hSelectedRec = a1 === "F4" ? String(sheet.getRange("F4").getValue()).trim() : "";
    populateHeatForm3BySerialAndRecord(hSn3, hSelectedRec);
  }
}

/**
 * ฟังก์ชันรับ GET Request (รองรับทั้งการเช็คสถานะ และการดึงข้อมูลย้อนหลัง)
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "query") {
      var res = queryHistoricalData({
        category: e.parameter.category || "FLUE_GAS",
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

    var catCfg = resolveCategoryConfig(data.category);
    var eq = data.equipment || {
      name: data.eqName || catCfg.defaultName,
      model: data.model || "",
      serial: data.serial || "",
      code: data.eqCode || ""
    };

    var summaryMessages = [];

    // =========================================================================
    // กรณีที่ 1: บันทึกจากแท็บที่ 1 (แบบฟอร์มที่ 1: สอบเทียบและบำรุงรักษา)
    // =========================================================================
    if (data.formType === "FORM_1_MAINTENANCE_CAL" || !data.formType) {
      recordForm1Data(data, eq, catCfg.db1);
      recordLegacyChecklist(data, eq, catCfg.category);
      populateForm1BySerial(eq.serial, "", "", catCfg.category);
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มที่ 1 (สอบเทียบและบำรุงรักษา) เรียบร้อย");

      if (data.syncToForm2) {
        recordForm2Data({
          date: data.date,
          damageDetail: data.repairDetail || data.notes || "พบรายการต้องแก้ไขจากการตรวจสภาพเครื่องมือ",
          operator: data.operator || data.technician || "",
          supervisor: data.supervisor || "",
          notes: "เชื่อมโยงอัตโนมัติจากแบบฟอร์มที่ 1 (" + (data.result || "ไม่ผ่าน") + ")"
        }, eq, catCfg.db2);
        populateForm2BySerial(eq.serial, "", "", catCfg.category);
        summaryMessages.push("✓ เชื่อมโยงข้อมูลเข้าแบบฟอร์มที่ 2 (ประวัติความเสียหาย/ซ่อมแซม) อัตโนมัติ");
      }
    }

    // =========================================================================
    // กรณีที่ 2: บันทึกจากแท็บที่ 2 (แบบฟอร์มที่ 2: ประวัติความเสียหาย/ซ่อมแซม)
    // =========================================================================
    else if (data.formType === "FORM_2_REPAIR_HISTORY") {
      recordForm2Data(data, eq, catCfg.db2);
      populateForm2BySerial(eq.serial, "", "", catCfg.category);
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มที่ 2 (ประวัติเครื่องมือ/การซ่อมแซม) เรียบร้อย");
    }

    // =========================================================================
    // กรณีที่ 3A: บันทึก Flue Gas Analyzer Verification
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
          description: "ทวนสอบค่าก๊าซมาตรฐาน (Flue Gas Verification: O2, CO, NO, NO2, SO2)",
          correction: "Max Error: " + (data.maxErrorPct || "-"),
          uncertainty: "-",
          tolerance: data.tolerancePct || "±5%",
          result: data.result || "ผ่าน",
          agency: "ภายในแผนกปฏิบัติการภาคสนาม (ETC)",
          operator: data.calibratedBy || "",
          supervisor: data.approvedBy || "",
          notes: "Temp: " + (data.temp || "-") + "°C, P: " + (data.pressure || "-") + " Mbar, RH: " + (data.humidity || "-") + "%"
        }, eq, CONFIG.DB_FORM1_SHEET);
        populateForm1BySerial(eq.serial, "", "", "FLUE_GAS");
        summaryMessages.push("✓ ส่งผลสรุปเข้าแบบฟอร์มที่ 1 (บันทึกการสอบเทียบและบำรุงรักษา) อัตโนมัติ");
      }
    }

    // =========================================================================
    // กรณีที่ 3B: บันทึก Sound Level Meter Pre/Post Calibration Check
    // =========================================================================
    else if (data.formType === "FORM_3_SOUND_VERIFICATION") {
      recordSoundForm3Data(data, eq);
      populateSoundForm3BySerialAndRecord(eq.serial, "");
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มทวนสอบเครื่องวัดเสียง (Pre/Post Calibration Check) เรียบร้อย");

      if (data.syncToForm1) {
        recordForm1Data({
          subType: "verification",
          date: data.date,
          nextDate: data.nextDate || "",
          description: "ทวนสอบเครื่องวัดเสียง Pre/Post Calibration (" + (data.calPointsSummary || "94/114 dB") + ", " + (data.freq || "1 kHz") + ", " + (data.freqWeighting || "A") + "/" + (data.timeWeighting || "Fast") + ")",
          correction: "Max Dev: " + (data.maxDevDb || "-"),
          uncertainty: "Drift: " + (data.maxDriftDb || "-"),
          tolerance: data.toleranceDb || "±0.5 dB",
          result: data.result || "ผ่าน",
          agency: "ภายในแผนกปฏิบัติการภาคสนาม (ETC)",
          operator: data.calibratedBy || "",
          supervisor: data.approvedBy || "",
          notes: "Calibrator: " + (data.calibrator || "-") + (data.notes ? " | " + data.notes : "")
        }, eq, CONFIG.DB_SOUND_FORM1_SHEET);
        populateForm1BySerial(eq.serial, "", "", "SOUND");
        summaryMessages.push("✓ ส่งผลสรุปเข้าแบบฟอร์มที่ 1 (เครื่องวัดเสียง) อัตโนมัติ");
      }
    }

    // =========================================================================
    // กรณีที่ 3C: บันทึก Heat Stress Verification (Ta, Tw, Tg)
    // =========================================================================
    else if (data.formType === "FORM_3_HEAT_VERIFICATION") {
      recordHeatForm3Data(data, eq);
      populateHeatForm3BySerialAndRecord(eq.serial, "");
      summaryMessages.push("✓ บันทึกลงฐานข้อมูลและแบบฟอร์มทวนสอบเครื่องวัดความร้อน (Heat Stress Verification) เรียบร้อย");

      if (data.syncToForm1) {
        recordForm1Data({
          subType: "verification",
          date: data.date,
          nextDate: data.nextDate || "",
          description: "ทวนสอบเครื่องวัดความร้อน (" + (data.method || "Verification Module / Ref Thermometer") + ": Ta, Tw, Tg)",
          correction: "Max Err: " + (data.maxErrorC || "-"),
          uncertainty: "-",
          tolerance: data.toleranceC || "±0.5 °C",
          result: data.result || "ผ่าน",
          agency: "ภายในแผนกปฏิบัติการภาคสนาม (ETC)",
          operator: data.calibratedBy || "",
          supervisor: data.approvedBy || "",
          notes: "Ref: " + (data.refDevice || "-") + (data.notes ? " | " + data.notes : "")
        }, eq, CONFIG.DB_HEAT_FORM1_SHEET);
        populateForm1BySerial(eq.serial, "", "", "HEAT");
        summaryMessages.push("✓ ส่งผลสรุปเข้าแบบฟอร์มที่ 1 (เครื่องวัดความร้อน) อัตโนมัติ");
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
  var catCfg = resolveCategoryConfig(params.category);
  var targetForm = params.targetForm || "FORM_1";
  var cleanSerial = String(params.serial || "").replace(/^'/, "").trim();
  var startDt = parseDateValue(params.startDate);
  var endDt = parseDateValue(params.endDate);
  if (endDt) {
    endDt = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate(), 23, 59, 59);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var records = [];
  var dbSheetName = catCfg.db1;
  if (targetForm === "FORM_2") dbSheetName = catCfg.db2;
  if (targetForm === "FORM_3") dbSheetName = catCfg.db3;

  var eqInfo = getEquipmentInfoBySerial(cleanSerial || catCfg.defaultSerial, dbSheetName, catCfg.category);
  var sheet = ss.getSheetByName(dbSheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      status: "success",
      category: catCfg.category,
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

  // ดึงข้อมูลแบบฟอร์มที่ 3 (แยก Flue Gas vs Sound vs Heat Stress)
  else if (targetForm === "FORM_3" && catCfg.category === "HEAT") {
    var seqH = 1;
    for (var hIdx = 1; hIdx < allData.length; hIdx++) {
      var rh = allData[hIdx];
      var rowSnH = String(rh[3] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSnH !== cleanSerial) continue;

      var recDateH = parseDateValue(rh[1]);
      if (startDt && recDateH && recDateH < startDt) continue;
      if (endDt && recDateH && recDateH > endDt) continue;

      var parsedHeat = [];
      try {
        parsedHeat = JSON.parse(rh[13] || "[]");
      } catch (errH) {
        parsedHeat = [];
      }

      records.push({
        seq: seqH++,
        rowIndex: hIdx + 1,
        date: formatCellDate(rh[1]),
        rawDate: formatIsoDate(rh[1]),
        time: formatCellTime(rh[2]),
        serial: rowSnH,
        eqModel: rh[4] || eqInfo.model,
        eqCode: rh[5] || eqInfo.code,
        standard: rh[6] || eqInfo.standard || "ISO 7243",
        refDevice: rh[7] || "-",
        method: rh[8] || "-",
        toleranceC: rh[9] || "±0.5 °C",
        calibratedBy: rh[10] || "",
        approvedBy: rh[11] || "",
        result: rh[12] || "ผ่าน",
        heatData: parsedHeat,
        maxErrorC: rh[14] || "-",
        wbgtCheck: rh[15] || "-",
        notes: rh[16] || ""
      });
    }
  }
  else if (targetForm === "FORM_3" && catCfg.category === "SOUND") {
    var seqS = 1;
    for (var sIdx = 1; sIdx < allData.length; sIdx++) {
      var rs = allData[sIdx];
      var rowSnS = String(rs[4] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSnS !== cleanSerial) continue;

      var recDateS = parseDateValue(rs[1]);
      if (startDt && recDateS && recDateS < startDt) continue;
      if (endDt && recDateS && recDateS > endDt) continue;

      var parsedSound = [];
      try {
        parsedSound = JSON.parse(rs[16] || "[]");
      } catch (errS) {
        parsedSound = [];
      }

      records.push({
        seq: seqS++,
        rowIndex: sIdx + 1,
        date: formatCellDate(rs[1]),
        rawDate: formatIsoDate(rs[1]),
        preTime: formatCellTime(rs[2]),
        postTime: formatCellTime(rs[3]),
        serial: rowSnS,
        eqModel: rs[5] || eqInfo.model,
        eqCode: rs[6] || eqInfo.code,
        standard: rs[7] || eqInfo.standard || "",
        calibrator: rs[8] || "-",
        freq: rs[9] || "1 kHz",
        freqWeighting: rs[10] || "A-Weighting (dBA)",
        timeWeighting: rs[11] || "Fast (F)",
        toleranceDb: rs[12] || "±0.5 dB",
        calibratedBy: rs[13] || "",
        approvedBy: rs[14] || "",
        result: rs[15] || "ผ่าน",
        maxDevDb: rs[17] || "-",
        maxDriftDb: rs[18] || "-",
        notes: rs[19] || "",
        soundData: parsedSound
      });
    }
  }
  else if (targetForm === "FORM_3") {
    var seq3 = 1;
    for (var k = 1; k < allData.length; k++) {
      var r3 = allData[k];
      var rowSn3 = String(r3[4] || "").replace(/^'/, "").trim();
      if (cleanSerial && cleanSerial !== "ALL" && rowSn3 !== cleanSerial) continue;

      var recDate3 = parseDateValue(r3[1]);
      if (startDt && recDate3 && recDate3 < startDt) continue;
      if (endDt && recDate3 && recDate3 > endDt) continue;

      var parsedGas = { O2: [], CO: [], NO: [], NO2: [], SO2: [] };
      try {
        parsedGas = JSON.parse(r3[13] || "{}");
      } catch (err) {
        parsedGas = { O2: [], CO: [], NO: [], NO2: [], SO2: [] };
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
    category: catCfg.category,
    targetForm: targetForm,
    equipment: eqInfo,
    records: records
  };
}

// =============================================================================
// ส่วนจัดการฐานข้อมูล (Database Sheets)
// =============================================================================

function recordForm1Data(data, eq, targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sName = targetSheetName || CONFIG.DB_FORM1_SHEET;
  var sheet = ss.getSheetByName(sName);
  if (!sheet) sheet = ss.insertSheet(sName);

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
    data.agency || "ภายในแผนกปฏิบัติการภาคสนาม (ETC)",
    data.operator || "",
    data.supervisor || "",
    data.notes || ""
  ]);
}

function recordForm2Data(data, eq, targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sName = targetSheetName || CONFIG.DB_FORM2_SHEET;
  var sheet = ss.getSheetByName(sName);
  if (!sheet) sheet = ss.insertSheet(sName);

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

function recordSoundForm3Data(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DB_SOUND_FORM3_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DB_SOUND_FORM3_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp", "Calibrated Date", "Pre-Cal Time", "Post-Cal Time", "Serial No.",
      "Model", "Equipment Code", "Standard / Class", "Acoustic Calibrator (Model/SN)",
      "Frequency", "Frequency Weighting", "Time Weighting", "Tolerance (±dB)",
      "Calibrated By", "Approve By", "Result", "SoundData_JSON", "Max Deviation (dB)", "Max Drift (dB)", "Notes"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#0f766e");
  }

  var snClean = String(eq.serial || "").trim();
  var masterMeta = SOUND_EQUIPMENT_MASTER[snClean] || {};
  sheet.appendRow([
    new Date(),
    data.date || "",
    data.preTime || "-",
    data.postTime || "-",
    "'" + snClean,
    eq.model || masterMeta.model || "",
    eq.code || masterMeta.code || "",
    data.standard || masterMeta.standard || "",
    data.calibrator || "-",
    data.freq || "1 kHz",
    data.freqWeighting || "A-Weighting (dBA)",
    data.timeWeighting || "Fast (F)",
    data.toleranceDb || "±0.5 dB",
    data.calibratedBy || "",
    data.approvedBy || "",
    data.result || "ผ่าน",
    JSON.stringify(data.soundData || []),
    data.maxDevDb || "-",
    data.maxDriftDb || "-",
    data.notes || ""
  ]);
}

function recordHeatForm3Data(data, eq) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DB_HEAT_FORM3_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DB_HEAT_FORM3_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp", "Calibrated Date", "Calibrated Time", "Serial No.", "Model",
      "Equipment Code", "Standard", "Ref Device / Module", "Method",
      "Tolerance (±°C)", "Calibrated By", "Approve By", "Result",
      "HeatData_JSON", "Max Error (°C)", "WBGT Check", "Notes"
    ];
    sheet.appendRow(headers);
    styleDbHeader(sheet, headers.length, "#c2410c");
  }

  var snClean = String(eq.serial || "").trim();
  var masterMeta = HEAT_EQUIPMENT_MASTER[snClean] || {};
  sheet.appendRow([
    new Date(),
    data.date || "",
    data.time || "-",
    "'" + snClean,
    eq.model || masterMeta.model || "",
    eq.code || masterMeta.code || "",
    data.standard || masterMeta.standard || "ISO 7243",
    data.refDevice || "-",
    data.method || "-",
    data.toleranceC || "±0.5 °C",
    data.calibratedBy || "",
    data.approvedBy || "",
    data.result || "ผ่าน",
    JSON.stringify(data.heatData || []),
    data.maxErrorC || "-",
    data.wbgtCheck || "-",
    data.notes || ""
  ]);
}

function recordLegacyChecklist(data, eq, category) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LEGACY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.LEGACY_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = [
      "วันเวลาที่บันทึก (Timestamp)", "วันที่ซ่อมบำรุง", "ชื่อช่าง", "รุ่นเครื่อง", "Serial No.",
      "กรองฝุ่น", "ปั๊มอากาศ", "แบตเตอรี่", "ความสะอาด", "เซ็นเซอร์", "หมายเหตุ / รายละเอียด", "หมวดหมู่เครื่องมือ"
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
    data.notes || data.description || "",
    category || "FLUE_GAS"
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
  populateForm1BySerial("01563443", "", "", "FLUE_GAS");
  populateForm2BySerial("01563443", "", "", "FLUE_GAS");
  populateForm3BySerialAndRecord("01563443", "");
}

function initializeSoundSheetsAndForms() {
  populateForm1BySerial("01120943", "", "", "SOUND");
  populateForm2BySerial("01120943", "", "", "SOUND");
  populateSoundForm3BySerialAndRecord("01120943", "");
}

function initializeHeatSheetsAndForms() {
  populateForm1BySerial("TPH 060001", "", "", "HEAT");
  populateForm2BySerial("TPH 060001", "", "", "HEAT");
  populateHeatForm3BySerialAndRecord("TPH 060001", "");
}

function refreshAllFormViews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s1 = ss.getSheetByName(CONFIG.FORM1_TEMPLATE_SHEET);
  if (s1) {
    populateForm1BySerial(String(s1.getRange("I4").getValue()).trim() || "01563443", s1.getRange("I5").getValue(), s1.getRange("K5").getValue(), "FLUE_GAS");
  } else {
    populateForm1BySerial("01563443", "", "", "FLUE_GAS");
  }

  var s2 = ss.getSheetByName(CONFIG.FORM2_TEMPLATE_SHEET);
  if (s2) {
    populateForm2BySerial(String(s2.getRange("E5").getValue()).trim() || "01563443", s2.getRange("B6").getValue(), s2.getRange("E6").getValue(), "FLUE_GAS");
  } else {
    populateForm2BySerial("01563443", "", "", "FLUE_GAS");
  }

  var s3 = ss.getSheetByName(CONFIG.FORM3_TEMPLATE_SHEET);
  if (s3) {
    populateForm3BySerialAndRecord(String(s3.getRange("F3").getValue()).trim() || "01563443", String(s3.getRange("F4").getValue()).trim());
  } else {
    populateForm3BySerialAndRecord("01563443", "");
  }

  var ss1 = ss.getSheetByName(CONFIG.SOUND_FORM1_TEMPLATE_SHEET);
  if (ss1) {
    populateForm1BySerial(String(ss1.getRange("I4").getValue()).trim() || "01120943", ss1.getRange("I5").getValue(), ss1.getRange("K5").getValue(), "SOUND");
  }
  var ss2 = ss.getSheetByName(CONFIG.SOUND_FORM2_TEMPLATE_SHEET);
  if (ss2) {
    populateForm2BySerial(String(ss2.getRange("E5").getValue()).trim() || "01120943", ss2.getRange("B6").getValue(), ss2.getRange("E6").getValue(), "SOUND");
  }
  var ss3 = ss.getSheetByName(CONFIG.SOUND_FORM3_TEMPLATE_SHEET);
  if (ss3) {
    populateSoundForm3BySerialAndRecord(String(ss3.getRange("F3").getValue()).trim() || "01120943", String(ss3.getRange("F4").getValue()).trim());
  }

  var hs1 = ss.getSheetByName(CONFIG.HEAT_FORM1_TEMPLATE_SHEET);
  if (hs1) {
    populateForm1BySerial(String(hs1.getRange("I4").getValue()).trim() || "TPH 060001", hs1.getRange("I5").getValue(), hs1.getRange("K5").getValue(), "HEAT");
  }
  var hs2 = ss.getSheetByName(CONFIG.HEAT_FORM2_TEMPLATE_SHEET);
  if (hs2) {
    populateForm2BySerial(String(hs2.getRange("E5").getValue()).trim() || "TPH 060001", hs2.getRange("B6").getValue(), hs2.getRange("E6").getValue(), "HEAT");
  }
  var hs3 = ss.getSheetByName(CONFIG.HEAT_FORM3_TEMPLATE_SHEET);
  if (hs3) {
    populateHeatForm3BySerialAndRecord(String(hs3.getRange("F3").getValue()).trim() || "TPH 060001", String(hs3.getRange("F4").getValue()).trim());
  }
}

/**
 * 1. แบบฟอร์มที่ 1: บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ (รองรับกรอง S/N และช่วงวันที่)
 */
function populateForm1BySerial(targetSerial, startDateVal, endDateVal, category) {
  var catCfg = resolveCategoryConfig(category);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(catCfg.tpl1);
  if (!sheet) sheet = ss.insertSheet(catCfg.tpl1);

  var cleanSerial = String(targetSerial || catCfg.defaultSerial).replace(/^'/, "").trim();
  var query = queryHistoricalData({
    category: catCfg.category,
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
  sheet.getRange("A2:L2").merge().setValue("แผนกปฏิบัติการภาคสนาม บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด")
       .setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A3:B3").merge().setValue("ชื่อเครื่องมือ").setFontWeight("bold");
  sheet.getRange("C3:F3").merge().setValue(eqInfo.name);
  sheet.getRange("G3:H3").merge().setValue("ยี่ห้อ/รุ่น").setFontWeight("bold");
  sheet.getRange("I3:L3").merge().setValue(eqInfo.model);

  sheet.getRange("A4:B4").merge().setValue("หมายเลขเครื่องมือ").setFontWeight("bold");
  sheet.getRange("C4:F4").merge().setValue(eqInfo.code || "-");
  sheet.getRange("G4:H4").merge().setValue("หมายเลขเครื่อง (S/N)").setFontWeight("bold");
  sheet.getRange("I4").setValue("'" + cleanSerial).setBackground("#fef9c3").setFontWeight("bold");

  var snList1 = catCfg.category === "SOUND"
    ? getSoundSerialList()
    : (catCfg.category === "HEAT" ? getHeatSerialList() : ["01563443", "01807527", "01859560", "60378478", "63455616", "63455658"]);
  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(snList1, true)
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

  // รหัสเอกสารมุมขวาล่างของแบบฟอร์ม
  var docCodeRow1 = 8 + rowsToInsert.length;
  sheet.getRange(docCodeRow1, 10, 1, 3).merge()
       .setValue("FM-LAB-029 / 4 / 17-10-67")
       .setFontSize(9)
       .setHorizontalAlignment("right")
       .setVerticalAlignment("middle");
}

/**
 * 2. แบบฟอร์มที่ 2: บันทึกประวัติเครื่องมือ (รองรับกรอง S/N และช่วงวันที่)
 */
function populateForm2BySerial(targetSerial, startDateVal, endDateVal, category) {
  var catCfg = resolveCategoryConfig(category);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(catCfg.tpl2);
  if (!sheet) sheet = ss.insertSheet(catCfg.tpl2);

  var cleanSerial = String(targetSerial || catCfg.defaultSerial).replace(/^'/, "").trim();
  var query = queryHistoricalData({
    category: catCfg.category,
    targetForm: "FORM_2",
    serial: cleanSerial,
    startDate: startDateVal,
    endDate: endDateVal
  });
  var eqInfo = query.equipment;

  sheet.clear();

  sheet.getRange("A1:F1").merge().setValue("บันทึกประวัติเครื่องมือ")
       .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:F2").merge().setValue("แผนกปฏิบัติการภาคสนาม  บริษัท  อีสเทิร์น ไทย คอนซัลติ้ง  1992  จำกัด")
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

  var snList2 = catCfg.category === "SOUND"
    ? getSoundSerialList()
    : (catCfg.category === "HEAT" ? getHeatSerialList() : ["01563443", "01807527", "01859560", "60378478", "63455616", "63455658"]);
  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(snList2, true)
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

  // รหัสเอกสารมุมขวาล่างของแบบฟอร์ม
  var docCodeRow2 = 8 + rowsToInsert.length;
  sheet.getRange(docCodeRow2, 5, 1, 2).merge()
       .setValue("FM-LAB-026 / 1 / 30-08-56")
       .setFontSize(9)
       .setHorizontalAlignment("right")
       .setVerticalAlignment("middle");
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
    gasData: { O2: [], CO: [], NO: [], NO2: [], SO2: [] }
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
    { key: "O2", title: "O2 Calibrated Data", startRow: 6 },
    { key: "CO", title: "CO Calibrated Data", startRow: 12 },
    { key: "NO", title: "NO Calibrated Data", startRow: 18 },
    { key: "NO2", title: "NO2 Calibrated Data", startRow: 24 },
    { key: "SO2", title: "SO2 Calibrated Data", startRow: 30 }
  ];

  var gasObj = data.gasData || {};

  for (var g = 0; g < gasList.length; g++) {
    var item = gasList[g];
    var rIdx = item.startRow;

    sheet.getRange(rIdx, 1, 1, 4).merge().setValue(item.title)
         .setFontWeight("bold").setBackground("#f1f5f9");

    var unitLabel = item.key === "O2" ? "PPM / %" : "PPM";
    var colHeaders = [
      "Standard Gas Value (" + unitLabel + ")",
      "Flue Gas Analyzer Response (" + unitLabel + ")",
      "Difference (" + unitLabel + ")",
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

  sheet.getRange("A36:B36").merge().setValue("Calibrated By .......................................................")
       .setHorizontalAlignment("center");
  sheet.getRange("C36:D36").merge().setValue("Approve By .......................................................")
       .setHorizontalAlignment("center");

  sheet.getRange("A37:B37").merge().setValue("( " + (data.calibratedBy || "........................................") + " )")
       .setHorizontalAlignment("center");
  sheet.getRange("C37:D37").merge().setValue("( " + (data.approvedBy || "........................................") + " )")
       .setHorizontalAlignment("center");

  // รหัสเอกสารมุมขวาล่างของแบบฟอร์ม
  sheet.getRange("C39:D39").merge().setValue("FM-LAB-265/0/31-10-61")
       .setFontSize(9)
       .setHorizontalAlignment("right");

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 210);
}

/**
 * 3B. แบบฟอร์มที่ 3 (Sound Level Meter): Pre/Post Calibration Check
 * (ตามข้อกำหนด: ยังไม่ต้องใส่รหัส FM-LAB)
 */
function populateSoundForm3BySerialAndRecord(targetSerial, selectedRecordLabel) {
  var cleanSerial = String(targetSerial || "01120943").replace(/^'/, "").trim();
  var query = queryHistoricalData({
    category: "SOUND",
    targetForm: "FORM_3",
    serial: cleanSerial
  });
  var eqInfo = query.equipment;
  var records = query.records;

  var recordOptions = [];
  var chosenRecord = null;

  for (var i = records.length - 1; i >= 0; i--) {
    var r = records[i];
    var label = "#" + r.seq + " | " + r.date + " " + (r.preTime || "-") + " (" + r.result + ")";
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
    preTime: "-",
    postTime: "-",
    calibrator: "-",
    freq: "1 kHz",
    freqWeighting: "A-Weighting (dBA)",
    timeWeighting: "Fast (F)",
    toleranceDb: "±0.5 dB",
    calibratedBy: "........................................",
    approvedBy: "........................................",
    notes: "-",
    soundData: [
      { std: "94.0", pre: "", post: "", preErr: "", postErr: "", drift: "" },
      { std: "114.0", pre: "", post: "", preErr: "", postErr: "", drift: "" }
    ]
  };

  renderSoundForm3Sheet(payload, eqInfo, recordOptions, selectedRecordLabel || "ล่าสุด");
}

function renderSoundForm3Sheet(data, eq, recordOptions, currentLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SOUND_FORM3_TEMPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SOUND_FORM3_TEMPLATE_SHEET);

  var cleanSerial = String(eq.serial || "01120943").replace(/^'/, "").trim();
  sheet.clear();

  sheet.getRange("A1:F1").merge().setValue("Sound Level Meter Verification (Pre / Post Calibration Check)")
       .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:F2").merge().setValue("แผนกปฏิบัติการภาคสนาม บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด")
       .setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A3").setValue("Calibrated Date :").setFontWeight("bold");
  sheet.getRange("B3").setValue(data.date || "-");
  sheet.getRange("C3").setValue("Pre-Cal / Post-Cal Time :").setFontWeight("bold");
  sheet.getRange("D3").setValue((data.preTime || "-") + " / " + (data.postTime || "-"));

  sheet.getRange("A4").setValue("Sound Level Meter Model :").setFontWeight("bold");
  sheet.getRange("B4").setValue((eq.model || "") + (eq.standard ? " [" + eq.standard + "]" : ""));
  sheet.getRange("C4").setValue("Serial No. / Code :").setFontWeight("bold");
  sheet.getRange("D4").setValue("'" + cleanSerial + (eq.code ? " (" + eq.code + ")" : ""));

  sheet.getRange("A5").setValue("Acoustic Calibrator :").setFontWeight("bold");
  sheet.getRange("B5").setValue(data.calibrator || "-");
  sheet.getRange("C5").setValue("Frequency / Tolerance :").setFontWeight("bold");
  sheet.getRange("D5").setValue((data.freq || "1 kHz") + " | " + (data.toleranceDb || "±0.5 dB"));

  sheet.getRange("A6").setValue("Frequency Weighting :").setFontWeight("bold");
  sheet.getRange("B6").setValue(data.freqWeighting || "A-Weighting (dBA)");
  sheet.getRange("C6").setValue("Time Weighting :").setFontWeight("bold");
  sheet.getRange("D6").setValue(data.timeWeighting || "Fast (F)");

  // กล่องเลือก S/N และเลือกรายการย้อนหลังด้านข้าง (G-H หรือ E-F)
  var snRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getSoundSerialList(), true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("E2:F2").merge().setValue("🔍 ค้นหาใบ Verification เครื่องเสียง").setFontWeight("bold").setBackground("#ccfbf1").setHorizontalAlignment("center");
  sheet.getRange("E3").setValue("1. เลือก S/N:").setFontWeight("bold").setFontSize(9);
  sheet.getRange("F3").setValue("'" + cleanSerial).setDataValidation(snRule).setBackground("#fef9c3").setFontWeight("bold");
  sheet.getRange("E4").setValue("2. เลือกครั้งที่ทวนสอบ:").setFontWeight("bold").setFontSize(9);
  sheet.getRange("F4").setValue(currentLabel).setBackground("#fef9c3").setFontSize(9);
  if (recordOptions && recordOptions.length > 0) {
    var recRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(recordOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange("F4").setDataValidation(recRule);
  } else {
    sheet.getRange("F4").clearDataValidations();
  }

  // ตารางตรวจวัด Pre/Post
  var headers = [
    "Standard Sound Level (dB)",
    "Pre-Measurement (dB)",
    "Post-Measurement (dB)",
    "Pre Error (dB)",
    "Post Error (dB)",
    "Drift |Post-Pre| (dB)"
  ];
  sheet.getRange(8, 1, 1, 6).setValues([headers])
       .setBackground("#e2e8f0").setFontWeight("bold").setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  var rows = Array.isArray(data.soundData) ? data.soundData : [];
  var tableRows = [];
  for (var i = 0; i < 2; i++) {
    var pt = rows[i] || {};
    tableRows.push([
      pt.std || (i === 0 ? "94.0" : "114.0"),
      pt.pre || "",
      pt.post || "",
      pt.preErr || "",
      pt.postErr || "",
      pt.drift || ""
    ]);
  }
  sheet.getRange(9, 1, 2, 6).setValues(tableRows).setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  sheet.getRange("A12").setValue("Result :").setFontWeight("bold");
  sheet.getRange("B12:C12").merge().setValue((data.result || "ผ่าน") + " (Max Dev: " + (data.maxDevDb || "-") + ", Drift: " + (data.maxDriftDb || "-") + ")");
  sheet.getRange("D12").setValue("Notes :").setFontWeight("bold");
  sheet.getRange("E12:F12").merge().setValue(data.notes || "-");

  sheet.getRange("A14:C14").merge().setValue("Calibrated By .......................................................").setHorizontalAlignment("center");
  sheet.getRange("D14:F14").merge().setValue("Approve By .......................................................").setHorizontalAlignment("center");
  sheet.getRange("A15:C15").merge().setValue("( " + (data.calibratedBy || "........................................") + " )").setHorizontalAlignment("center");
  sheet.getRange("D15:F15").merge().setValue("( " + (data.approvedBy || "........................................") + " )").setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 175);
  sheet.setColumnWidth(3, 175);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 175);
}

/**
 * แสดงหน้าต่างป๊อปอัปใน Google Sheets เพื่อกดดาวน์โหลดแบบฟอร์มปัจจุบันเป็น PDF หรือ Excel (.xlsx)
 */
/**
 * 3C. แบบฟอร์มที่ 3 (Heat Stress): Verification Module / Reference Thermometer Check (Ta, Tw, Tg)
 * (ตามข้อกำหนด: ยังไม่ต้องใส่รหัส FM-LAB)
 */
function populateHeatForm3BySerialAndRecord(targetSerial, selectedRecordLabel) {
  var cleanSerial = String(targetSerial || "TPH 060001").replace(/^'/, "").trim();
  var query = queryHistoricalData({
    category: "HEAT",
    targetForm: "FORM_3",
    serial: cleanSerial
  });
  var eqInfo = query.equipment;
  var records = query.records;

  var recordOptions = [];
  var chosenRecord = null;
  for (var i = records.length - 1; i >= 0; i--) {
    var r = records[i];
    var label = "#" + r.seq + " | " + r.date + " " + (r.time || "-") + " (" + r.result + ")";
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
    refDevice: "-",
    method: "Verification Module / Reference Thermometer",
    toleranceC: "±0.5 °C",
    maxErrorC: "-",
    wbgtCheck: "-",
    calibratedBy: "........................................",
    approvedBy: "........................................",
    notes: "-",
    heatData: [
      { sensor: "Dry Bulb (Ta)", ref: "", reading: "", err: "", status: "" },
      { sensor: "Wet Bulb (Tw)", ref: "", reading: "", err: "", status: "" },
      { sensor: "Globe Temp (Tg)", ref: "", reading: "", err: "", status: "" }
    ]
  };
  renderHeatForm3Sheet(payload, eqInfo, recordOptions, selectedRecordLabel || "ล่าสุด");
}

function renderHeatForm3Sheet(data, eq, recordOptions, currentLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.HEAT_FORM3_TEMPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.HEAT_FORM3_TEMPLATE_SHEET);

  var cleanSerial = String(eq.serial || "TPH 060001").replace(/^'/, "").trim();
  sheet.clear();

  sheet.getRange("A1:E1").merge().setValue("Heat Stress (WBGT) Verification Form")
       .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:E2").merge().setValue("แผนกปฏิบัติการภาคสนาม บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด")
       .setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A3").setValue("Calibrated Date :").setFontWeight("bold");
  sheet.getRange("B3").setValue(data.date || "-");
  sheet.getRange("C3").setValue("Calibrated Time :").setFontWeight("bold");
  sheet.getRange("D3").setValue(data.time || "-");

  sheet.getRange("A4").setValue("Heat Stress Model :").setFontWeight("bold");
  sheet.getRange("B4").setValue((eq.model || "") + " [" + (eq.standard || "ISO 7243") + "]");
  sheet.getRange("C4").setValue("Serial No. :").setFontWeight("bold");
  sheet.getRange("D4").setValue("'" + cleanSerial);

  sheet.getRange("A5").setValue("Ref / Verification Module :").setFontWeight("bold");
  sheet.getRange("B5").setValue(data.refDevice || "-");
  sheet.getRange("C5").setValue("Tolerance :").setFontWeight("bold");
  sheet.getRange("D5").setValue(data.toleranceC || "±0.5 °C");

  sheet.getRange("A6").setValue("Verification Method :").setFontWeight("bold");
  sheet.getRange("B6:D6").merge().setValue(data.method || "-");

  sheet.getRange("E2:F2").merge().setValue("🔍 ค้นหาใบ Verification เครื่องวัดความร้อน")
       .setFontWeight("bold").setBackground("#ffedd5").setHorizontalAlignment("center");
  sheet.getRange("E3").setValue("1. เลือก S/N:").setFontWeight("bold").setFontSize(9);
  var snRule = SpreadsheetApp.newDataValidation().requireValueInList(getHeatSerialList(), true).setAllowInvalid(true).build();
  sheet.getRange("F3").setValue("'" + cleanSerial).setDataValidation(snRule).setBackground("#fef9c3").setFontWeight("bold");
  sheet.getRange("E4").setValue("2. เลือกครั้งที่ทวนสอบ:").setFontWeight("bold").setFontSize(9);
  sheet.getRange("F4").setValue(currentLabel).setBackground("#fef9c3").setFontSize(9);
  if (recordOptions && recordOptions.length > 0) {
    sheet.getRange("F4").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(recordOptions, true).setAllowInvalid(true).build());
  } else {
    sheet.getRange("F4").clearDataValidations();
  }

  var headers = ["Sensor Parameter", "Reference / Module (°C)", "Meter Reading (°C)", "Difference / Error (°C)", "Status"];
  sheet.getRange(8, 1, 1, 5).setValues([headers])
       .setBackground("#e2e8f0").setFontWeight("bold").setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  var rows = Array.isArray(data.heatData) ? data.heatData : [];
  var tableRows = [];
  for (var i = 0; i < Math.max(rows.length, 3); i++) {
    var pt = rows[i] || {};
    tableRows.push([pt.sensor || "-", pt.ref || "", pt.reading || "", pt.err || "", pt.status || "-"]);
  }
  sheet.getRange(9, 1, tableRows.length, 5).setValues(tableRows).setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  var rSummary = 9 + tableRows.length + 1;
  sheet.getRange(rSummary, 1).setValue("Summary / Result :").setFontWeight("bold");
  sheet.getRange(rSummary, 2, 1, 4).merge().setValue((data.result || "ผ่าน") + " | Max Error: " + (data.maxErrorC || "-") + " | WBGT Check: " + (data.wbgtCheck || "-"));
  sheet.getRange(rSummary + 1, 1).setValue("Notes :").setFontWeight("bold");
  sheet.getRange(rSummary + 1, 2, 1, 4).merge().setValue(data.notes || "-");

  var rSig = rSummary + 3;
  sheet.getRange(rSig, 1, 1, 2).merge().setValue("Calibrated By .......................................................").setHorizontalAlignment("center");
  sheet.getRange(rSig, 3, 1, 3).merge().setValue("Approve By .......................................................").setHorizontalAlignment("center");
  sheet.getRange(rSig + 1, 1, 1, 2).merge().setValue("( " + (data.calibratedBy || "........................................") + " )").setHorizontalAlignment("center");
  sheet.getRange(rSig + 1, 3, 1, 3).merge().setValue("( " + (data.approvedBy || "........................................") + " )").setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 195);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 165);
  sheet.setColumnWidth(4, 165);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 200);
}

function showDownloadDialogForActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ssId = ss.getId();
  var sheetId = sheet.getSheetId();
  var sheetName = sheet.getName();

  var isLandscape = (sheetName === CONFIG.FORM1_TEMPLATE_SHEET || sheetName === CONFIG.SOUND_FORM1_TEMPLATE_SHEET || sheetName === CONFIG.HEAT_FORM1_TEMPLATE_SHEET);
  var exportRange = "";
  if (sheetName === CONFIG.FORM3_TEMPLATE_SHEET || sheetName === CONFIG.SOUND_FORM3_TEMPLATE_SHEET || sheetName === CONFIG.HEAT_FORM3_TEMPLATE_SHEET) {
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

function getEquipmentInfoBySerial(serial, dbSheetName, category) {
  var clean = String(serial || "").replace(/^'/, "").trim();
  var catUpper = String(category || "").toUpperCase();
  var isHeat = catUpper === "HEAT" || Boolean(HEAT_EQUIPMENT_MASTER[clean]);
  var isSound = catUpper === "SOUND" || (!isHeat && Boolean(SOUND_EQUIPMENT_MASTER[clean]));
  var info;

  if (isHeat && HEAT_EQUIPMENT_MASTER[clean]) {
    var hm = HEAT_EQUIPMENT_MASTER[clean];
    info = {
      name: hm.name,
      model: hm.model,
      code: hm.code,
      serial: clean,
      standard: hm.standard,
      status: hm.status
    };
  } else if (isSound && SOUND_EQUIPMENT_MASTER[clean]) {
    var sm = SOUND_EQUIPMENT_MASTER[clean];
    info = {
      name: sm.name,
      model: sm.model,
      code: sm.code,
      serial: clean,
      standard: sm.standard,
      status: sm.status
    };
  } else if (EQUIPMENT_MASTER[clean]) {
    info = {
      name: EQUIPMENT_MASTER[clean].name,
      model: EQUIPMENT_MASTER[clean].model,
      code: EQUIPMENT_MASTER[clean].code,
      serial: clean
    };
  } else {
    info = {
      name: isHeat ? "Heat Stress Monitor" : (isSound ? "Sound Level Meter" : "Flue Gas Analyzer"),
      model: (isHeat || isSound) ? "" : "Testo",
      code: "",
      serial: clean
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(dbSheetName);
  if (dbSheet && dbSheet.getLastRow() > 1) {
    var data = dbSheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var snColIdx = (dbSheetName === CONFIG.DB_HEAT_FORM3_SHEET) ? 3 : 4;
      var rowSn = String(data[i][snColIdx] || "").replace(/^'/, "").trim();
      if (rowSn === clean) {
        if (dbSheetName === CONFIG.DB_FORM3_SHEET) {
          if (data[i][3]) info.model = data[i][3];
          if (data[i][5]) info.code = data[i][5];
        } else if (dbSheetName === CONFIG.DB_SOUND_FORM3_SHEET) {
          if (data[i][5]) info.model = data[i][5];
          if (data[i][6]) info.code = data[i][6];
        } else if (dbSheetName === CONFIG.DB_HEAT_FORM3_SHEET) {
          if (data[i][4]) info.model = data[i][4];
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
// ส่วนรายงาน PDF และอีเมลสรุปประจำสัปดาห์ (Weekly Executive Dashboard Report)
// สรุปครบทั้ง 3 แบบฟอร์ม (สอบเทียบ, ประวัติซ่อม, Gas Verification) + Checklist
// =============================================================================

// Helper อ่านข้อมูล Form 1 ประจำสัปดาห์ตามชื่อชีท
function readWeeklyForm1Records(ss, sheetName, cutoffDate, defaultModel) {
  var records = [];
  var failCount = 0;
  var sheet = ss.getSheetByName(sheetName);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var d = parseDateValue(r[5]);
      if (!d || d < cutoffDate) continue;
      var isFail = String(r[11] || "").indexOf("ไม่ผ่าน") !== -1;
      if (isFail) failCount++;
      records.push({
        date: formatDateDisplay(d),
        model: r[2] || defaultModel,
        serial: String(r[4] || "").replace(/^'/, ""),
        code: r[3] || "-",
        description: r[7] || "-",
        correction: r[8] || "-",
        uncertainty: r[9] || "-",
        tolerance: r[10] || "-",
        result: r[11] || "ผ่าน",
        isFail: isFail,
        agency: r[12] || "-",
        operator: r[13] || "-",
        supervisor: r[14] || "-",
        notes: r[15] || "-"
      });
    }
  }
  return { records: records, fails: failCount };
}

// Helper อ่านข้อมูล Form 2 ประจำสัปดาห์ตามชื่อชีท
function readWeeklyForm2Records(ss, sheetName, cutoffDate, defaultModel) {
  var records = [];
  var sheet = ss.getSheetByName(sheetName);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var j = 1; j < data.length; j++) {
      var r = data[j];
      var d = parseDateValue(r[5]);
      if (!d || d < cutoffDate) continue;
      records.push({
        date: formatDateDisplay(d),
        model: r[2] || defaultModel,
        serial: String(r[4] || "").replace(/^'/, ""),
        code: r[3] || "-",
        damageDetail: r[6] || "-",
        operator: r[7] || "-",
        supervisor: r[8] || "-",
        notes: r[9] || "-"
      });
    }
  }
  return { records: records };
}

function sendWeeklyDashboardReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var cutoffDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - CONFIG.REPORT_LOOKBACK_DAYS, 0, 0, 0);

  // =========================================================================
  // หมวด 1: Flue Gas Analyzer (Testo)
  // =========================================================================
  var gasForm1 = readWeeklyForm1Records(ss, CONFIG.DB_FORM1_SHEET, cutoffDate, "Testo");
  var gasForm2 = readWeeklyForm2Records(ss, CONFIG.DB_FORM2_SHEET, cutoffDate, "Testo");

  var form3Records = [];
  var form3FailCount = 0;
  var maxErrorAll = 0;
  var sheet3 = ss.getSheetByName(CONFIG.DB_FORM3_SHEET);
  if (sheet3 && sheet3.getLastRow() > 1) {
    var data3 = sheet3.getDataRange().getValues();
    for (var k = 1; k < data3.length; k++) {
      var r3 = data3[k];
      var d3 = parseDateValue(r3[1]);
      if (!d3 || d3 < cutoffDate) continue;
      var isFail3 = String(r3[11] || "").indexOf("ไม่ผ่าน") !== -1;
      if (isFail3) form3FailCount++;
      var errNum = parseFloat(String(r3[12] || "").replace("%", ""));
      if (!isNaN(errNum) && Math.abs(errNum) > maxErrorAll) {
        maxErrorAll = Math.abs(errNum);
      }
      var parsedGas = {};
      try { parsedGas = JSON.parse(r3[13] || "{}"); } catch (e) { parsedGas = {}; }
      form3Records.push({
        date: formatDateDisplay(d3),
        time: formatCellTime(r3[2]),
        model: r3[3] || "Testo",
        serial: String(r3[4] || "").replace(/^'/, ""),
        code: r3[5] || "-",
        temp: r3[6] || "-",
        pressure: r3[7] || "-",
        humidity: r3[8] || "-",
        calibratedBy: r3[9] || "-",
        approvedBy: r3[10] || "-",
        result: r3[11] || "ผ่าน",
        isFail: isFail3,
        maxErrorPct: r3[12] || "-",
        gasData: parsedGas
      });
    }
  }

  // =========================================================================
  // หมวด 2: Sound Level Meter (เครื่องวัดระดับเสียง)
  // =========================================================================
  var soundForm1 = readWeeklyForm1Records(ss, CONFIG.DB_SOUND_FORM1_SHEET, cutoffDate, "Sound Level Meter");
  var soundForm2 = readWeeklyForm2Records(ss, CONFIG.DB_SOUND_FORM2_SHEET, cutoffDate, "Sound Level Meter");

  var soundForm3Records = [];
  var soundForm3FailCount = 0;
  var maxSoundDevAll = 0;
  var sheetSound3 = ss.getSheetByName(CONFIG.DB_SOUND_FORM3_SHEET);
  if (sheetSound3 && sheetSound3.getLastRow() > 1) {
    var dataS3 = sheetSound3.getDataRange().getValues();
    for (var ks = 1; ks < dataS3.length; ks++) {
      var rs3 = dataS3[ks];
      var ds3 = parseDateValue(rs3[1]);
      if (!ds3 || ds3 < cutoffDate) continue;
      var isFailS3 = String(rs3[15] || "").indexOf("ไม่ผ่าน") !== -1;
      if (isFailS3) soundForm3FailCount++;
      var devNum = parseFloat(String(rs3[17] || "").replace(/[^0-9.\-]/g, ""));
      if (!isNaN(devNum) && Math.abs(devNum) > maxSoundDevAll) {
        maxSoundDevAll = Math.abs(devNum);
      }
      soundForm3Records.push({
        date: formatDateDisplay(ds3),
        preTime: formatCellTime(rs3[2]),
        postTime: formatCellTime(rs3[3]),
        serial: String(rs3[4] || "").replace(/^'/, ""),
        model: rs3[5] || "Sound Level Meter",
        code: rs3[6] || "-",
        standard: rs3[7] || "",
        calibrator: rs3[8] || "-",
        freq: rs3[9] || "1 kHz",
        freqWeighting: rs3[10] || "A",
        timeWeighting: rs3[11] || "Fast",
        toleranceDb: rs3[12] || "±0.5 dB",
        calibratedBy: rs3[13] || "-",
        approvedBy: rs3[14] || "-",
        result: rs3[15] || "ผ่าน",
        isFail: isFailS3,
        maxDevDb: rs3[17] || "-",
        maxDriftDb: rs3[18] || "-",
        notes: rs3[19] || "-"
      });
    }
  }

  // =========================================================================
  // หมวด 3: Heat Stress Monitor (เครื่องวัดความร้อน WBGT)
  // =========================================================================
  var heatForm1 = readWeeklyForm1Records(ss, CONFIG.DB_HEAT_FORM1_SHEET, cutoffDate, "Heat Stress Monitor");
  var heatForm2 = readWeeklyForm2Records(ss, CONFIG.DB_HEAT_FORM2_SHEET, cutoffDate, "Heat Stress Monitor");

  var heatForm3Records = [];
  var heatForm3FailCount = 0;
  var maxHeatErrorAll = 0;
  var sheetHeat3 = ss.getSheetByName(CONFIG.DB_HEAT_FORM3_SHEET);
  if (sheetHeat3 && sheetHeat3.getLastRow() > 1) {
    var dataH3 = sheetHeat3.getDataRange().getValues();
    for (var kh = 1; kh < dataH3.length; kh++) {
      var rh3 = dataH3[kh];
      var dh3 = parseDateValue(rh3[1]);
      if (!dh3 || dh3 < cutoffDate) continue;
      var isFailH3 = String(rh3[12] || "").indexOf("ไม่ผ่าน") !== -1;
      if (isFailH3) heatForm3FailCount++;
      var errHNum = parseFloat(String(rh3[14] || "").replace(/[^0-9.\-]/g, ""));
      if (!isNaN(errHNum) && Math.abs(errHNum) > maxHeatErrorAll) {
        maxHeatErrorAll = Math.abs(errHNum);
      }
      heatForm3Records.push({
        date: formatDateDisplay(dh3),
        time: formatCellTime(rh3[2]),
        serial: String(rh3[3] || "").replace(/^'/, ""),
        model: rh3[4] || "Heat Stress Monitor",
        code: rh3[5] || "-",
        standard: rh3[6] || "ISO 7243",
        refDevice: rh3[7] || "-",
        method: rh3[8] || "-",
        toleranceC: rh3[9] || "±0.5 °C",
        calibratedBy: rh3[10] || "-",
        approvedBy: rh3[11] || "-",
        result: rh3[12] || "ผ่าน",
        isFail: isFailH3,
        maxErrorC: rh3[14] || "-",
        wbgtCheck: rh3[15] || "-",
        notes: rh3[16] || "-"
      });
    }
  }

  // =========================================================================
  // Checklist 5 รายการ (แยก Flue Gas vs Sound)
  // =========================================================================
  var checklistGas = [];
  var checklistSound = [];
  var checklistHeat = [];
  var checklistIssueGas = 0;
  var checklistIssueSound = 0;
  var checklistIssueHeat = 0;
  var sheetLegacy = ss.getSheetByName(CONFIG.LEGACY_SHEET);
  if (sheetLegacy && sheetLegacy.getLastRow() > 1) {
    var dataLegacy = sheetLegacy.getDataRange().getValues();
    for (var m = 1; m < dataLegacy.length; m++) {
      var rL = dataLegacy[m];
      var dL = parseDateValue(rL[1]);
      if (!dL || dL < cutoffDate) continue;
      var hasIssueL = [rL[5], rL[6], rL[7], rL[8], rL[9]].some(function(val) {
        var str = String(val);
        return str.indexOf("ต้องแก้ไข") !== -1 || str.indexOf("เสื่อม") !== -1 || str.indexOf("ผิดปกติ") !== -1;
      });
      var snL = String(rL[4] || "-").replace(/^'/, "").trim();
      var catL = String(rL[11] || "").toUpperCase();
      var isHeatChk = catL === "HEAT" || Boolean(HEAT_EQUIPMENT_MASTER[snL]);
      var isSoundChk = !isHeatChk && (catL === "SOUND" || Boolean(SOUND_EQUIPMENT_MASTER[snL]));
      var chkItem = {
        date: formatDateDisplay(dL),
        technician: rL[2] || "-",
        model: rL[3] || "-",
        serial: snL,
        filter: rL[5] || "-",
        pump: rL[6] || "-",
        battery: rL[7] || "-",
        clean: rL[8] || "-",
        sensor: rL[9] || "-",
        notes: rL[10] || "-",
        hasIssue: hasIssueL
      };
      if (isHeatChk) {
        if (hasIssueL) checklistIssueHeat++;
        checklistHeat.push(chkItem);
      } else if (isSoundChk) {
        if (hasIssueL) checklistIssueSound++;
        checklistSound.push(chkItem);
      } else {
        if (hasIssueL) checklistIssueGas++;
        checklistGas.push(chkItem);
      }
    }
  }

  // รวมสถิติทั้ง 3 หมวดเครื่องมือ
  var totalGasIssues = gasForm1.fails + gasForm2.records.length + form3FailCount + checklistIssueGas;
  var totalSoundIssues = soundForm1.fails + soundForm2.records.length + soundForm3FailCount + checklistIssueSound;
  var totalHeatIssues = heatForm1.fails + heatForm2.records.length + heatForm3FailCount + checklistIssueHeat;
  var totalIssues = totalGasIssues + totalSoundIssues + totalHeatIssues;

  var totalGasActivities = gasForm1.records.length + gasForm2.records.length + form3Records.length + checklistGas.length;
  var totalSoundActivities = soundForm1.records.length + soundForm2.records.length + soundForm3Records.length + checklistSound.length;
  var totalHeatActivities = heatForm1.records.length + heatForm2.records.length + heatForm3Records.length + checklistHeat.length;
  var totalActivities = totalGasActivities + totalSoundActivities + totalHeatActivities;

  var summaryData = {
    cutoffDate: cutoffDate,
    today: today,
    totalIssues: totalIssues,
    totalActivities: totalActivities,
    form1: gasForm1,
    form2: gasForm2,
    form3: { records: form3Records, fails: form3FailCount, maxError: maxErrorAll },
    checklist: { records: checklistGas, issues: checklistIssueGas },
    soundForm1: soundForm1,
    soundForm2: soundForm2,
    soundForm3: { records: soundForm3Records, fails: soundForm3FailCount, maxDevDb: maxSoundDevAll },
    soundChecklist: { records: checklistSound, issues: checklistIssueSound },
    heatForm1: heatForm1,
    heatForm2: heatForm2,
    heatForm3: { records: heatForm3Records, fails: heatForm3FailCount, maxErrorC: maxHeatErrorAll },
    heatChecklist: { records: checklistHeat, issues: checklistIssueHeat }
  };

  // สร้างเนื้อหา PDF และแปลงเป็น Blob
  var htmlPdfContent = buildWeeklyHtmlReport(summaryData);
  var pdfBlob = Utilities.newBlob(htmlPdfContent, MimeType.HTML)
    .setName("Field_Equipment_Weekly_Report_" + Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd") + ".pdf")
    .getAs(MimeType.PDF);

  // หัวข้ออีเมล
  var emailSubject = "รายงานสรุปงานบำรุงรักษาและทวนสอบเครื่อง Flue Gas Analyzer, เครื่องวัดเสียง และเครื่องวัดความร้อน ประจำสัปดาห์ (ระหว่างวันที่ " +
    formatDateDisplay(cutoffDate) + " - " + formatDateDisplay(today) + ")";

  // เนื้อหาอีเมล HTML
  var emailBody = buildWeeklyEmailHtmlBody(summaryData, ss.getUrl());

  // ส่งอีเมล
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: emailSubject,
    htmlBody: emailBody,
    attachments: [pdfBlob]
  });

  try {
    ss.toast("ส่งอีเมลรายงานสรุป (Flue Gas + เครื่องวัดเสียง) ไปยัง " + CONFIG.EMAIL_RECIPIENTS + " เรียบร้อยแล้ว", "ส่งอีเมลสำเร็จ", 5);
  } catch (err) {}
}

/**
 * สร้างเนื้อหา HTML สำหรับแสดงในกล่องข้อความอีเมล (รวม Flue Gas Analyzer + Sound Level Meter)
 */
function buildWeeklyEmailHtmlBody(s, sheetUrl) {
  var isOk = s.totalIssues === 0;
  var statusBadge = isOk
    ? "<span style='background:#dcfce7; color:#15803d; padding:6px 12px; border-radius:999px; font-weight:bold; font-size:13px;'>✓ อุปกรณ์ทุกเครื่องพร้อมใช้งานปกติ</span>"
    : "<span style='background:#fee2e2; color:#b91c1c; padding:6px 12px; border-radius:999px; font-weight:bold; font-size:13px;'>⚠️ พบรายการที่ต้องติดตาม " + s.totalIssues + " รายการ</span>";

  function buildForm1EmailRows(records, emptyMsg) {
    if (!records || records.length === 0) {
      return "<tr><td colspan='5' style='text-align:center; padding:10px; color:#6b7280; font-size:12px;'>" + emptyMsg + "</td></tr>";
    }
    var out = "";
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var resColor = r.isFail ? "#dc2626" : "#16a34a";
      out += "<tr style='border-bottom:1px solid #e5e7eb; font-size:12px;'>" +
        "<td style='padding:8px 6px; text-align:center;'>" + r.date + "</td>" +
        "<td style='padding:8px 6px;'><b>" + escapeXml(r.model) + "</b><br><small style='color:#6b7280;'>S/N: " + escapeXml(r.serial) + "</small></td>" +
        "<td style='padding:8px 6px;'>" + escapeXml(r.description) + "</td>" +
        "<td style='padding:8px 6px; text-align:center; font-weight:bold; color:" + resColor + ";'>" + escapeXml(r.result) + "</td>" +
        "<td style='padding:8px 6px; text-align:center;'>" + escapeXml(r.operator) + "</td>" +
      "</tr>";
    }
    return out;
  }

  function buildForm2EmailRows(records, emptyMsg) {
    if (!records || records.length === 0) {
      return "<tr><td colspan='4' style='text-align:center; padding:10px; color:#15803d; font-size:12px;'>" + emptyMsg + "</td></tr>";
    }
    var out = "";
    for (var j = 0; j < records.length; j++) {
      var r = records[j];
      out += "<tr style='border-bottom:1px solid #fed7aa; background-color:#fff7ed; font-size:12px;'>" +
        "<td style='padding:8px 6px; text-align:center; font-weight:bold; color:#9a3412;'>" + r.date + "</td>" +
        "<td style='padding:8px 6px;'><b>" + escapeXml(r.model) + "</b><br><small style='color:#6b7280;'>S/N: " + escapeXml(r.serial) + "</small></td>" +
        "<td style='padding:8px 6px; color:#c2410c;'><b>" + escapeXml(r.damageDetail) + "</b></td>" +
        "<td style='padding:8px 6px; text-align:center;'>" + escapeXml(r.operator) + "</td>" +
      "</tr>";
    }
    return out;
  }

  // ตารางย่อ Flue Gas Analyzer
  var f1Rows = buildForm1EmailRows(s.form1.records, "ไม่มีรายการสอบเทียบ/บำรุงรักษา Flue Gas ในรอบสัปดาห์นี้");
  var f2Rows = buildForm2EmailRows(s.form2.records, "✓ ไม่พบประวัติความเสียหายหรือส่งซ่อม Flue Gas ในรอบสัปดาห์นี้");

  var f3Rows = "";
  if (s.form3.records.length === 0) {
    f3Rows = "<tr><td colspan='5' style='text-align:center; padding:10px; color:#6b7280; font-size:12px;'>ไม่มีรายการทวนสอบก๊าซในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var k = 0; k < s.form3.records.length; k++) {
      var r3 = s.form3.records[k];
      var res3Color = r3.isFail ? "#dc2626" : "#16a34a";
      f3Rows += "<tr style='border-bottom:1px solid #e5e7eb; font-size:12px;'>" +
        "<td style='padding:8px 6px; text-align:center;'>" + r3.date + "<br><small style='color:#6b7280;'>" + r3.time + "</small></td>" +
        "<td style='padding:8px 6px;'><b>" + escapeXml(r3.model) + "</b><br><small style='color:#6b7280;'>S/N: " + escapeXml(r3.serial) + "</small></td>" +
        "<td style='padding:8px 6px; text-align:center; font-family:monospace;'>" + escapeXml(r3.maxErrorPct) + "</td>" +
        "<td style='padding:8px 6px; text-align:center; font-weight:bold; color:" + res3Color + ";'>" + escapeXml(r3.result) + "</td>" +
        "<td style='padding:8px 6px; text-align:center;'>" + escapeXml(r3.calibratedBy) + "</td>" +
      "</tr>";
    }
  }

  // ตารางย่อ Sound Level Meter
  var sf1Rows = buildForm1EmailRows(s.soundForm1.records, "ไม่มีรายการสอบเทียบ/บำรุงรักษาเครื่องวัดเสียงในรอบสัปดาห์นี้");
  var sf2Rows = buildForm2EmailRows(s.soundForm2.records, "✓ ไม่พบประวัติความเสียหายหรือส่งซ่อมเครื่องวัดเสียงในรอบสัปดาห์นี้");

  var hf1Rows = buildForm1EmailRows(s.heatForm1.records, "ไม่มีรายการสอบเทียบ/บำรุงรักษาเครื่องวัดความร้อนในรอบสัปดาห์นี้");
  var hf2Rows = buildForm2EmailRows(s.heatForm2.records, "✓ ไม่พบประวัติความเสียหายหรือส่งซ่อมเครื่องวัดความร้อนในรอบสัปดาห์นี้");

  var hf3Rows = "";
  if (s.heatForm3.records.length === 0) {
    hf3Rows = "<tr><td colspan='5' style='text-align:center; padding:10px; color:#6b7280; font-size:12px;'>ไม่มีรายการทวนสอบเครื่องวัดความร้อนในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var kh = 0; kh < s.heatForm3.records.length; kh++) {
      var rh3 = s.heatForm3.records[kh];
      var resH3Color = rh3.isFail ? "#dc2626" : "#16a34a";
      hf3Rows += "<tr style='border-bottom:1px solid #e5e7eb; font-size:12px;'>" +
        "<td style='padding:8px 6px; text-align:center;'>" + rh3.date + "<br><small style='color:#6b7280;'>" + rh3.time + "</small></td>" +
        "<td style='padding:8px 6px;'><b>" + escapeXml(rh3.model) + "</b><br><small style='color:#6b7280;'>S/N: " + escapeXml(rh3.serial) + "</small></td>" +
        "<td style='padding:8px 6px; text-align:center; font-family:monospace;'>" + escapeXml(rh3.maxErrorC) + "</td>" +
        "<td style='padding:8px 6px; text-align:center; font-weight:bold; color:" + resH3Color + ";'>" + escapeXml(rh3.result) + "</td>" +
        "<td style='padding:8px 6px; text-align:center;'>" + escapeXml(rh3.calibratedBy) + "</td>" +
      "</tr>";
    }
  }

  var sf3Rows = "";
  if (s.soundForm3.records.length === 0) {
    sf3Rows = "<tr><td colspan='5' style='text-align:center; padding:10px; color:#6b7280; font-size:12px;'>ไม่มีรายการทวนสอบ Pre/Post เครื่องวัดเสียงในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var ks = 0; ks < s.soundForm3.records.length; ks++) {
      var rs3 = s.soundForm3.records[ks];
      var resS3Color = rs3.isFail ? "#dc2626" : "#16a34a";
      sf3Rows += "<tr style='border-bottom:1px solid #e5e7eb; font-size:12px;'>" +
        "<td style='padding:8px 6px; text-align:center;'>" + rs3.date + "<br><small style='color:#6b7280;'>" + rs3.preTime + " / " + rs3.postTime + "</small></td>" +
        "<td style='padding:8px 6px;'><b>" + escapeXml(rs3.model) + "</b><br><small style='color:#6b7280;'>S/N: " + escapeXml(rs3.serial) + "</small></td>" +
        "<td style='padding:8px 6px; text-align:center; font-family:monospace;'>" + escapeXml(rs3.maxDevDb) + "<br><small style='color:#6b7280;'>Drift: " + escapeXml(rs3.maxDriftDb) + "</small></td>" +
        "<td style='padding:8px 6px; text-align:center; font-weight:bold; color:" + resS3Color + ";'>" + escapeXml(rs3.result) + "</td>" +
        "<td style='padding:8px 6px; text-align:center;'>" + escapeXml(rs3.calibratedBy) + "</td>" +
      "</tr>";
    }
  }

  var html = "<div style='font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif; max-width:680px; margin:0 auto; background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; overflow:hidden;'>" +
    "<div style='background:linear-gradient(135deg, #0f172a 0%, #1e40af 100%); color:#ffffff; padding:24px 20px; text-align:center;'>" +
      "<div style='font-size:12px; letter-spacing:1px; text-transform:uppercase; opacity:0.9; margin-bottom:4px;'>แผนกปฏิบัติการภาคสนาม | บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด</div>" +
      "<h2 style='margin:0 0 8px 0; font-size:19px; font-weight:bold;'>รายงานสรุปงานบำรุงรักษาและทวนสอบเครื่องมือภาคสนาม</h2>" +
      "<div style='font-size:13px; opacity:0.95;'>Flue Gas • Sound Level • Heat Stress | รอบสัปดาห์: " + formatDateDisplay(s.cutoffDate) + " ถึง " + formatDateDisplay(s.today) + "</div>" +
      "<div style='margin-top:14px;'>" + statusBadge + "</div>" +
    "</div>" +

    "<div style='padding:20px;'>" +
      "<p style='margin:0 0 16px 0; font-size:14px; color:#374151; line-height:1.7;'>" +
        "เรียนทุกท่าน<br>" +
        "สรุปผลการดำเนินงานเกี่ยวกับการบำรุงรักษาเครื่อง Flue Gas Analyzer, เครื่องวัดระดับเสียง (Sound Level Meter) และเครื่องวัดความร้อน (Heat Stress Monitor) ในรอบ 7 วันที่ผ่านมา ดังนี้ครับ<br>" +
        "(รายละเอียดเพิ่มเติมดังเอกสารแนบ)<br>" +
        "ขอบคุณครับ" +
      "</p>" +

      // การ์ดสรุป 2 หมวด (KPI Summary)
      "<table style='width:100%; border-collapse:separate; border-spacing:8px; margin-bottom:16px;'>" +
        "<tr>" +
          "<td style='width:50%; background:#eff6ff; border:1px solid #bfdbfe; border-radius:12px; padding:12px;'>" +
            "<div style='font-size:12px; font-weight:bold; color:#1e40af;'>🔥 หมวดที่ 1: Flue Gas Analyzer (Testo)</div>" +
            "<div style='font-size:12px; color:#1e293b; margin-top:6px; line-height:1.6;'>" +
              "• สอบเทียบ/ตรวจสภาพ (F1): <b>" + s.form1.records.length + "</b> (ไม่ผ่าน " + s.form1.fails + ")<br>" +
              "• แจ้งซ่อม/เสียหาย (F2): <b>" + s.form2.records.length + "</b><br>" +
              "• ทวนสอบก๊าซ (F3): <b>" + s.form3.records.length + "</b> (Max Err: " + (s.form3.maxError > 0 ? s.form3.maxError.toFixed(2) + "%" : "-") + ")<br>" +
              "• Checklist 5 ระบบ: <b>" + s.checklist.records.length + "</b> (ต้องแก้ไข " + s.checklist.issues + ")" +
            "</div>" +
          "</td>" +
          "<td style='width:50%; background:#f0fdfa; border:1px solid #99f6e4; border-radius:12px; padding:12px;'>" +
            "<div style='font-size:12px; font-weight:bold; color:#0f766e;'>🔊 หมวดที่ 2: Sound Level Meter (เครื่องวัดเสียง)</div>" +
            "<div style='font-size:12px; color:#1e293b; margin-top:6px; line-height:1.6;'>" +
              "• สอบเทียบ/ตรวจสภาพ (F1): <b>" + s.soundForm1.records.length + "</b> (ไม่ผ่าน " + s.soundForm1.fails + ")<br>" +
              "• แจ้งซ่อม/เสียหาย (F2): <b>" + s.soundForm2.records.length + "</b><br>" +
              "• Pre/Post Check (F3): <b>" + s.soundForm3.records.length + "</b> (Max Dev: " + (s.soundForm3.maxDevDb > 0 ? s.soundForm3.maxDevDb.toFixed(2) + " dB" : "-") + ")<br>" +
              "• Checklist 5 ระบบ: <b>" + s.soundChecklist.records.length + "</b> (ต้องแก้ไข " + s.soundChecklist.issues + ")" +
            "</div>" +
          "</td>" +
        "</tr>" +
        "<tr>" +
          "<td colspan='2' style='background:#fff7ed; border:1px solid #fed7aa; border-radius:12px; padding:12px;'>" +
            "<div style='font-size:12px; font-weight:bold; color:#c2410c;'>🌡️ หมวดที่ 3: Heat Stress Monitor (เครื่องวัดความร้อน WBGT - ISO 7243)</div>" +
            "<div style='font-size:12px; color:#1e293b; margin-top:4px;'>" +
              "สอบเทียบ/ตรวจสภาพ (F1): <b>" + s.heatForm1.records.length + "</b> (ไม่ผ่าน " + s.heatForm1.fails + ") &nbsp;|&nbsp; " +
              "แจ้งซ่อม (F2): <b>" + s.heatForm2.records.length + "</b> &nbsp;|&nbsp; " +
              "ทวนสอบ Ta/Tw/Tg (F3): <b>" + s.heatForm3.records.length + "</b> (Max Err: " + (s.heatForm3.maxErrorC > 0 ? s.heatForm3.maxErrorC.toFixed(2) + " °C" : "-") + ") &nbsp;|&nbsp; " +
              "Checklist: <b>" + s.heatChecklist.records.length + "</b> (แก้ไข " + s.heatChecklist.issues + ")" +
            "</div>" +
          "</td>" +
        "</tr>" +
      "</table>" +

      // --- ส่วนที่ 1: Flue Gas Analyzer ---
      "<div style='font-size:14px; font-weight:bold; color:#1e3a8a; margin:18px 0 8px 0; border-bottom:2px solid #dbeafe; padding-bottom:4px;'>🔥 หมวด Flue Gas Analyzer (Testo)</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#f8fafc; padding:8px 12px; font-weight:bold; font-size:12px; color:#1e293b; border-bottom:1px solid #e5e7eb;'>📋 แบบฟอร์มที่ 1: การสอบเทียบและบำรุงรักษา</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายการ</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ตรวจ</th></tr></thead><tbody>" + f1Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid " + (s.form2.records.length > 0 ? "#fed7aa" : "#e5e7eb") + "; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:" + (s.form2.records.length > 0 ? "#ffedd5" : "#f8fafc") + "; padding:8px 12px; font-weight:bold; font-size:12px; color:" + (s.form2.records.length > 0 ? "#9a3412" : "#1e293b") + "; border-bottom:1px solid #e5e7eb;'>🛠️ แบบฟอร์มที่ 2: ประวัติความเสียหาย / ซ่อมแซม</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายละเอียด</th><th style='padding:6px; text-align:center;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + f2Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#f8fafc; padding:8px 12px; font-weight:bold; font-size:12px; color:#1e293b; border-bottom:1px solid #e5e7eb;'>🧪 แบบฟอร์มที่ 3: Gas Verification (O2, CO, NO, NO2, SO2)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วัน/เวลา</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px; text-align:center;'>Max Error</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ทวนสอบ</th></tr></thead><tbody>" + f3Rows + "</tbody></table>" +
      "</div>" +

      // --- ส่วนที่ 2: Sound Level Meter ---
      "<div style='font-size:14px; font-weight:bold; color:#0f766e; margin:22px 0 8px 0; border-bottom:2px solid #ccfbf1; padding-bottom:4px;'>🔊 หมวด Sound Level Meter (เครื่องวัดระดับเสียง)</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#f0fdfa; padding:8px 12px; font-weight:bold; font-size:12px; color:#115e59; border-bottom:1px solid #e5e7eb;'>📋 แบบฟอร์มที่ 1: การสอบเทียบและบำรุงรักษา (เครื่องวัดเสียง)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายการ</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ตรวจ</th></tr></thead><tbody>" + sf1Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid " + (s.soundForm2.records.length > 0 ? "#fed7aa" : "#e5e7eb") + "; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:" + (s.soundForm2.records.length > 0 ? "#ffedd5" : "#f0fdfa") + "; padding:8px 12px; font-weight:bold; font-size:12px; color:" + (s.soundForm2.records.length > 0 ? "#9a3412" : "#115e59") + "; border-bottom:1px solid #e5e7eb;'>🛠️ แบบฟอร์มที่ 2: ประวัติความเสียหาย / ซ่อมแซม (เครื่องวัดเสียง)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายละเอียด</th><th style='padding:6px; text-align:center;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + sf2Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#f0fdfa; padding:8px 12px; font-weight:bold; font-size:12px; color:#115e59; border-bottom:1px solid #e5e7eb;'>🎙️ แบบฟอร์มที่ 3: Pre/Post Calibration Check (94 dB / 114 dB)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วัน/เวลา</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px; text-align:center;'>Max Dev / Drift</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ทวนสอบ</th></tr></thead><tbody>" + sf3Rows + "</tbody></table>" +
      "</div>" +

      // --- ส่วนที่ 3: Heat Stress Monitor ---
      "<div style='font-size:14px; font-weight:bold; color:#c2410c; margin:22px 0 8px 0; border-bottom:2px solid #fed7aa; padding-bottom:4px;'>🌡️ หมวด Heat Stress Monitor (เครื่องวัดความร้อน WBGT)</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#fff7ed; padding:8px 12px; font-weight:bold; font-size:12px; color:#9a3412; border-bottom:1px solid #e5e7eb;'>📋 แบบฟอร์มที่ 1: การสอบเทียบและบำรุงรักษา (เครื่องวัดความร้อน)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายการ</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ตรวจ</th></tr></thead><tbody>" + hf1Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid " + (s.heatForm2.records.length > 0 ? "#fed7aa" : "#e5e7eb") + "; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:" + (s.heatForm2.records.length > 0 ? "#ffedd5" : "#fff7ed") + "; padding:8px 12px; font-weight:bold; font-size:12px; color:#9a3412; border-bottom:1px solid #e5e7eb;'>🛠️ แบบฟอร์มที่ 2: ประวัติความเสียหาย / ซ่อมแซม (เครื่องวัดความร้อน)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วันที่</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px;'>รายละเอียด</th><th style='padding:6px; text-align:center;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + hf2Rows + "</tbody></table>" +
      "</div>" +
      "<div style='margin-top:10px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;'>" +
        "<div style='background:#fff7ed; padding:8px 12px; font-weight:bold; font-size:12px; color:#9a3412; border-bottom:1px solid #e5e7eb;'>🌡️ แบบฟอร์มที่ 3: Heat Stress Verification (Ta, Tw, Tg ±0.5 °C)</div>" +
        "<table style='width:100%; border-collapse:collapse;'><thead><tr style='background:#f1f5f9; font-size:11px; color:#475569;'><th style='padding:6px; text-align:center;'>วัน/เวลา</th><th style='padding:6px;'>รุ่น / S/N</th><th style='padding:6px; text-align:center;'>Max Error</th><th style='padding:6px; text-align:center;'>ผล</th><th style='padding:6px; text-align:center;'>ผู้ทวนสอบ</th></tr></thead><tbody>" + hf3Rows + "</tbody></table>" +
      "</div>" +

      "<div style='margin-top:24px; text-align:center;'>" +
        "<a href='" + sheetUrl + "' style='display:inline-block; background:#1e40af; color:#ffffff; text-decoration:none; padding:10px 20px; border-radius:8px; font-size:13px; font-weight:bold; margin-right:8px;'>📊 เปิดดู Google Sheets ฐานข้อมูล</a>" +
        "<a href='" + CONFIG.WEBAPP_URL + "' style='display:inline-block; background:#0f172a; color:#ffffff; text-decoration:none; padding:10px 20px; border-radius:8px; font-size:13px; font-weight:bold;'>🌐 เปิดหน้าเว็บแอปบันทึกข้อมูล</a>" +
      "</div>" +

      "<div style='margin-top:20px; padding:12px; background:#f8fafc; border-radius:8px; border:1px solid #e2e8f0; font-size:12px; color:#64748b; text-align:center;'>" +
        "📎 เอกสารรายงานสรุปฉบับเต็มครอบคลุมทั้ง <b>Flue Gas Analyzer</b> และ <b>Sound Level Meter</b> ถูกแนบมาเป็นไฟล์ PDF ท้ายอีเมลฉบับนี้ครับ" +
      "</div>" +
    "</div>" +

    "<div style='background:#f1f5f9; padding:14px 20px; text-align:center; font-size:11px; color:#64748b; border-top:1px solid #e2e8f0;'>" +
      "ระบบรายงานผลอัตโนมัติ | แผนกปฏิบัติการภาคสนาม บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด<br>" +
      "อีเมลนี้ส่งไปยัง " + escapeXml(CONFIG.EMAIL_RECIPIENTS) + " ตามกำหนดการรายงานสรุปประจำสัปดาห์" +
    "</div>" +
  "</div>";

  return html;
}

/**
 * สร้างเอกสาร PDF สรุปประจำสัปดาห์ (รวมทั้ง Flue Gas Analyzer และ Sound Level Meter)
 */
function buildWeeklyHtmlReport(s) {
  function renderBadgePdf(val) {
    if (!val || val === "-") return "<span style='color:#666;'>-</span>";
    var str = String(val);
    if (str.indexOf("ต้องแก้ไข") !== -1 || str.indexOf("เสื่อม") !== -1 || str.indexOf("ผิดปกติ") !== -1 || str.indexOf("ไม่ผ่าน") !== -1) {
      return "<span style='background:#fee2e2; color:#b91c1c; padding:2px 5px; border-radius:3px; font-weight:bold; font-size:9px;'>ต้องแก้ไข</span>";
    }
    return "<span style='background:#dcfce7; color:#15803d; padding:2px 5px; border-radius:3px; font-size:9px;'>ปกติ</span>";
  }

  function buildForm1PdfRows(records) {
    if (!records || records.length === 0) return "<tr><td colspan='7' style='text-align:center; padding:8px; color:#6b7280;'>ไม่มีรายการในรอบสัปดาห์นี้</td></tr>";
    var out = "";
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var bg = r.isFail ? "#fee2e2" : (i % 2 === 0 ? "#ffffff" : "#f9fafb");
      out += "<tr style='background-color:" + bg + ";'>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + r.date + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px;'><b>" + escapeXml(r.model) + "</b> (" + escapeXml(r.serial) + ")</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px;'>" + escapeXml(r.description) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + escapeXml(r.correction) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + escapeXml(r.tolerance) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-weight:bold; font-size:9px; color:" + (r.isFail ? "#b91c1c" : "#15803d") + ";'>" + escapeXml(r.result) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(r.operator) + "</td>" +
      "</tr>";
    }
    return out;
  }

  function buildForm2PdfRows(records) {
    if (!records || records.length === 0) return "<tr><td colspan='5' style='text-align:center; padding:8px; color:#15803d;'>✓ ไม่พบประวัติความเสียหายหรือส่งซ่อมในรอบสัปดาห์นี้ (ปกติ)</td></tr>";
    var out = "";
    for (var j = 0; j < records.length; j++) {
      var r = records[j];
      out += "<tr style='background-color:#fff7ed;'>" +
        "<td style='padding:5px; border:1px solid #fed7aa; text-align:center; font-size:9px; font-weight:bold; color:#9a3412;'>" + r.date + "</td>" +
        "<td style='padding:5px; border:1px solid #fed7aa; font-size:9px;'><b>" + escapeXml(r.model) + "</b> (" + escapeXml(r.serial) + ")</td>" +
        "<td style='padding:5px; border:1px solid #fed7aa; font-size:9px; color:#9a3412; font-weight:bold;'>" + escapeXml(r.damageDetail) + "</td>" +
        "<td style='padding:5px; border:1px solid #fed7aa; font-size:9px; text-align:center;'>" + escapeXml(r.operator) + "</td>" +
        "<td style='padding:5px; border:1px solid #fed7aa; font-size:9px;'>" + escapeXml(r.notes) + "</td>" +
      "</tr>";
    }
    return out;
  }

  function buildChecklistPdfRows(records) {
    if (!records || records.length === 0) return "<tr><td colspan='9' style='text-align:center; padding:8px; color:#6b7280;'>ไม่มีรายการตรวจสภาพในรอบสัปดาห์นี้</td></tr>";
    var out = "";
    for (var m = 0; m < records.length; m++) {
      var rL = records[m];
      var bgL = rL.hasIssue ? "#fff7ed" : (m % 2 === 0 ? "#ffffff" : "#f9fafb");
      out += "<tr style='background-color:" + bgL + ";'>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + rL.date + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; font-size:9px;'>" + escapeXml(rL.technician) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; font-size:9px;'>" + escapeXml(rL.model) + " (" + escapeXml(rL.serial) + ")</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center;'>" + renderBadgePdf(rL.filter) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center;'>" + renderBadgePdf(rL.pump) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center;'>" + renderBadgePdf(rL.battery) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center;'>" + renderBadgePdf(rL.clean) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; text-align:center;'>" + renderBadgePdf(rL.sensor) + "</td>" +
        "<td style='padding:4px; border:1px solid #cbd5e1; font-size:8px;'>" + escapeXml(rL.notes) + "</td>" +
      "</tr>";
    }
    return out;
  }

  var f1RowsPdf = buildForm1PdfRows(s.form1.records);
  var f2RowsPdf = buildForm2PdfRows(s.form2.records);
  var chkRowsPdf = buildChecklistPdfRows(s.checklist.records);

  var f3RowsPdf = "";
  if (s.form3.records.length === 0) {
    f3RowsPdf = "<tr><td colspan='7' style='text-align:center; padding:8px; color:#6b7280;'>ไม่มีรายการทวนสอบก๊าซในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var k = 0; k < s.form3.records.length; k++) {
      var r3 = s.form3.records[k];
      var bg3 = r3.isFail ? "#fee2e2" : (k % 2 === 0 ? "#ffffff" : "#f9fafb");
      var env = "T: " + r3.temp + "°C | P: " + r3.pressure + " | RH: " + r3.humidity + "%";
      f3RowsPdf += "<tr style='background-color:" + bg3 + ";'>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + r3.date + " " + r3.time + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px;'><b>" + escapeXml(r3.model) + "</b> (" + escapeXml(r3.serial) + ")</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:8px; text-align:center;'>" + escapeXml(env) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px; font-weight:bold;'>" + escapeXml(r3.maxErrorPct) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-weight:bold; font-size:9px; color:" + (r3.isFail ? "#b91c1c" : "#15803d") + ";'>" + escapeXml(r3.result) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(r3.calibratedBy) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(r3.approvedBy) + "</td>" +
      "</tr>";
    }
  }

  // หมวด 2: Sound Level Meter (PDF)
  var sf1RowsPdf = buildForm1PdfRows(s.soundForm1.records);
  var sf2RowsPdf = buildForm2PdfRows(s.soundForm2.records);
  var schkRowsPdf = buildChecklistPdfRows(s.soundChecklist.records);

  var sf3RowsPdf = "";
  if (s.soundForm3.records.length === 0) {
    sf3RowsPdf = "<tr><td colspan='7' style='text-align:center; padding:8px; color:#6b7280;'>ไม่มีรายการทวนสอบ Pre/Post เครื่องวัดเสียงในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var ks = 0; ks < s.soundForm3.records.length; ks++) {
      var rs3 = s.soundForm3.records[ks];
      var bgs3 = rs3.isFail ? "#fee2e2" : (ks % 2 === 0 ? "#ffffff" : "#f9fafb");
      var setupInfo = (rs3.calibrator || "-") + " | " + (rs3.freq || "1 kHz") + " | " + (rs3.freqWeighting || "A") + "/" + (rs3.timeWeighting || "F");
      sf3RowsPdf += "<tr style='background-color:" + bgs3 + ";'>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + rs3.date + "<br>" + rs3.preTime + " - " + rs3.postTime + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px;'><b>" + escapeXml(rs3.model) + "</b> (" + escapeXml(rs3.serial) + ")</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:8px; text-align:center;'>" + escapeXml(setupInfo) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px; font-weight:bold;'>" + escapeXml(rs3.maxDevDb) + " (Drift: " + escapeXml(rs3.maxDriftDb) + ")</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-weight:bold; font-size:9px; color:" + (rs3.isFail ? "#b91c1c" : "#15803d") + ";'>" + escapeXml(rs3.result) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(rs3.calibratedBy) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(rs3.approvedBy) + "</td>" +
      "</tr>";
    }
  }

  // หมวด 3: Heat Stress (PDF)
  var hf1RowsPdf = buildForm1PdfRows(s.heatForm1.records);
  var hf2RowsPdf = buildForm2PdfRows(s.heatForm2.records);
  var hchkRowsPdf = buildChecklistPdfRows(s.heatChecklist.records);

  var hf3RowsPdf = "";
  if (s.heatForm3.records.length === 0) {
    hf3RowsPdf = "<tr><td colspan='7' style='text-align:center; padding:8px; color:#6b7280;'>ไม่มีรายการทวนสอบเครื่องวัดความร้อนในรอบสัปดาห์นี้</td></tr>";
  } else {
    for (var kh = 0; kh < s.heatForm3.records.length; kh++) {
      var rh3 = s.heatForm3.records[kh];
      var bgh3 = rh3.isFail ? "#fee2e2" : (kh % 2 === 0 ? "#ffffff" : "#f9fafb");
      hf3RowsPdf += "<tr style='background-color:" + bgh3 + ";'>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px;'>" + rh3.date + " " + rh3.time + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px;'><b>" + escapeXml(rh3.model) + "</b> (" + escapeXml(rh3.serial) + ")</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:8px; text-align:center;'>" + escapeXml(rh3.refDevice) + "<br>" + escapeXml(rh3.method) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-size:9px; font-weight:bold;'>" + escapeXml(rh3.maxErrorC) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; text-align:center; font-weight:bold; font-size:9px; color:" + (rh3.isFail ? "#b91c1c" : "#15803d") + ";'>" + escapeXml(rh3.result) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(rh3.calibratedBy) + "</td>" +
        "<td style='padding:5px; border:1px solid #cbd5e1; font-size:9px; text-align:center;'>" + escapeXml(rh3.approvedBy) + "</td>" +
      "</tr>";
    }
  }

  return "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<style>" +
      "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');" +
      "body { font-family: 'Sarabun', Tahoma, sans-serif; color:#1e293b; margin:0; padding:15px; font-size:10px; }" +
      ".header { text-align:center; border-bottom:2px solid #1e40af; padding-bottom:8px; margin-bottom:12px; }" +
      ".header h1 { font-size:15px; margin:0 0 4px 0; color:#1e40af; font-weight:bold; }" +
      ".cat-banner { font-size:12px; font-weight:bold; color:#ffffff; background:#1e3a8a; padding:5px 10px; border-radius:4px; margin:14px 0 8px 0; }" +
      ".cat-banner.sound { background:#0f766e; } .cat-banner.heat { background:#c2410c; }" +
      ".section-title { font-size:10.5px; font-weight:bold; color:#1e3a8a; margin:10px 0 5px 0; border-left:3px solid #2563eb; padding-left:6px; }" +
      ".section-title.sound { color:#0f766e; border-left-color:#0d9488; } .section-title.heat { color:#c2410c; border-left-color:#ea580c; }" +
      "table { width:100%; border-collapse:collapse; margin-bottom:8px; }" +
      "th { background-color:#1e40af; color:#ffffff; padding:5px 4px; font-size:9px; border:1px solid #1e3a8a; text-align:center; }" +
      "table.sound th { background-color:#0f766e; border-color:#115e59; } table.heat th { background-color:#c2410c; border-color:#9a3412; }" +
      "td { padding:4px; font-size:9px; }" +
    "</style></head><body>" +
    "<div class='header'>" +
      "<div style='font-size:10px; font-weight:bold; color:#64748b; margin-bottom:2px;'>แผนกปฏิบัติการภาคสนาม | บริษัท อีสเทิร์น ไทย คอนซัลติ้ง 1992 จำกัด</div>" +
      "<h1>รายงานสรุปการสอบเทียบ บำรุงรักษา และทวนสอบเครื่อง Flue Gas Analyzer, เครื่องวัดเสียง และเครื่องวัดความร้อน ประจำสัปดาห์</h1>" +
      "<div style='font-size:10px; color:#475569;'>รอบวันที่: " + formatDateDisplay(s.cutoffDate) + " ถึง " + formatDateDisplay(s.today) + " &nbsp;|&nbsp; " +
      "กิจกรรมทั้งหมด: " + s.totalActivities + " รายการ &nbsp;|&nbsp; " +
      "<span style='color:" + (s.totalIssues > 0 ? "#b91c1c" : "#15803d") + "; font-weight:bold;'>" +
        (s.totalIssues > 0 ? "⚠️ พบรายการต้องติดตาม " + s.totalIssues + " รายการ" : "✓ อุปกรณ์ทุกเครื่องอยู่ในสภาพปกติ") +
      "</span></div>" +
    "</div>" +

    // PART I: FLUE GAS ANALYZER
    "<div class='cat-banner'>หมวดที่ 1: เครื่องวัดประสิทธิภาพการเผาไหม้ (Flue Gas Analyzer - Testo)</div>" +
    "<div class='section-title'>1.1 แบบฟอร์มที่ 1: บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ (" + s.form1.records.length + " รายการ)</div>" +
    "<table><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการตรวจสภาพ / สอบเทียบ</th><th style='width:60px;'>Correction</th><th style='width:75px;'>Tolerance</th><th style='width:45px;'>ผล</th><th style='width:80px;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + f1RowsPdf + "</tbody></table>" +

    "<div class='section-title'>1.2 แบบฟอร์มที่ 2: ประวัติความเสียหาย ความผิดปกติ หรือการซ่อมแซม (" + s.form2.records.length + " รายการ)</div>" +
    "<table><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการเสียหาย / ความผิดปกติ / ซ่อมแซม</th><th style='width:90px;'>ผู้ดำเนินการ</th><th style='width:110px;'>หมายเหตุ</th></tr></thead><tbody>" + f2RowsPdf + "</tbody></table>" +

    "<div class='section-title'>1.3 แบบฟอร์มที่ 3: Flue Gas Analyzer Verification (O2, CO, NO, NO2, SO2) (" + s.form3.records.length + " รายการ)</div>" +
    "<table><thead><tr><th style='width:80px;'>วัน/เวลา</th><th style='width:130px;'>รุ่น / S/N</th><th style='width:110px;'>สภาวะแวดล้อม</th><th style='width:65px;'>Max Error</th><th style='width:45px;'>ผล</th><th style='width:80px;'>ผู้ทวนสอบ</th><th style='width:80px;'>ผู้อนุมัติ</th></tr></thead><tbody>" + f3RowsPdf + "</tbody></table>" +

    "<div class='section-title'>1.4 ตรวจสอบสภาพเครื่องมือประจำสัปดาห์ (Checklist 5 รายการ) (" + s.checklist.records.length + " รายการ)</div>" +
    "<table><thead><tr><th style='width:60px;'>วันที่</th><th style='width:80px;'>ผู้ตรวจ</th><th style='width:120px;'>รุ่น / S/N</th><th style='width:50px;'>กรอง</th><th style='width:50px;'>ปั๊ม</th><th style='width:50px;'>แบต</th><th style='width:50px;'>สะอาด</th><th style='width:55px;'>เซ็นเซอร์</th><th>หมายเหตุ</th></tr></thead><tbody>" + chkRowsPdf + "</tbody></table>" +

    // PART II: SOUND LEVEL METER
    "<div class='cat-banner sound'>หมวดที่ 2: เครื่องวัดระดับเสียง (Sound Level Meter)</div>" +
    "<div class='section-title sound'>2.1 แบบฟอร์มที่ 1: บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ (" + s.soundForm1.records.length + " รายการ)</div>" +
    "<table class='sound'><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการตรวจสภาพ / สอบเทียบ</th><th style='width:60px;'>Correction</th><th style='width:75px;'>Tolerance</th><th style='width:45px;'>ผล</th><th style='width:80px;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + sf1RowsPdf + "</tbody></table>" +

    "<div class='section-title sound'>2.2 แบบฟอร์มที่ 2: ประวัติความเสียหาย ความผิดปกติ หรือการซ่อมแซม (" + s.soundForm2.records.length + " รายการ)</div>" +
    "<table class='sound'><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการเสียหาย / ความผิดปกติ / ซ่อมแซม</th><th style='width:90px;'>ผู้ดำเนินการ</th><th style='width:110px;'>หมายเหตุ</th></tr></thead><tbody>" + sf2RowsPdf + "</tbody></table>" +

    "<div class='section-title sound'>2.3 แบบฟอร์มที่ 3: Sound Level Meter Pre/Post Calibration Check (94 dB / 114 dB) (" + s.soundForm3.records.length + " รายการ)</div>" +
    "<table class='sound'><thead><tr><th style='width:80px;'>วัน/เวลา</th><th style='width:130px;'>รุ่น / S/N</th><th style='width:130px;'>Calibrator / Setup</th><th style='width:85px;'>Max Dev / Drift</th><th style='width:45px;'>ผล</th><th style='width:75px;'>ผู้ทวนสอบ</th><th style='width:75px;'>ผู้อนุมัติ</th></tr></thead><tbody>" + sf3RowsPdf + "</tbody></table>" +

    "<div class='section-title sound'>2.4 ตรวจสอบสภาพเครื่องวัดเสียงประจำสัปดาห์ (Checklist 5 รายการ) (" + s.soundChecklist.records.length + " รายการ)</div>" +
    "<table class='sound'><thead><tr><th style='width:60px;'>วันที่</th><th style='width:80px;'>ผู้ตรวจ</th><th style='width:120px;'>รุ่น / S/N</th><th style='width:55px;'>ตัวเครื่อง</th><th style='width:55px;'>ไมค์/โฟม</th><th style='width:50px;'>แบต</th><th style='width:55px;'>สาย/SD</th><th style='width:55px;'>การตอบสนอง</th><th>หมายเหตุ</th></tr></thead><tbody>" + schkRowsPdf + "</tbody></table>" +

    // PART III: HEAT STRESS MONITOR
    "<div class='cat-banner heat'>หมวดที่ 3: เครื่องวัดความร้อน (Heat Stress Monitor - WBGT / ISO 7243)</div>" +
    "<div class='section-title heat'>3.1 แบบฟอร์มที่ 1: บันทึกการสอบเทียบและบำรุงรักษาเครื่องมือ (" + s.heatForm1.records.length + " รายการ)</div>" +
    "<table class='heat'><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการตรวจสภาพ / สอบเทียบ</th><th style='width:60px;'>Correction</th><th style='width:75px;'>Tolerance</th><th style='width:45px;'>ผล</th><th style='width:80px;'>ผู้ดำเนินการ</th></tr></thead><tbody>" + hf1RowsPdf + "</tbody></table>" +

    "<div class='section-title heat'>3.2 แบบฟอร์มที่ 2: ประวัติความเสียหาย ความผิดปกติ หรือการซ่อมแซม (" + s.heatForm2.records.length + " รายการ)</div>" +
    "<table class='heat'><thead><tr><th style='width:65px;'>วันที่</th><th style='width:140px;'>รุ่น / S/N</th><th>รายการเสียหาย / ความผิดปกติ / ซ่อมแซม</th><th style='width:90px;'>ผู้ดำเนินการ</th><th style='width:110px;'>หมายเหตุ</th></tr></thead><tbody>" + hf2RowsPdf + "</tbody></table>" +

    "<div class='section-title heat'>3.3 แบบฟอร์มที่ 3: Heat Stress Verification (Dry Bulb, Wet Bulb, Globe ±0.5 °C) (" + s.heatForm3.records.length + " รายการ)</div>" +
    "<table class='heat'><thead><tr><th style='width:80px;'>วัน/เวลา</th><th style='width:130px;'>รุ่น / S/N</th><th style='width:130px;'>Ref / Method</th><th style='width:75px;'>Max Error</th><th style='width:45px;'>ผล</th><th style='width:75px;'>ผู้ทวนสอบ</th><th style='width:75px;'>ผู้อนุมัติ</th></tr></thead><tbody>" + hf3RowsPdf + "</tbody></table>" +

    "<div class='section-title heat'>3.4 ตรวจสอบสภาพเครื่องวัดความร้อนประจำสัปดาห์ (Checklist 5 รายการ) (" + s.heatChecklist.records.length + " รายการ)</div>" +
    "<table class='heat'><thead><tr><th style='width:60px;'>วันที่</th><th style='width:80px;'>ผู้ตรวจ</th><th style='width:120px;'>รุ่น / S/N</th><th style='width:55px;'>Wet Bulb</th><th style='width:55px;'>Globe</th><th style='width:50px;'>แบต</th><th style='width:55px;'>Dry Bulb</th><th style='width:55px;'>ตัวเครื่อง</th><th>หมายเหตุ</th></tr></thead><tbody>" + hchkRowsPdf + "</tbody></table>" +

    "<table style='width:100%; margin-top:25px; border:none;'>" +
      "<tr style='border:none;'>" +
        "<td style='width:50%; text-align:center; border:none; font-size:10px;'>" +
          "<div>ลงชื่อ ........................................................... ผู้รายงาน</div>" +
          "<div style='margin-top:4px;'>( หัวหน้าส่วนงานเครื่องมือวัด )</div>" +
        "</td>" +
        "<td style='width:50%; text-align:center; border:none; font-size:10px;'>" +
          "<div>ลงชื่อ ........................................................... ผู้รับทราบ / อนุมัติ</div>" +
          "<div style='margin-top:4px;'>( ผู้จัดการแผนกปฏิบัติการภาคสนาม )</div>" +
        "</td>" +
      "</tr>" +
    "</table>" +

    "</body></html>";
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
