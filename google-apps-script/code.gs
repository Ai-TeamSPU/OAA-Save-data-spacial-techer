/**
 * ===================================================
 *  Google Apps Script — รับข้อมูลอาจารย์พิเศษ (v2 - แก้ CORS)
 *  วิธีใช้:
 *  1. เปิด Google Sheet → Extensions > Apps Script
 *  2. วางโค้ดนี้ทั้งหมด แทนที่ code เดิม → Save
 *  3. Deploy > New deployment > Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Copy "Web app URL" ไปวางใน Settings ของโปรแกรม
 *  ** ถ้า Deploy แล้วต้องทำ "New deployment" ใหม่ทุกครั้งที่แก้โค้ด **
 * ===================================================
 */

var SHEET_NAME = "อาจารย์พิเศษ";

// รับข้อมูลผ่าน GET (แก้ปัญหา CORS จาก browser)
function doGet(e) {
  try {
    var payload = e.parameter.data;
    if (!payload) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", message: "Apps Script พร้อมใช้งาน" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var data = JSON.parse(payload);
    return saveData(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// รับข้อมูลผ่าน POST (สำหรับ same-origin / server-side)
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    return saveData(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = [
      "วันที่บันทึก", "ภาคการศึกษา", "คณะ", "สาขาวิชา",
      "คำนำหน้า", "ชื่อ (ไทย)", "นามสกุล (ไทย)", "First Name", "Last Name",
      "วุฒิ 1 (ระดับ)", "วุฒิ 1 (ปี)", "วุฒิ 1 (หลักสูตร)", "วุฒิ 1 (สาขา)", "วุฒิ 1 (สถาบัน)",
      "วุฒิ 2 (ระดับ)", "วุฒิ 2 (ปี)", "วุฒิ 2 (หลักสูตร)", "วุฒิ 2 (สาขา)", "วุฒิ 2 (สถาบัน)",
      "วุฒิ 3 (ระดับ)", "วุฒิ 3 (ปี)", "วุฒิ 3 (หลักสูตร)", "วุฒิ 3 (สาขา)", "วุฒิ 3 (สถาบัน)",
      "ประสบการณ์ 1 (ตำแหน่ง)", "ประสบการณ์ 1 (บริษัท)", "ประสบการณ์ 1 (เริ่ม)", "ประสบการณ์ 1 (สิ้นสุด)", "ประสบการณ์ 1 (ระยะเวลา)",
      "รายวิชา 1 (รหัส)", "รายวิชา 1 (ชื่อ)", "รายวิชา 1 (หน่วยกิต)",
      "รายวิชา 2 (รหัส)", "รายวิชา 2 (ชื่อ)", "รายวิชา 2 (หน่วยกิต)",
      "สัดส่วนการสอน", "ชั่วโมง/สัปดาห์",
      "กลุ่ม", "ประเภทอาจารย์พิเศษ", "รายละเอียด",
      "ความเชี่ยวชาญ", "หมายเหตุ",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  var edu = data.educations || [];
  var work = data.experiences || [];
  var courses = data.courses || [];
  var subs = data.qualSubs || {};

  var qualGroup = "";
  var qualDetailsArr = [];

  if (subs.qual1_a || subs.qual1_b) { 
    qualGroup = "กลุ่มที่ 1"; 
    if (subs.qual1_a) qualDetailsArr.push("คุณวุฒิปริญญาโท สอดคล้องตามเกณฑ์มาตรฐานหลักสูตร พ.ศ. 2565");
    if (subs.qual1_b) {
      var branchStr = data.qualFields && data.qualFields.qual1_branch ? " (" + data.qualFields.qual1_branch + ")" : "";
      qualDetailsArr.push("มีตำแหน่ง ศาสตราจารย์ รองศาสตราจารย์ หรือ ผู้ช่วยศาสตราจารย์ ในสาขาวิชา" + branchStr);
    }
  }
  else if (subs.qual2_a || subs.qual2_c) { 
    qualGroup = "กลุ่มที่ 2"; 
    if (subs.qual2_a) qualDetailsArr.push("คุณวุฒิปริญญาตรี และมีประสบการณ์ทำงานภาคอุตสาหกรรม อย่างต่อเนื่องมาแล้ว 5 ปีขึ้นไป");
    if (subs.qual2_c) qualDetailsArr.push("มีความรู้และประสบการณ์เป็นที่ยอมรับซึ่งตรงหรือสัมพันธ์กับรายวิชาที่สอน หรือผลงานเป็นที่ประจักษ์ในวิชาชีพ");
  }
  else if (subs.qual3_a) { 
    qualGroup = "กลุ่มที่ 3"; 
    if (subs.qual3_a) qualDetailsArr.push("คุณวุฒิปริญญาตรี และมีประสบการณ์ทำงานภาคอุตสาหกรรม น้อยกว่า 5 ปี");
  }

  var qualDetails = qualDetailsArr.join(", ");

  var teacherType = "";
  if (qualGroup === "กลุ่มที่ 1") teacherType = "อาจารย์พิเศษ";
  else if (qualGroup === "กลุ่มที่ 2") teacherType = "อาจารย์พิเศษร่วมสอน";
  else if (qualGroup === "กลุ่มที่ 3") teacherType = "อาจารย์พิเศษช่วยสอน";

  var titlePrefix = data.titlePrefix === "อื่นๆ / Other" ? (data.titleCustom || "") : (data.titlePrefix || "");

  // คำนวณระยะเวลาประสบการณ์
  function calcDur(start, end) {
    if (!start) return "";
    var s = new Date(start), e2 = end ? new Date(end) : new Date();
    var months = (e2.getFullYear() - s.getFullYear()) * 12 + (e2.getMonth() - s.getMonth());
    if (months < 0) return "";
    var yrs = Math.floor(months / 12), rem = months % 12;
    if (yrs === 0 && rem === 0) return "น้อยกว่า 1 เดือน";
    if (yrs === 0) return rem + " เดือน";
    if (rem === 0) return yrs + " ปี";
    return yrs + " ปี " + rem + " เดือน";
  }

  var sheetHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];

  // ตรวจสอบและเพิ่มคอลัมน์ประสบการณ์แบบไดนามิก
  var insertIndexWork = sheetHeaders.indexOf("รายวิชา 1 (รหัส)");
  if (insertIndexWork === -1) insertIndexWork = sheetHeaders.indexOf("รายวิชา 1 (รหัส-ชื่อ)");
  if (insertIndexWork === -1) insertIndexWork = sheetHeaders.length; // fallback
  
  var headersToAddWork = [];
  for (var w = 1; w <= work.length; w++) {
    var posHeader = "ประสบการณ์ " + w + " (ตำแหน่ง)";
    if (sheetHeaders.indexOf(posHeader) === -1) {
      headersToAddWork.push("ประสบการณ์ " + w + " (ตำแหน่ง)");
      headersToAddWork.push("ประสบการณ์ " + w + " (บริษัท)");
      headersToAddWork.push("ประสบการณ์ " + w + " (เริ่ม)");
      headersToAddWork.push("ประสบการณ์ " + w + " (สิ้นสุด)");
      headersToAddWork.push("ประสบการณ์ " + w + " (ระยะเวลา)");
    }
  }

  if (headersToAddWork.length > 0) {
    sheet.insertColumnsBefore(insertIndexWork + 1, headersToAddWork.length);
    sheet.getRange(1, insertIndexWork + 1, 1, headersToAddWork.length).setValues([headersToAddWork]);
    sheet.getRange(1, insertIndexWork + 1, 1, headersToAddWork.length).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // ตรวจสอบและเพิ่มคอลัมน์รายวิชาแบบไดนามิก
  var insertIndex = sheetHeaders.indexOf("สัดส่วนการสอน");
  if (insertIndex === -1) insertIndex = sheetHeaders.length; // fallback
  var headersToAdd = [];
  for (var c = 1; c <= courses.length; c++) {
    var codeHeader = "รายวิชา " + c + " (รหัส)";
    if (sheetHeaders.indexOf(codeHeader) === -1) {
      headersToAdd.push("รายวิชา " + c + " (รหัส)");
      headersToAdd.push("รายวิชา " + c + " (ชื่อ)");
      headersToAdd.push("รายวิชา " + c + " (หน่วยกิต)");
      headersToAdd.push("รายวิชา " + c + " (จำนวนครั้งที่สอน)");
    }
  }

  if (headersToAdd.length > 0) {
    sheet.insertColumnsBefore(insertIndex + 1, headersToAdd.length);
    sheet.getRange(1, insertIndex + 1, 1, headersToAdd.length).setValues([headersToAdd]);
    sheet.getRange(1, insertIndex + 1, 1, headersToAdd.length).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    // โหลด headers ใหม่หลังจากเพิ่มคอลัมน์
    sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  function getCourseParts(c) {
    if (!c || !c.subject) return { code: "", name: "" };
    var parts = c.subject.split(" - ");
    return {
      code: parts[0] ? parts[0].trim() : "",
      name: parts.slice(1).join(" - ").trim()
    };
  }

  var rowData = {
    "วันที่บันทึก": new Date().toLocaleDateString("th-TH"),
    "ภาคการศึกษา": data.semester || "",
    "คณะ": data.faculty || "",
    "สาขาวิชา": data.branch || "",
    "คำนำหน้า": titlePrefix,
    "ชื่อ (ไทย)": data.firstNameTH || "",
    "นามสกุล (ไทย)": data.lastNameTH || "",
    "First Name": data.firstNameEN || "",
    "Last Name": data.lastNameEN || "",
    "วุฒิ 1 (ระดับ)": edu[0] ? edu[0].level : "",
    "วุฒิ 1 (ปี)": edu[0] ? edu[0].year : "",
    "วุฒิ 1 (หลักสูตร)": edu[0] ? edu[0].curriculum : "",
    "วุฒิ 1 (สาขา)": edu[0] ? edu[0].major : "",
    "วุฒิ 1 (สถาบัน)": edu[0] ? edu[0].institution : "",
    "วุฒิ 2 (ระดับ)": edu[1] ? edu[1].level : "",
    "วุฒิ 2 (ปี)": edu[1] ? edu[1].year : "",
    "วุฒิ 2 (หลักสูตร)": edu[1] ? edu[1].curriculum : "",
    "วุฒิ 2 (สาขา)": edu[1] ? edu[1].major : "",
    "วุฒิ 2 (สถาบัน)": edu[1] ? edu[1].institution : "",
    "วุฒิ 3 (ระดับ)": edu[2] ? edu[2].level : "",
    "วุฒิ 3 (ปี)": edu[2] ? edu[2].year : "",
    "วุฒิ 3 (หลักสูตร)": edu[2] ? edu[2].curriculum : "",
    "วุฒิ 3 (สาขา)": edu[2] ? edu[2].major : "",
    "วุฒิ 3 (สถาบัน)": edu[2] ? edu[2].institution : "",
    "สัดส่วนการสอน": courses.map(c => c.proportion || "ไม่ระบุ").join(", "),
    "ชั่วโมง/สัปดาห์": data.teachingHours || "",
    "กลุ่ม": qualGroup,
    "ประเภทอาจารย์พิเศษ": teacherType,
    "รายละเอียด": qualDetails,
    "ความเชี่ยวชาญ": data.expertise || "",
    "หมายเหตุ": data.note || ""
  };

  // แมพข้อมูลประสบการณ์แบบไดนามิก
  work.forEach(function(w, index) {
    var num = index + 1;
    rowData["ประสบการณ์ " + num + " (ตำแหน่ง)"] = w.position || "";
    rowData["ประสบการณ์ " + num + " (บริษัท)"] = w.company || "";
    rowData["ประสบการณ์ " + num + " (เริ่ม)"] = w.startDate || "";
    rowData["ประสบการณ์ " + num + " (สิ้นสุด)"] = w.isCurrent ? "ปัจจุบัน" : (w.endDate || "");
    rowData["ประสบการณ์ " + num + " (ระยะเวลา)"] = calcDur(w.startDate, w.isCurrent ? null : w.endDate);
  });

  // แมพข้อมูลรายวิชาแบบไดนามิก
  courses.forEach(function(c, index) {
    var cParts = getCourseParts(c);
    var num = index + 1;
    rowData["รายวิชา " + num + " (รหัส)"] = cParts.code;
    rowData["รายวิชา " + num + " (ชื่อ)"] = cParts.name;
    rowData["รายวิชา " + num + " (หน่วยกิต)"] = c.credits || "";
    rowData["รายวิชา " + num + " (จำนวนครั้งที่สอน)"] = c.teachCount || "";
  });

  var finalRow = [];
  for (var i = 0; i < sheetHeaders.length; i++) {
    var h = sheetHeaders[i] ? sheetHeaders[i].toString().trim() : "";
    finalRow.push(rowData[h] !== undefined ? rowData[h] : "");
  }

  if (finalRow.length > 0) {
    sheet.appendRow(finalRow);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "บันทึกข้อมูลสำเร็จ" }))
    .setMimeType(ContentService.MimeType.JSON);
}
