const SPREADSHEET_ID = "1Gs6cnHjpI-5TOfiNinqiuWLb5Q0oXvjUNM3m2uqLFoM";

// ชื่อชีตต่างๆ
const SHEET_EXCHANGE_LOG = "บันทึกการแลกคาบ";
const SHEET_TEACHERS = "รายชื่อครู";
const SHEET_SCHEDULE = "ตารางสอน";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบแลกเปลี่ยนคาบสอน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- Helper Functions ---
function parseDateLocal(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split('-');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

function getThaiDayName(date) {
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  return days[date.getDay()];
}

function setupSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetLog = ss.getSheetByName(SHEET_EXCHANGE_LOG);
    if (!sheetLog) {
      sheetLog = ss.insertSheet(SHEET_EXCHANGE_LOG);
      sheetLog.appendRow(["Timestamp", "วันที่ขอแลก", "ผู้ขอแลก", "คาบ", "รหัสวิชา", "ชั้น", "ประเภท", "เหตุผล", "กลุ่มสาระ", "วันที่ไปแลก", "ผู้รับแลก", "คาบ", "รหัสวิชา", "ชั้น", "สถานที่แลก"]);
    }
    let sheetTeachers = ss.getSheetByName(SHEET_TEACHERS);
    if (!sheetTeachers) {
      sheetTeachers = ss.insertSheet(SHEET_TEACHERS);
      sheetTeachers.appendRow(["ชื่อ-สกุล", "กลุ่มสาระ", "ตำแหน่ง", "วิทยฐานะ"]);
    }
    let sheetSchedule = ss.getSheetByName(SHEET_SCHEDULE);
    if (!sheetSchedule) {
      sheetSchedule = ss.insertSheet(SHEET_SCHEDULE);
      sheetSchedule.appendRow(["ชื่อครู", "วัน", "คาบ", "เวลา", "รหัสวิชา", "ชื่อวิชา", "ห้องเรียน", "ประเภทคาบ"]);
    }
    return "Setup Complete";
  } catch (e) { return "Error: " + e.toString(); }
}

function getMasterData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_TEACHERS);
    if (!sheet) { setupSheets(); return { departments: [], teachers: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { departments: [], teachers: [] };
    const headers = data.shift();
    const idxName = headers.indexOf("ชื่อ-สกุล");
    const idxDept = headers.indexOf("กลุ่มสาระ");
    let departments = new Set();
    let teachers = [];
    data.forEach(row => {
      const name = row[idxName];
      const dept = row[idxDept];
      if (name && dept) {
        departments.add(dept);
        teachers.push({ name: name, dept: dept });
      }
    });
    return {
      departments: Array.from(departments).sort(),
      teachers: teachers.sort((a, b) => a.name.localeCompare(b.name))
    };
  } catch (e) { return { departments: [], teachers: [] }; }
}

function getSchedule(teacherName, dateOrDayStr) {
  try {
    if (!dateOrDayStr) return [];
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SCHEDULE);
    if (!sheet) return [];

    let targetDayName = "";
    const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    
    if (days.includes(dateOrDayStr)) {
       targetDayName = dateOrDayStr;
    } else {
       const date = parseDateLocal(dateOrDayStr);
       if (date) targetDayName = getThaiDayName(date);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const col = {
      teacher: headers.indexOf("ชื่อครู"), day: headers.indexOf("วัน"), period: headers.indexOf("คาบ"),
      time: headers.indexOf("เวลา"), code: headers.indexOf("รหัสวิชา"), name: headers.indexOf("ชื่อวิชา"),
      room: headers.indexOf("ห้องเรียน"), type: headers.indexOf("ประเภทคาบ")
    };
    let schedules = [];
    data.forEach(row => {
      if (row[col.teacher] == teacherName && row[col.day] == targetDayName) {
        schedules.push({
          period: String(row[col.period]), time: row[col.time], subjectCode: row[col.code],
          subjectName: row[col.name], classRoom: row[col.room], type: row[col.type]
        });
      }
    });
    return schedules.sort((a, b) => parseInt(String(a.period).split('-')[0]) - parseInt(String(b.period).split('-')[0]));
  } catch (e) { return []; }
}

// --- 3. ค้นหาคู่แลก (Logic ใหม่: วนลูป จันทร์-ศุกร์ ไม่กำหนดวันที่) ---
function findMatchingCandidatesAuto(criteria, reqDayName, requesterName) {
  try {
    if (!reqDayName) return [];
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SCHEDULE);
    if (!sheet) return [];

    const allData = sheet.getDataRange().getValues();
    const headers = allData.shift();
    
    const col = {
      teacher: headers.indexOf("ชื่อครู"), day: headers.indexOf("วัน"), period: headers.indexOf("คาบ"),
      time: headers.indexOf("เวลา"), code: headers.indexOf("รหัสวิชา"), name: headers.indexOf("ชื่อวิชา"),
      room: headers.indexOf("ห้องเรียน"), type: headers.indexOf("ประเภทคาบ")
    };

    let results = [];
    const targetDays = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];

    const getOccupiedRow = (dayName, room, p) => {
        return allData.find(row => {
            if (row[col.day] != dayName || row[col.room] != room) return false;
            let pStr = String(row[col.period]);
            if (pStr.includes('-')) {
                const [start, end] = pStr.split('-').map(Number);
                return p >= start && p <= end;
            } else {
                return parseInt(pStr) === p;
            }
        });
    };

    const isTeacherBusy = (tName, dayName, p, ignorePeriodStr) => {
        return allData.some(row => {
            if (String(row[col.teacher]).trim() != String(tName).trim() || row[col.day] != dayName) return false;
            let rowPStr = String(row[col.period]);
            if (ignorePeriodStr && rowPStr === ignorePeriodStr) return false;
            if (rowPStr.includes('-')) {
                const [start, end] = rowPStr.split('-').map(Number);
                return p >= start && p <= end;
            } else {
                return parseInt(rowPStr) === p;
            }
        });
    };

    const isTargetFreeForMySlot = (targetName, myDayName, myPeriodStr, ignoreTargetPeriod) => {
         let periodsToCheck = [];
         if (String(myPeriodStr).includes('-')) {
             const [s, e] = String(myPeriodStr).split('-').map(Number);
             for(let i=s; i<=e; i++) periodsToCheck.push(i);
         } else {
             periodsToCheck.push(parseInt(myPeriodStr));
         }
         for (let p of periodsToCheck) {
             if (isTeacherBusy(targetName, myDayName, p, ignoreTargetPeriod)) return false;
         }
         return true;
    };

    // วนลูป 5 วัน (จันทร์-ศุกร์)
    for (let i = 0; i < targetDays.length; i++) {
      const currentDayName = targetDays[i];

      for (let p = 1; p <= 9; p++) {
          
          // 1. คุณ (Requester) ต้องว่างในคาบที่จะย้ายไป (ในวัน currentDayName)
          // ถ้า currentDayName คือวันเดียวกับ reqDayName เราจะ ignore คาบเดิมของเรา (criteria.period) เพราะเรากำลังย้ายออก
          const ignoreReqPeriod = (currentDayName === reqDayName) ? criteria.period : null;
          
          if (isTeacherBusy(requesterName, currentDayName, p, ignoreReqPeriod)) continue;

          if (criteria.type === "คาบคู่") {
              if (p >= 9) continue;
              if (isTeacherBusy(requesterName, currentDayName, p + 1, ignoreReqPeriod)) continue;
          }

          const occupiedRow = getOccupiedRow(currentDayName, criteria.classRoom, p);
          
          if (occupiedRow) {
              let pStr = String(occupiedRow[col.period]);
              let pStart = parseInt(pStr.split('-')[0]);
              
              if (pStart === p && occupiedRow[col.type] === criteria.type) {
                  const targetName = occupiedRow[col.teacher];
                  
                  // 2. คู่แลก (Target) ต้องว่างในเวลาเดิมของคุณ (reqDayName)
                  // ถ้าวันเดียวกัน เราจะ ignore คาบของ Target ที่เขากำลังจะย้ายออก (pStr)
                  const ignoreTargetPeriod = (currentDayName === reqDayName) ? pStr : null;

                  if (isTargetFreeForMySlot(targetName, reqDayName, criteria.period, ignoreTargetPeriod)) {
                       results.push({
                          status: "Swap", 
                          date: currentDayName, // ส่งชื่อวันกลับไปแทนวันที่
                          teacher: targetName,
                          period: pStr, 
                          time: occupiedRow[col.time], 
                          subjectCode: occupiedRow[col.code],
                          subjectName: occupiedRow[col.name], 
                          classRoom: occupiedRow[col.room], 
                          type: occupiedRow[col.type]
                       });
                  }
              }
          } 
      }
    }
    
    return results; // ส่งผลลัพธ์กลับ (Frontend จะเรียงวันเอง)
  } catch (e) { Logger.log(e); return []; }
}

// --- 4. บันทึกข้อมูล ---
function processForm(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_EXCHANGE_LOG);
    if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_EXCHANGE_LOG); }

    sheet.appendRow([
      new Date(), data.requester.date, data.requester.name, data.requester.period,
      data.requester.subjectCode, data.requester.classRoom, data.requester.type,
      data.requester.reason || "-", data.requester.dept || "-",
      data.target.date, data.target.teacher, data.target.period,
      data.target.subjectCode, data.target.classRoom, data.requester.location || "-"
    ]);
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อย' };
  } catch (e) { return { status: 'error', message: "Error: " + e.toString() }; }
}
