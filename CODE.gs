const SHEET_STAFF = 'Staff';
const SHEET_LOG = 'Log';
const SHEET_USAGE = 'Usage';
const SHEET_SETTING = 'Setting';

// ดึงค่าการตั้งค่าจาก Sheet
function getSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SETTING);
    if(!sheet) return { startTime: 6, alertTime: 16.5, logoUrl: '', title: 'ระบบลงเวลา', sheetUrl: '' };
    
    const data = sheet.getDataRange().getValues();
    // C3=StartTime, C4=AlertTime(EndTime), C6=SheetUrl, C8=Logo, C9=Title
    let settings = {
      startTime: parseTimeToDecimal(data[2][2]), 
      alertTime: parseTimeToDecimal(data[3][2]), // ใช้เป็นเวลาสิ้นสุด Check-in ด้วย
      sheetUrl: data[5][2],                      
      logoUrl: data[7][2],                       
      title: data[8][2]                          
    };
    return settings;
  } catch(e) {
    return { startTime: 6, alertTime: 16.5, logoUrl: '', title: 'ระบบลงเวลา', sheetUrl: '' };
  }
}

// แปลงค่าเวลาเป็นทศนิยม (รองรับทั้ง . และ :)
function parseTimeToDecimal(timeVal) {
  if (!timeVal) return 0;
  if (typeof timeVal === 'number') {
    let d = new Date((timeVal - 25569) * 86400 * 1000); 
    return d.getHours() + (d.getMinutes() / 60);
  }
  if (typeof timeVal === 'string') {
    // รองรับทั้ง "16:30" และ "16.30"
    let parts = timeVal.split(/[.:]/);
    if (parts.length >= 2) {
      return parseInt(parts[0]) + (parseInt(parts[1]) / 60);
    }
  }
  return 0;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบลงเวลาปฏิบัติงาน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getServerTime() {
  return new Date().getTime(); 
}

function getInitData() {
  return {
    time: getServerTime(),
    settings: getSettings(),
    staff: getStaffData(),
    logs: getTodayLogs(),
    pending: getPendingCheckouts()
  };
}

function getTodayLogs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_LOG) || ss.getSheetByName('log'); 
    if (!sheet) return []; 
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 

    const todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy");
    let logs = [];
    
    for (let i = 1; i < data.length; i++) { 
      try {
        let logDateObj = data[i][0]; 
        if (!logDateObj) continue;
        let parsedDate = new Date(logDateObj);
        if (isNaN(parsedDate.getTime())) continue; 

        let logDateStr = Utilities.formatDate(parsedDate, "Asia/Bangkok", "dd/MM/yyyy");
        
        if (logDateStr === todayStr) {
          let name = data[i][1] ? String(data[i][1]).trim() : ""; 
          let timeInObj = data[i][3]; 
          let timeOutObj = data[i][4]; 
          
          let timeIn = "-";
          let timeOut = "-";
          
          if (timeInObj) {
            let pTimeIn = new Date(timeInObj);
            if (!isNaN(pTimeIn.getTime())) timeIn = Utilities.formatDate(pTimeIn, "Asia/Bangkok", "HH:mm") + " น.";
          }
          if (timeOutObj) {
            let pTimeOut = new Date(timeOutObj);
            if (!isNaN(pTimeOut.getTime())) timeOut = Utilities.formatDate(pTimeOut, "Asia/Bangkok", "HH:mm") + " น.";
          }
          logs.push({ name: name, timeIn: timeIn, timeOut: timeOut });
        }
      } catch(e) { continue; } 
    }
    return logs.reverse();
  } catch (error) {
    return [];
  }
}

function getPendingCheckouts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG) || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let pendings = [];
    const todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy");

    for (let i = 1; i < data.length; i++) {
      let timeOut = data[i][4];
      if (!timeOut || timeOut === "") {
        let name = data[i][1] ? String(data[i][1]).trim() : "";
        let dateObj = data[i][0];
        let dateStr = "-";
        let timeInObj = data[i][3];
        let timeInStr = "-";

        if (dateObj) {
            let d = new Date(dateObj);
            if(!isNaN(d.getTime())) dateStr = Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy");
        }
        if (timeInObj) {
            let t = new Date(timeInObj);
            if(!isNaN(t.getTime())) timeInStr = Utilities.formatDate(t, "Asia/Bangkok", "HH:mm") + " น.";
        }
        
        if (dateStr !== todayStr) pendings.push({ name: name, date: dateStr, timeIn: timeInStr });
      }
    }
    return pendings;
  } catch(e) { return []; }
}

function getStaffData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
  if(!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let staffObj = {};
  for (let i = 1; i < data.length; i++) {
    let name = data[i][0] ? String(data[i][0]).trim() : "";  
    let group = data[i][1] ? String(data[i][1]).trim() : ""; 
    let dept = data[i][2] ? String(data[i][2]).trim() : "";  
    if (!dept || !name) continue; 
    if (!staffObj[dept]) { staffObj[dept] = {}; }
    if (!staffObj[dept][group]) { staffObj[dept][group] = []; }
    staffObj[dept][group].push(name);
  }
  return staffObj;
}

function checkStatus(name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOG) || ss.getSheetByName('log');
    if (!logSheet) return { latestStatus: "ยังไม่ได้ปฏิบัติงาน", checkInTimeMs: null, hasPassword: false };

    const logData = logSheet.getDataRange().getValues();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy");
    
    let result = { 
      latestStatus: "ยังไม่ได้ปฏิบัติงาน", checkInTimeMs: null, finishedTimeMs: null, hasPassword: false,
      isPendingOld: false, pendingDateStr: null, pendingRow: null
    };

    const staffSheet = ss.getSheetByName(SHEET_STAFF);
    if(staffSheet) {
      const staffData = staffSheet.getDataRange().getValues();
      for(let i=1; i<staffData.length; i++) {
        if(String(staffData[i][0]).trim() === name) {
          if(staffData[i][3] && String(staffData[i][3]).trim() !== "") result.hasPassword = true;
          break;
        }
      }
    }

    for (let i = logData.length - 1; i > 0; i--) {
      let logName = logData[i][1] ? String(logData[i][1]).trim() : "";
      if (logName === name) {
        let logDateObj = logData[i][0]; 
        let timeOutObj = logData[i][4]; 
        let timeInObj = logData[i][3];  
        
        if (!timeOutObj || timeOutObj === "") {
           let parsedDate = new Date(logDateObj);
           if (isNaN(parsedDate.getTime())) continue;
           let logDateStr = Utilities.formatDate(parsedDate, "Asia/Bangkok", "dd/MM/yyyy");
           
           if (logDateStr === todayStr) {
             result.latestStatus = "เริ่มปฏิบัติงาน";
             if (timeInObj) {
                let pIn = new Date(timeInObj);
                if (!isNaN(pIn.getTime())) result.checkInTimeMs = pIn.getTime();
             }
             break;
           } else {
             result.latestStatus = "ค้างปฏิบัติงาน";
             result.isPendingOld = true;
             result.pendingDateStr = logDateStr;
             result.pendingRow = i + 1;
             if (timeInObj) {
                let pIn = new Date(timeInObj);
                if (!isNaN(pIn.getTime())) result.checkInTimeMs = pIn.getTime();
             }
             break; 
           }
        } else {
           let parsedDate = new Date(logDateObj);
           if (!isNaN(parsedDate.getTime())) {
             let logDateStr = Utilities.formatDate(parsedDate, "Asia/Bangkok", "dd/MM/yyyy");
             if (logDateStr === todayStr) {
               result.latestStatus = "เสร็จสิ้นการปฏิบัติงาน";
               let pOut = new Date(timeOutObj);
               if(!isNaN(pOut.getTime())) result.finishedTimeMs = pOut.getTime();
               break;
             }
           }
        }
      }
    }
    return result;
  } catch (error) {
    return { latestStatus: "ยังไม่ได้ปฏิบัติงาน", checkInTimeMs: null, hasPassword: false };
  }
}

function saveMissedCheckout(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOG) || ss.getSheetByName('log');
    if (!logSheet) return { status: 'error', message: 'ไม่พบ Sheet Log' };
    
    let row = formData.row;
    let timeStr = formData.timeStr; 
    
    let rowData = logSheet.getRange(row, 1, 1, 6).getValues()[0];
    let dateVal = rowData[0]; 
    let startTime = new Date(dateVal); 
    
    let parts = timeStr.split(':');
    let hours = parseInt(parts[0]);
    let mins = parseInt(parts[1]);
    
    let checkoutDate = new Date(startTime);
    checkoutDate.setHours(hours);
    checkoutDate.setMinutes(mins);
    checkoutDate.setSeconds(0);
    
    let checkInDate = new Date(rowData[3]); 
    let diffMs = checkoutDate.getTime() - checkInDate.getTime();
    if(diffMs < 0) diffMs = 0;
    let diffHrs = Math.floor(diffMs / 3600000);
    let diffMins = Math.floor((diffMs % 3600000) / 60000);
    let durationStr = diffHrs.toString().padStart(2, '0') + ":" + diffMins.toString().padStart(2, '0');
    
    logSheet.getRange(row, 5).setValue(checkoutDate);
    logSheet.getRange(row, 6).setValue(durationStr);
    
    return { status: 'success', message: 'บันทึกข้อมูลย้อนหลังเรียบร้อย' };
    
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function setupPassword(name, pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
    if(!sheet) return {status: 'error', message: 'ไม่พบ Sheet Staff'};
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]).trim() === name) {
        if(data[i][3] && String(data[i][3]).trim() !== "") return {status: 'error', message: 'ผู้ใช้นี้มีรหัสผ่านแล้ว'};
        sheet.getRange(i+1, 4).setValue(pin); 
        return {status: 'success', message: 'ตั้งรหัสผ่านสำเร็จ'};
      }
    }
    return {status: 'error', message: 'ไม่พบชื่อผู้ใช้'};
  } catch(e) {
    return {status: 'error', message: e.toString()};
  }
}

function verifyPassword(name, pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
    if(!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]).trim() === name) {
        let storedPin = data[i][3] ? String(data[i][3]).trim() : "";
        return storedPin === pin;
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

function saveLog(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(SHEET_LOG) || ss.getSheetByName('log');
    let usageSheet = ss.getSheetByName(SHEET_USAGE) || ss.getSheetByName('usage');

    if (!logSheet || !usageSheet) throw new Error("ไม่พบชีตฐานข้อมูล");

    const serverTime = new Date(); 
    const todayStr = Utilities.formatDate(serverTime, "Asia/Bangkok", "dd/MM/yyyy");
    
    const settings = getSettings();
    const currentHour = serverTime.getHours() + serverTime.getMinutes()/60;
    
    usageSheet.appendRow([serverTime, formData.name, formData.actionType]);

    if (formData.actionType === "เริ่มปฏิบัติงาน") {
      // *** แก้ไข: เช็คช่วงเวลา C3 - C4 ***
      if (currentHour < settings.startTime || currentHour > settings.alertTime) {
        return { status: 'error', message: 'blocked_time' };
      }

      const data = logSheet.getDataRange().getValues();
      for (let i = data.length - 1; i > 0; i--) {
         let logName = data[i][1] ? String(data[i][1]).trim() : "";
         if (logName === formData.name) {
           let logDateObj = data[i][0];
           if(logDateObj) {
             let logDateStr = Utilities.formatDate(new Date(logDateObj), "Asia/Bangkok", "dd/MM/yyyy");
             if (logDateStr === todayStr) {
               let existingTimeOut = data[i][4];
               if (existingTimeOut && existingTimeOut !== "") return { status: 'error', message: 'ท่านได้ลงเวลาออกงานไปแล้วในวันนี้' };
               else return { status: 'error', message: 'ท่านได้ลงเวลาเข้างานไปแล้วในวันนี้' };
             }
           }
         }
      }
      logSheet.appendRow([serverTime, formData.name, formData.dept, serverTime, "", ""]);
      return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    } 
    else if (formData.actionType === "เสร็จสิ้นการปฏิบัติงาน") {
      const data = logSheet.getDataRange().getValues();
      let rowToUpdate = -1;
      let checkInTimeObj = null;

      for (let i = data.length - 1; i > 0; i--) {
        try {
          let logName = data[i][1] ? String(data[i][1]).trim() : "";
          let existingTimeOut = data[i][4]; 
          
          if (logName === formData.name && (!existingTimeOut || existingTimeOut === "")) {
            let logDateObj = data[i][0];
            if(logDateObj) {
              let logDateStr = Utilities.formatDate(new Date(logDateObj), "Asia/Bangkok", "dd/MM/yyyy");
              if (logDateStr === todayStr) {
                rowToUpdate = i + 1; 
                checkInTimeObj = data[i][3]; 
                break; 
              }
            }
          }
        } catch(e) { continue; }
      }

      if (rowToUpdate !== -1) {
        let startTime = checkInTimeObj ? new Date(checkInTimeObj) : serverTime;
        if (isNaN(startTime.getTime())) startTime = serverTime;
        let diffMs = serverTime.getTime() - startTime.getTime();
        if (diffMs < 0) diffMs = 0;
        let diffHrs = Math.floor(diffMs / 3600000);
        let diffMins = Math.floor((diffMs % 3600000) / 60000);
        let durationStr = diffHrs.toString().padStart(2, '0') + ":" + diffMins.toString().padStart(2, '0');

        logSheet.getRange(rowToUpdate, 5).setValue(serverTime); 
        logSheet.getRange(rowToUpdate, 6).setValue(durationStr); 
        return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
      } else {
         return { status: 'error', message: 'ไม่พบข้อมูลการเริ่มปฏิบัติงานในวันนี้' };
      }
    }
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}
