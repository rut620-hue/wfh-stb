// ==========================================
// ส่วนของการทำ API Router เพื่อเชื่อมกับ GitHub Pages
// ==========================================

// รองรับการดึงข้อมูลแบบ GET (เช่น การดึงข้อมูลเริ่มต้นตอนโหลดหน้าเว็บ)
function doGet(e) {
  try {
    let action = e.parameter.action;
    let result = {};
    
    // แยกประเภทคำสั่งที่ส่งมา
    if (action === 'getInitData') {
      result = getInitData();
    } else {
      result = { status: 'success', message: 'API is running. Ready to connect with GitHub Pages.' };
    }
    
    // ส่งข้อมูลกลับไปเป็น JSON
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// รองรับการส่งข้อมูลแบบ POST (เช่น การบันทึกเวลา, ตรวจสอบสถานะ, ตั้งรหัสผ่าน)
function doPost(e) {
  try {
    // แปลงข้อมูลที่หน้าเว็บ GitHub ส่งมาให้อยู่ในรูปแบบ JSON
    let params = JSON.parse(e.postData.contents);
    let action = params.action;
    let data = params.data; // ข้อมูลที่แนบมาด้วย เช่น ชื่อ, รหัสผ่าน
    let result = {};

    // แยกประเภทคำสั่งและเรียกใช้ฟังก์ชันเดิมที่มีอยู่
    if (action === 'checkStatus') {
      result = checkStatus(data.name);
    } else if (action === 'saveLog') {
      result = saveLog(data);
    } else if (action === 'saveMissedCheckout') {
      result = saveMissedCheckout(data);
    } else if (action === 'setupPassword') {
      result = setupPassword(data.name, data.pin);
    } else if (action === 'verifyPassword') {
      result = verifyPassword(data.name, data.pin);
    } else {
      result = { status: 'error', message: 'ไม่พบคำสั่งนี้ในระบบ' };
    }

    // ส่งข้อมูลกลับไปเป็น JSON และอนุญาตให้ข้ามโดเมนได้ (CORS)
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
