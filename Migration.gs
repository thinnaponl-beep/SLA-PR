// ==========================================
// สคริปต์สำหรับยิงข้อมูลจาก Google Sheets เข้า Supabase
// ==========================================

function migrateCasesDataToSupabase() {
  // 🌟 1. นำ URL และ Key ของคุณมาประกาศไว้ "ข้างใน" ฟังก์ชัน 
  const MIG_URL = 'https://ddixdzgjaiuqqchxvqnc.supabase.co'; 
  const MIG_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRkaXhkemdqYWl1cXFjaHh2cW5jIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ0ODk0NzcsImV4cCI6MjA5MDA2NTQ3N30.iXZB-Fm2zWX0qT5jrr3_dC6OzeCYbVjGvgdZcNyo_Bk';

  // 🌟 เปลี่ยนชื่อชีตเป้าหมายเป็น Database_Cases
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Database_Cases");
  
  if (!sheet) {
    try { SpreadsheetApp.getUi().alert("ไม่พบชีต Database_Cases"); } catch(e){}
    return;
  }

  // ดึงข้อมูลทั้งหมดมา 
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getDisplayValues();
  const payloadArray = [];

  Logger.log(`กำลังเตรียมย้ายข้อมูลจำนวน ${data.length} แถว...`);

  // ฟังก์ชันแปลงวันที่และเวลา ให้เป็นมาตรฐาน SQL
  function cleanDateTime(dateStr) {
    let str = String(dateStr).trim().replace(/,/g, ''); 
    if (!str || str === "" || str === "-") return null; 

    let parts = str.split(' ');
    
    if (parts.length >= 1) {
      let datePart = parts[0].split('/');
      if (datePart.length === 3) {
        let d = datePart[0].padStart(2, '0');
        let m = datePart[1].padStart(2, '0');
        let y = parseInt(datePart[2]);
        if (y > 2400) y -= 543; // แปลง พ.ศ. เป็น ค.ศ.
        
        let timePart = parts[1] ? parts[1] : '00:00:00';
        let timeArray = timePart.split(':');
        let hh = (timeArray[0] || '00').padStart(2, '0');
        let min = (timeArray[1] || '00').padStart(2, '0');
        let ss = (timeArray[2] || '00').padStart(2, '0');

        return `${y}-${m}-${d} ${hh}:${min}:${ss}`; 
      }
      
      if (parts[0].includes('-')) return str;
    }
    return null; 
  }

  // วนลูปจับคู่ข้อมูลเข้าชื่อคอลัมน์ของ Supabase
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // 🌟 ป้องกัน Error คอลัมน์ว่าง
    let caseId = String(row[0] || "").trim();
    if (!caseId) continue; 

    payloadArray.push({
      case_id: caseId,
      status: String(row[1] || "").trim(),
      time_created: cleanDateTime(row[2]),
      time_accepted: cleanDateTime(row[3]),
      time_closed: cleanDateTime(row[4]),
      creator: String(row[5] || "").trim(),
      assignee: String(row[6] || "").trim(),
      maid_id: String(row[7] || "").trim(),
      maid_name: String(row[8] || "").trim(),
      topic: String(row[9] || "").trim(),       // 🌟 บังคับเป็น Text ธรรมดา
      chat_link: String(row[10] || "").trim(),
      action_details: String(row[11] || "").trim(),
      history_logs: String(row[12] || "").trim() // 🌟 บังคับเป็น Text ธรรมดา
    });
  }

  const chunkSize = 500;
  for (let i = 0; i < payloadArray.length; i += chunkSize) {
    const chunk = payloadArray.slice(i, i + chunkSize);
    
    const options = {
      method: 'post',
      headers: {
        'apikey': MIG_KEY,
        'Authorization': 'Bearer ' + MIG_KEY,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=minimal'
      },
      payload: JSON.stringify(chunk),
      muteHttpExceptions: true
    };

    // ยิง API ไปยังตาราง Database_Cases
    const response = UrlFetchApp.fetch(MIG_URL + '/rest/v1/Database_Cases', options);
    
    // 🌟 แจ้งเตือน Error กลางหน้าจอ Google Sheets ถ้ามีปัญหา
    if (response.getResponseCode() !== 201 && response.getResponseCode() !== 200) {
      let errorMsg = response.getContentText();
      Logger.log(`❌ Error: ${errorMsg}`);
      try {
          SpreadsheetApp.getUi().alert(`❌ เกิดข้อผิดพลาดตอนอัปโหลด:\n\n${errorMsg}\n\n(ลองเช็คชนิดข้อมูลคอลัมน์ใน Supabase ดูนะครับ)`);
      } catch(e){}
      return; // หยุดการทำงานทันทีถ้าเจอ Error
    }
  }

  Logger.log("🎉 การย้ายข้อมูลเสร็จสมบูรณ์!");
  
  // 🌟 เด้ง Popup แจ้งเตือนเมื่อสำเร็จ
  try {
     SpreadsheetApp.getUi().alert("🎉 การย้ายข้อมูลเสร็จสมบูรณ์ 100%! ไปเช็คใน Supabase ได้เลยครับ");
  } catch(e){}
}
