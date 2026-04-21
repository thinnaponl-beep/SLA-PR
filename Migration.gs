// ==========================================
// สคริปต์สำหรับย้ายข้อมูล Config และ Maids ไป Supabase
// (ใช้ย้ายข้อมูลเดิมที่ค้างอยู่ใน Google Sheets)
// ==========================================


function migrateRemainingDataToSupabase() {
  const MIG_URL = ''; 
  const MIG_KEY = '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ----------------------------------------
  // 1. ดึงและส่งข้อมูลรายชื่อแม่บ้าน (Database_Maids)
  // ----------------------------------------
  const maidSheet = ss.getSheetByName("Database_Maids");
  if (maidSheet && maidSheet.getLastRow() > 1) {
    Logger.log("เริ่มดึงข้อมูลแม่บ้าน...");
    const maidData = maidSheet.getRange(2, 1, maidSheet.getLastRow() - 1, 3).getDisplayValues();
    const maidPayload = maidData.map(r => ({
      maid_id: String(r[0] || "").trim(),
      maid_name: String(r[1] || "").trim(),
      note: String(r[2] || "").trim()
    })).filter(m => m.maid_id !== "");
    
    if (maidPayload.length > 0) {
      const chunkSize = 500;
      for (let i = 0; i < maidPayload.length; i += chunkSize) {
          const chunk = maidPayload.slice(i, i + chunkSize);
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
          UrlFetchApp.fetch(MIG_URL + '/rest/v1/Database_Maids', options);
      }
      Logger.log(`✅ ส่งข้อมูลแม่บ้านเข้า Supabase สำเร็จ จำนวน ${maidPayload.length} รายการ`);
    }
  }

  // ----------------------------------------
  // 2. ดึงและส่งข้อมูลตั้งค่า (Config_Cases)
  // ----------------------------------------
  const configSheet = ss.getSheetByName("Config_Cases");
  if (configSheet && configSheet.getLastRow() > 1) {
    Logger.log("เริ่มดึงข้อมูลตั้งค่าระบบ...");
    const configData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 5).getDisplayValues();
    const configPayload = [];
    
    configData.forEach(r => {
      // ผู้รับผิดชอบ (Admin)
      if (String(r[0]).trim() !== "") {
        configPayload.push({
          config_type: 'admin',
          value1: String(r[0]).trim(),
          value2: String(r[1] || "").trim()
        });
      }
      // หัวข้อ (Topic)
      if (String(r[2]).trim() !== "") {
        configPayload.push({
          config_type: 'topic',
          value1: String(r[2]).trim(),
          value2: String(r[3] || "").trim()
        });
      }
      // สิทธิ์แอดมิน (Super Admin)
      if (String(r[4]).trim() !== "") {
        configPayload.push({
          config_type: 'super_admin',
          value1: String(r[4]).trim(),
          value2: ""
        });
      }
    });

    if (configPayload.length > 0) {
      const options = {
        method: 'post',
        headers: {
          'apikey': MIG_KEY,
          'Authorization': 'Bearer ' + MIG_KEY,
          'Content-Type': 'application/json',
          'Prefer': 'resolution=minimal'
        },
        payload: JSON.stringify(configPayload),
        muteHttpExceptions: true
      };
      
      const res = UrlFetchApp.fetch(MIG_URL + '/rest/v1/Config_System', options);
      if(res.getResponseCode() === 201) {
          Logger.log(`✅ ส่งข้อมูล Config เข้า Supabase สำเร็จ จำนวน ${configPayload.length} รายการ`);
      } else {
          Logger.log("❌ ข้อผิดพลาด Config: " + res.getContentText());
      }
    }
  }

  try {
      SpreadsheetApp.getUi().alert("🎉 ย้ายข้อมูล แม่บ้าน & Config ไปที่ Supabase เสร็จสมบูรณ์!");
  } catch(e) {
      Logger.log("การทำงานเสร็จสิ้นสมบูรณ์");
  }
}
