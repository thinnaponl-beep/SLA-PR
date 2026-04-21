
// --- Supabase Config ---
const SUPABASE_URL = 'https://apsnkqgrkjtfsiydjcfd.supabase.co'; 
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFwc25rcWdya2p0ZnNpeWRqY2ZkIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NjMwNjcwMywiZXhwIjoyMDkxODgyNzAzfQ.xlnbmC9rTuvoa8vGmvi1Oufik-zxifQpBrYL9sXAySI';

/**
 * 1. Entry Point: จัดการ Routing เพื่อแสดงหน้า Case, Config หรือ Dashboard
 */
function doGet(e) {
  let page = e.parameter.page || 'case'; 
  let templateName = 'Case'; // Default

  if (page === 'config_case') {
    if (!checkIsSuperAdmin(Session.getActiveUser().getEmail())) {
       return HtmlService.createHtmlOutput('<h3 style="font-family:sans-serif; text-align:center; margin-top:50px; color:#dc3545;">⛔ Access Denied: คุณไม่มีสิทธิ์เข้าถึงหน้านี้</h3>');
    }
    templateName = 'Config_Case';
  } else if (page === 'dashboard') {
    templateName = 'Dashboard';
  }

  let template = HtmlService.createTemplateFromFile(templateName);
  template.url = getScriptUrl(); 

  return template.evaluate()
    .setTitle('Case Management Tracking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. Helper: ดึง URL ของ Web App สำหรับใช้เปลี่ยนหน้า
 */
function getScriptUrl() { 
  return ScriptApp.getService().getUrl(); 
}

/**
 * Helper: ตรวจสอบสิทธิ์ Super Admin ภายใน Server Script (เปลี่ยนมาดึงจาก Supabase)
 */
function checkIsSuperAdmin(email) {
  // 🌟 แก้ไขชื่อตารางเป็น Config_System
  const res = supabaseRequest(`Config_System?config_type=eq.super_admin&value1=eq.${encodeURIComponent(email)}&select=value1`, 'GET');
  if (res && !res.error && Array.isArray(res) && res.length > 0) {
    return true;
  }
  return false;
}

// ------------------------------------------
// --- ฟังก์ชันสำหรับทดสอบการดึงข้อมูล Config ---
// ------------------------------------------
function testFetchConfigs() {
  Logger.log("กำลังทดสอบดึงข้อมูลจากตาราง Config_System...");
  // 🌟 แก้ไขชื่อตารางเป็น Config_System
  const res = supabaseRequest('Config_System?select=*&limit=100&order=id.asc', 'GET');
  
  if (!res) {
    Logger.log("❌ ไม่ได้ข้อมูลกลับมาเลย (ผลลัพธ์เป็น null) กรุณาเช็ค URL หรือ KEY");
  } else if (res.error) {
    Logger.log("❌ เกิด Error จาก Supabase: " + JSON.stringify(res.error));
  } else {
    Logger.log("✅ ดึงข้อมูลสำเร็จ! พบทั้งหมด " + res.length + " รายการ");
    if (res.length > 0) {
      Logger.log("ตัวอย่างข้อมูลรายการแรก: " + JSON.stringify(res[0]));
    }
  }
}

// ------------------------------------------
// --- SUPABASE API FUNCTIONS ---
// ------------------------------------------

function supabaseRequest(endpoint, method = 'GET', payload = null) {
  const cleanUrl = SUPABASE_URL.trim().replace(/\/$/, "");
  const options = {
    method: method,
    headers: {
      'apikey': SUPABASE_KEY.trim(),
      'Authorization': 'Bearer ' + SUPABASE_KEY.trim(),
      'Content-Type': 'application/json',
      'Prefer': 'return=representation'
    },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  
  const response = UrlFetchApp.fetch(cleanUrl + '/rest/v1/' + endpoint, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  // 🌟 ดักจับ Error: ถ้ารหัสตั้งแต่ 400 ขึ้นไป (เช่น บันทึกไม่เข้า) ให้เตะ Error กลับไปโชว์ที่หน้าจอ
  if (code >= 400) {
    Logger.log(`🚨 Supabase API Error [${code}]: ${text}`);
    let errMsg = text;
    try { errMsg = JSON.parse(text).message || text; } catch(e){}
    throw new Error(`DB Error: ${errMsg}`);
  }

  if (!text || text.trim() === "") return true; // กรณีสำเร็จแต่ไม่มี data ส่งกลับมา
  
  try {
    return JSON.parse(text);
  } catch (e) {
    return text;
  }
}

/**
 * ฟังก์ชันช่วยเหลือ: แปลงวันที่จาก DB ให้เป็นรูปแบบไทย
 */
function formatDbDateToTh(dbDateStr) {
  if (!dbDateStr) return "";
  const str = dbDateStr.replace('T', ' ').split('+')[0].split('.')[0]; 
  const parts = str.split(' ');
  if (parts.length !== 2) return dbDateStr;
  const ymd = parts[0].split('-');
  if (ymd.length === 3) {
    return `${ymd[2]}/${ymd[1]}/${ymd[0]} ${parts[1]}`;
  }
  return dbDateStr;
}

/**
 * ดึงข้อมูล Case ทั้งหมดแบบ Pagination
 */
function fetchCasesAsArray() {
   Logger.log("กำลังดึงข้อมูล Case ทั้งหมดจาก Supabase...");
   let allCases = [];
   let offset = 0;
   const limit = 1000;
   let hasMore = true;

   while (hasMore) {
       try {
           const endpoint = `Database_Cases?select=*&limit=${limit}&offset=${offset}&order=%22Case%20ID%22.desc`;
           const res = supabaseRequest(endpoint, 'GET');
           
           if (res && Array.isArray(res) && res.length > 0) {
               allCases = allCases.concat(res);
               if (res.length < limit) hasMore = false; 
               else offset += limit; 
           } else {
               hasMore = false;
           }
       } catch (e) {
           Logger.log("Fetch Cases Error: " + e);
           hasMore = false;
       }
   }

   // 🌟 แปลงวันที่จากฐานข้อมูล (YYYY-MM-DD) เป็นของหน้าจอ (DD/MM/YYYY)
   return allCases.map(r => [
       r['Case ID'] || "", r['Status'] || "", 
       formatDbDateToTh(r['Time_Created']), 
       formatDbDateToTh(r['Time_Accepted']), 
       formatDbDateToTh(r['Time_Closed']),
       r['Creator'] || "", r['Assignee'] || "", r['Maid ID'] || "", r['Maid Name'] || "",
       typeof r['Topic'] === 'string' ? r['Topic'] : JSON.stringify(r['Topic'] || {}),
       r['Chat Link'] || "", r['Action Details'] || "",
       typeof r['History Logs'] === 'string' ? r['History Logs'] : JSON.stringify(r['History Logs'] || [])
   ]);
}

/**
 * ดึงข้อมูลแม่บ้านทั้งหมด (Pagination)
 */
function getAllMaids() {
  Logger.log("กำลังดึงข้อมูลแม่บ้านทั้งหมด...");
  let allMaids = [];
  let offset = 0;
  const limit = 1000;
  let hasMore = true;

  while (hasMore) {
    try {
      const res = supabaseRequest(`Database_Maids?select=maid_id,maid_name&limit=${limit}&offset=${offset}&order=maid_id.asc`, 'GET');
      if (res && Array.isArray(res) && res.length > 0) {
        allMaids = allMaids.concat(res);
        if (res.length < limit) hasMore = false;
        else offset += limit;
      } else {
        hasMore = false;
      }
    } catch (e) {
      hasMore = false;
    }
  }

  return allMaids.map(r => ({ id: String(r.maid_id).trim(), name: String(r.maid_name).trim() })).filter(item => item.id !== "");
}

function getCaseHistory(caseId) {
   try {
     const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}&select=%22History%20Logs%22`, 'GET');
     if (res && Array.isArray(res) && res.length > 0) {
         let logs = res[0]['History Logs'];
         if (typeof logs === 'string') {
             try { return JSON.parse(logs); } catch(e) { return []; }
         }
         return logs || [];
     }
   } catch(e) { }
   return [];
}

// ------------------------------------------
// --- MAID DATABASE FUNCTIONS (Supabase) ---
// ------------------------------------------

function getAllMaids() {
  console.log("กำลังดึงข้อมูลแม่บ้านจาก Supabase...");
  
  let allMaids = [];
  let offset = 0;
  const limit = 1000;
  let hasMore = true;

  while (hasMore) {
    const res = supabaseRequest(`Database_Maids?select=maid_id,maid_name&limit=${limit}&offset=${offset}&order=maid_id.asc`, 'GET');
    
    if (res && !res.error && Array.isArray(res)) {
      allMaids = allMaids.concat(res);
      if (res.length < limit) {
        hasMore = false;
      } else {
        offset += limit;
      }
    } else {
      hasMore = false;
    }
  }

  if (allMaids.length > 0) {
    console.log(`✅ ดึงข้อมูลแม่บ้านสำเร็จ รวมได้ทั้งหมด ${allMaids.length} คน`);
    return allMaids.map(r => ({ id: String(r.maid_id).trim(), name: String(r.maid_name).trim() })).filter(item => item.id !== "");
  }
  
  return [];
}

function searchMaidById(maidId) {
  if (!maidId) return { success: false, message: "กรุณาระบุรหัส" };
  const res = supabaseRequest(`Database_Maids?maid_id=eq.${encodeURIComponent(maidId)}&select=*`, 'GET');
  if (res && !res.error && Array.isArray(res) && res.length > 0) {
    return { success: true, found: true, name: res[0].maid_name };
  }
  return { success: true, found: false };
}

function addNewMaid(maidId, maidName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const check = supabaseRequest(`Database_Maids?maid_id=eq.${encodeURIComponent(maidId)}&select=maid_id`, 'GET');
    if (check && !check.error && Array.isArray(check) && check.length > 0) {
      return { success: false, message: "รหัสแม่บ้านนี้มีอยู่ในระบบแล้ว" };
    }
    
    const payload = {
      "maid_id": maidId,
      "maid_name": maidName,
      "note": "Added via Case Form"
    };
    
    const res = supabaseRequest('Database_Maids', 'POST', payload);
    if (res && res.error) throw new Error(res.error.message);
    
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// ------------------------------------------
// --- CONFIG DATA FUNCTIONS (Supabase) ---
// ------------------------------------------

function getCaseConfigs() {
  // 🌟 แก้ไขชื่อตารางเป็น Config_System
  const res = supabaseRequest('Config_System?select=*&limit=1000&order=id.asc', 'GET');
  
  let admins = []; 
  let topicData = []; 
  let adminDetails = [];
  let superAdmins = [];

  if (res && !res.error && Array.isArray(res)) {
      res.forEach(row => {
          if (row.config_type === 'admin') {
              if (row.value1) {
                  admins.push(row.value2 || row.value1); 
                  adminDetails.push({ name: row.value1, email: row.value2 || "" });
              }
          } else if (row.config_type === 'topic') {
              if (row.value1) {
                  topicData.push({
                      main: row.value1,
                      subs: row.value2 ? row.value2.split(',').map(s => s.trim()).filter(s => s !== "") : []
                  });
              }
          } else if (row.config_type === 'super_admin') {
              if (row.value1) {
                  superAdmins.push(row.value1);
              }
          }
      });
  }
  
  return { 
      admins: admins, 
      topicData: topicData,
      adminDetails: adminDetails,
      superAdmins: superAdmins
  };
}

function addCaseConfigItem(type, value, subValue) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    let payload = { config_type: type };
    if (type === 'admin') {
      let parts = value.split('|');
      payload.value1 = parts[0].trim();
      payload.value2 = parts.length > 1 ? parts[1].trim() : "";
    } else if (type === 'topic') {
      payload.value1 = value;
      payload.value2 = subValue || "";
    } else if (type === 'super_admin') {
      payload.value1 = value;
      payload.value2 = "";
    }
    
    // 🌟 แก้ไขชื่อตารางเป็น Config_System
    const res = supabaseRequest('Config_System', 'POST', payload);
    if (res && res.error) throw new Error(res.error.message);
    
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function updateCaseConfigTopic(oldName, newName, newSubs) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const payload = {
        "value1": newName,
        "value2": newSubs || ""
    };
    // 🌟 แก้ไขชื่อตารางเป็น Config_System
    const res = supabaseRequest(`Config_System?config_type=eq.topic&value1=eq.${encodeURIComponent(oldName)}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);
    
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function removeCaseConfigItem(type, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    // 🌟 แก้ไขชื่อตารางเป็น Config_System
    const res = supabaseRequest(`Config_System?config_type=eq.${type}&value1=eq.${encodeURIComponent(value)}`, 'DELETE');
    if (res && res.error) throw new Error(res.error.message);
    
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// ------------------------------------------
// --- CORE DATA FUNCTIONS (Main Case) ---
// ------------------------------------------

function getCasesData() {
  const currentUser = Session.getActiveUser().getEmail();
  const canConfig = checkIsSuperAdmin(currentUser);

  let data = [];
  let stats = { total: 0, pending: 0, progress: 0, closed: 0 };

  const values = fetchCasesAsArray();

  if (values.length > 0) {
    data = values.map((row, index) => {
      const status = row[1].trim();
      stats.total++;
      if (status === 'รอการประสานงาน') stats.pending++;
      else if (status === 'กำลังประสานงาน') stats.progress++;
      else if (status === 'ปิดเคส') stats.closed++;
      
      let systemCreatedTime = row[2];
      try {
        const logs = JSON.parse(row[12] || "[]");
        const createLog = logs.find(l => l.action === 'Create');
        if (createLog && createLog.timestamp) systemCreatedTime = createLog.timestamp;
      } catch(e){}

      return {
        rowIndex: index + 2, id: row[0], status: row[1],
        timeCreated: row[2], timeAccepted: row[3], timeClosed: row[4],
        creator: row[5], assignee: row[6],
        maidId: row[7], maidName: row[8],
        topic: row[9],
        chatLink: row[10], actionDetails: row[11],
        historyLogs: row[12],
        systemCreatedTime: systemCreatedTime
      };
    });
  }
  
  stats.pendingPct = stats.total > 0 ? ((stats.pending / stats.total) * 100).toFixed(1) : 0;
  stats.progressPct = stats.total > 0 ? ((stats.progress / stats.total) * 100).toFixed(1) : 0;
  stats.closedPct = stats.total > 0 ? ((stats.closed / stats.total) * 100).toFixed(1) : 0;
  
  return { 
      currentUser: currentUser, 
      canConfig: canConfig,
      data: data, 
      stats: stats, 
      configs: getCaseConfigs() 
  };
}

function createNewCase(form) {
  try {
    const user = Session.getActiveUser().getEmail();
    
    let createdTimeStr;
    if (form.caseDate && form.caseTime) {
      const [formYear, formMonth, formDay] = form.caseDate.split('-');
      createdTimeStr = `${formYear}-${formMonth}-${formDay} ${form.caseTime}:00`;
    } else {
      createdTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    }
    
    const year = new Date().getFullYear().toString().substr(-2);
    const prefix = `CASE-${year}-`;
    let maxSeq = 0;
    
    const latestCase = supabaseRequest(`Database_Cases?select=%22Case%20ID%22&%22Case%20ID%22=like.${prefix}*&order=%22Case%20ID%22.desc&limit=1`, 'GET');
    if (latestCase && latestCase.length > 0 && latestCase[0]['Case ID']) {
        maxSeq = parseInt(latestCase[0]['Case ID'].split('-')[2]);
    }
    const caseId = `${prefix}${String(maxSeq + 1).padStart(4, '0')}`;

    const newLog = {
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      action: "Create",
      user: user,
      details: "เปิดเคสใหม่"
    };

    const payload = {
        "Case ID": caseId,
        "Status": "รอการประสานงาน",
        "Time_Created": createdTimeStr,
        "Time_Accepted": null,
        "Time_Closed": null,
        "Creator": user,
        "Assignee": form.assignee,
        "Maid ID": form.maidId,
        "Maid Name": form.maidName,
        "Topic": JSON.stringify(form.topicObj || {}), 
        "Chat Link": form.chatLink,
        "Action Details": "",
        "History Logs": JSON.stringify([newLog]) 
    };

    const res = supabaseRequest('Database_Cases', 'POST', payload);
    if (!res || res.error) {
  throw new Error("Insert failed: " + JSON.stringify(res));
}
    
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function acceptCase(caseId) {
  try {
    console.log(`กำลังรับเคส: ${caseId}`);
    const user = Session.getActiveUser().getEmail();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    const history = getCaseHistory(caseId);
    history.push({ timestamp: now, action: "Accept", user: user, details: "รับเคส" });

    const payload = {
        "Status": "กำลังประสานงาน",
        "Time_Accepted": now,
        "History Logs": JSON.stringify(history)
    };

    const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);
    
    return { success: true };
  } catch (e) { 
    console.error("🚨 Error (acceptCase):", e.toString());
    return { success: false, message: e.toString() }; 
  }
}

function updateActionDetails(caseId, details) {
  try {
    const payload = { "Action Details": details };
    const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);
    return { success: true };
  } catch (e) { 
    console.error("🚨 Error (updateActionDetails):", e.toString());
    return { success: false }; 
  }
}

function updateCaseInfo(form) {
  try {
    console.log(`กำลังอัปเดตข้อมูลเคส: ${form.id}`);
    const user = Session.getActiveUser().getEmail();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    const history = getCaseHistory(form.id);
    history.push({ timestamp: now, action: "Edit", user: user, details: "แก้ไขข้อมูลเคส" });

    const payload = {
        "Maid ID": form.maidId,
        "Maid Name": form.maidName,
        "Topic": JSON.stringify(form.topicObj || {}),
        "Chat Link": form.chatLink,
        "History Logs": JSON.stringify(history)
    };

    const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${form.id}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);

    return { success: true };
  } catch (e) { 
    console.error("🚨 Error (updateCaseInfo):", e.toString());
    return { success: false, message: e.toString() }; 
  }
}

function closeCase(caseId, details) {
  try {
    console.log(`กำลังปิดเคส: ${caseId}`);
    const user = Session.getActiveUser().getEmail();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    const history = getCaseHistory(caseId);
    history.push({ timestamp: now, action: "Close", user: user, details: "ปิดเคส: " + details.substring(0, 30) + "..." });

    const payload = {
        "Status": "ปิดเคส",
        "Time_Closed": now,
        "Action Details": details,
        "History Logs": JSON.stringify(history)
    };

    const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);

    return { success: true };
  } catch (e) { 
    console.error("🚨 Error (closeCase):", e.toString());
    return { success: false }; 
  }
}

function reassignCase(caseId, newAssignee, note) {
  try {
    console.log(`กำลังส่งต่อเคส: ${caseId}`);
    const user = Session.getActiveUser().getEmail();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    const currentCase = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}&select=%22Assignee%22,%22History%20Logs%22`, 'GET');
    if (!currentCase || currentCase.length === 0) throw new Error("Case Not Found");

    const oldAssignee = currentCase[0]['Assignee'];
    let history = currentCase[0]['History Logs'];
    if (typeof history === 'string') {
        try { history = JSON.parse(history); } catch(e) { history = []; }
    }
    history = history || [];

    const noteText = note ? ` (Note: ${note})` : '';
    history.push({ timestamp: now, action: "Reassign", user: user, details: `ส่งต่อ: ${oldAssignee} -> ${newAssignee}${noteText}` });

    const payload = {
        "Assignee": newAssignee,
        "History Logs": JSON.stringify(history)
    };

    const res = supabaseRequest(`Database_Cases?%22Case%20ID%22=eq.${caseId}`, 'PATCH', payload);
    if (res && res.error) throw new Error(res.error.message);

    return { success: true };
  } catch (e) { 
    console.error("🚨 Error (reassignCase):", e.toString());
    return { success: false }; 
  }
}

// ------------------------------------------
// --- DASHBOARD & KPI FUNCTIONS ---
// ------------------------------------------

function getTopicCases(dateStr, topicName) {
  const parseDate = (str) => {
    if (!str) return null;
    try {
      const parts = str.split(' ');
      if (parts.length >= 2) {
        const dmy = parts[0].split('/');
        const hms = parts[1].split(':');
        let year = parseInt(dmy[2]); if (year > 2400) year -= 543;
        return new Date(year, parseInt(dmy[1])-1, parseInt(dmy[0]), parseInt(hms[0]), parseInt(hms[1])).getTime();
      }
      return null;
    } catch (e) { return null; }
  };

  const data = fetchCasesAsArray();
  if (data.length === 0) return [];
  
  let targetDateStr = dateStr; 
  
  const result = data.filter(row => {
      let systemCreatedStr = row[2];
      try {
        const logs = JSON.parse(row[12] || "[]");
        const createLog = logs.find(l => l.action === 'Create');
        if (createLog && createLog.timestamp) systemCreatedStr = createLog.timestamp;
      } catch(e){}
      
      const ts = parseDate(systemCreatedStr);
      if (!ts) return false;
      
      const d = new Date(ts);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      const rowDateKey = `${yyyy}-${mm}-${dd}`;
      
      if (rowDateKey !== targetDateStr) return false;

      let mainTopic = "ไม่ระบุ";
      try {
        const tObj = JSON.parse(row[9]);
        if(tObj && tObj.main) mainTopic = tObj.main;
      } catch(e) { if(row[9]) mainTopic = row[9]; }
      
      return mainTopic === topicName;
  }).map(row => {
      let subTopicsStr = "";
      try {
        const tObj = JSON.parse(row[9]);
        if(tObj && tObj.sub && tObj.sub.length > 0) subTopicsStr = tObj.sub.join(", ");
      } catch(e) {}
      
      return {
          id: row[0],
          time: row[2].split(' ')[1].substring(0,5), 
          maid: row[8], 
          assignee: row[6],
          status: row[1],
          subTopic: subTopicsStr, 
          lastAction: row[11] || "-" 
      };
  });
  
  return result;
}

function getDashboardStats(viewMode, filterValue, filterAssignee) {
  const emptyResult = { 
    avgOpenAccept: '-', 
    avgAcceptClose: '-', 
    assigneeStats: [],
    filterOptions: { months: [], assignees: [] },
    comparisonData: null,
    hourlyStats: null,
    topicStats: []
  };

  const data = fetchCasesAsArray();
  if (data.length === 0) return emptyResult;
  
  let totalOpenAccept = 0, countOpenAccept = 0;
  let totalAcceptClose = 0, countAcceptClose = 0;
  let assigneeMap = {};
  let topicCounts = {}; 
  let hourlyDataRaw = Array(24).fill(null).map(() => ({ count: 0, totalResponseTime: 0, responseCount: 0 }));

  const nowObj = new Date();
  const currentMonthKey = `${nowObj.getFullYear()}-${String(nowObj.getMonth() + 1).padStart(2, '0')}`;
  
  let comparisonData = { reference: { total: 0, closed: 0 }, selected: { total: 0, closed: 0 } };
  let availableMonths = new Set();
  let availableAssignees = new Set();

  const parseDate = (str) => {
    if (!str || str === "" || str === "-") return null;
    try {
      const parts = str.split(' ');
      if (parts.length >= 2) {
        const dmy = parts[0].split('/');
        const hms = parts[1].split(':');
        if (dmy.length === 3 && hms.length >= 2) {
          let year = parseInt(dmy[2]);
          if (year > 2400) year -= 543;
          const month = parseInt(dmy[1]) - 1;
          const day = parseInt(dmy[0]);
          return new Date(year, month, day, parseInt(hms[0]), parseInt(hms[1])).getTime();
        }
      }
      return new Date(str).getTime();
    } catch (e) { return null; }
  };

  const formatDuration = (ms) => {
    if (!ms || isNaN(ms) || ms <= 0) return '-';
    const totalMinutes = Math.floor(ms / 60000);
    const d = Math.floor(totalMinutes / (24 * 60));
    const h = Math.floor((totalMinutes % (24 * 60)) / 60);
    const m = totalMinutes % 60;
    let result = [];
    if (d > 0) result.push(`${d} วัน`);
    if (h > 0) result.push(`${h} ชม.`);
    result.push(`${m} นาที`);
    return result.join(' ');
  };

  const configs = getCaseConfigs();
  let adminNames = {};
  if (configs.adminDetails && configs.adminDetails.length > 0) {
      configs.adminDetails.forEach(a => { if(a.email) adminNames[a.email] = a.name; });
  }

  let refValue = ""; 
  if (viewMode === 'day' && filterValue) {
      const selectedDate = new Date(filterValue);
      selectedDate.setDate(selectedDate.getDate() - 1); 
      refValue = Utilities.formatDate(selectedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
      refValue = currentMonthKey;
  }

  data.forEach(row => {
    const rawCreatedEffective = row[2];
    const assignee = row[6] ? row[6].trim() : 'Unassigned';
    const status = row[1];
    
    let systemCreatedStr = rawCreatedEffective;
    try {
        const logs = JSON.parse(row[12] || "[]");
        const createLog = logs.find(l => l.action === 'Create');
        if (createLog && createLog.timestamp) systemCreatedStr = createLog.timestamp;
    } catch (e) {}

    const tCreatedSystem = parseDate(systemCreatedStr);       
    let rowMonthKey = ''; let rowDateKey = '';  
    
    if (tCreatedSystem) {
        const d = new Date(tCreatedSystem);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        rowMonthKey = `${yyyy}-${mm}`;
        rowDateKey = `${yyyy}-${mm}-${dd}`;
        availableMonths.add(rowMonthKey);
    }

    if (assignee !== 'Unassigned') {
        const displayName = adminNames[assignee] ? `${adminNames[assignee]} (${assignee})` : assignee;
        availableAssignees.add(JSON.stringify({email: assignee, label: displayName}));
    }

    const isTargetAssignee = (!filterAssignee || filterAssignee === 'all' || assignee === filterAssignee);
    if (isTargetAssignee) {
        let matchRef = false;
        if (viewMode === 'day' && rowDateKey === refValue) matchRef = true;
        if (viewMode === 'month' && rowMonthKey === refValue) matchRef = true;
        if (matchRef) {
            comparisonData.reference.total++;
            if (status === 'ปิดเคส') comparisonData.reference.closed++;
        }
        let matchSelected = false;
        if (viewMode === 'day' && rowDateKey === filterValue) matchSelected = true;
        if (viewMode === 'month' && rowMonthKey === filterValue) matchSelected = true;
        if (matchSelected) {
            comparisonData.selected.total++;
            if (status === 'ปิดเคส') comparisonData.selected.closed++;
        }
    }

    let includeRow = true;
    if (viewMode === 'day' && filterValue && rowDateKey !== filterValue) includeRow = false;
    if (viewMode === 'month' && filterValue && filterValue !== 'all' && rowMonthKey !== filterValue) includeRow = false;
    if (filterAssignee && filterAssignee !== 'all' && assignee !== filterAssignee) includeRow = false;

    if (!includeRow) return;

    let mainTopic = "ไม่ระบุ";
    try {
        const tObj = JSON.parse(row[9]); 
        if(tObj && tObj.main) mainTopic = tObj.main;
    } catch(e) {
        if(row[9] && row[9] !== "") mainTopic = row[9];
    }
    if(!topicCounts[mainTopic]) topicCounts[mainTopic] = 0;
    topicCounts[mainTopic]++;

    const tAccepted = parseDate(row[3]);
    const tClosed = parseDate(row[4]);

    if (viewMode === 'day' && tCreatedSystem) {
        const hour = new Date(tCreatedSystem).getHours();
        if (hour >= 0 && hour < 24) {
            hourlyDataRaw[hour].count++;
            if (tAccepted && tCreatedSystem && tAccepted >= tCreatedSystem) {
                hourlyDataRaw[hour].totalResponseTime += (tAccepted - tCreatedSystem);
                hourlyDataRaw[hour].responseCount++;
            }
        }
    }

    if (tCreatedSystem && tAccepted && tAccepted >= tCreatedSystem) {
      totalOpenAccept += (tAccepted - tCreatedSystem);
      countOpenAccept++;
    }
    if (tAccepted && tClosed && tClosed >= tAccepted) {
      totalAcceptClose += (tClosed - tAccepted);
      countAcceptClose++;
    }

    if (!assigneeMap[assignee]) {
      assigneeMap[assignee] = { email: assignee, total: 0, pending: 0, progress: 0, closed: 0, totalHandleTime: 0, closedWithTime: 0 };
    }
    assigneeMap[assignee].total++;
    if (status === 'รอการประสานงาน') assigneeMap[assignee].pending++;
    else if (status === 'กำลังประสานงาน') assigneeMap[assignee].progress++;
    else if (status === 'ปิดเคส') {
      assigneeMap[assignee].closed++;
      if (tCreatedSystem && tClosed && tClosed >= tCreatedSystem) {
        assigneeMap[assignee].totalHandleTime += (tClosed - tCreatedSystem);
        assigneeMap[assignee].closedWithTime++;
      }
    }
  });

  const avgOA = countOpenAccept > 0 ? formatDuration(totalOpenAccept / countOpenAccept) : '-';
  const avgAC = countAcceptClose > 0 ? formatDuration(totalAcceptClose / countAcceptClose) : '-';

  const assigneeStats = Object.values(assigneeMap).map(a => {
    const avgTime = a.closedWithTime > 0 ? formatDuration(a.totalHandleTime / a.closedWithTime) : '-';
    const displayName = adminNames[a.email] || a.email.split('@')[0];
    return { name: displayName, email: a.email, total: a.total, pending: a.pending, progress: a.progress, closed: a.closed, avgTime: avgTime };
  });

  assigneeStats.sort((a, b) => b.total - a.total);
  const sortedMonths = Array.from(availableMonths).sort().reverse();
  const sortedAssignees = Array.from(availableAssignees).map(j => JSON.parse(j)).sort((a,b) => a.label.localeCompare(b.label));

  let hourlyStats = null;
  if (viewMode === 'day') {
      hourlyStats = hourlyDataRaw.map((h, i) => ({
          hour: `${String(i).padStart(2, '0')}:00`,
          count: h.count,
          avgResponse: h.responseCount > 0 ? Math.round((h.totalResponseTime / h.responseCount) / 60000) : 0 
      }));
  }

  const topicStats = Object.keys(topicCounts).map(key => ({ label: key, count: topicCounts[key] })).sort((a,b) => b.count - a.count);

  return {
    avgOpenAccept: avgOA, avgAcceptClose: avgAC,
    assigneeStats: assigneeStats,
    filterOptions: { months: sortedMonths, assignees: sortedAssignees },
    comparisonData: comparisonData,
    hourlyStats: hourlyStats,
    topicStats: topicStats,
    refLabel: viewMode === 'day' ? 'เมื่อวาน' : 'เดือนปัจจุบัน'
  };
}

function getDashboardStatsAndCases(viewMode, filterValue, filterAssignee) {
  const stats = getDashboardStats(viewMode, filterValue, filterAssignee);
  let cases = [];
  
  const data = fetchCasesAsArray();
  
  if (data.length > 0) {
    const parseDate = (str) => {
      if (!str || str === "" || str === "-") return null;
      try {
        const parts = str.split(' ');
        if (parts.length >= 2) {
          const dmy = parts[0].split('/');
          const hms = parts[1].split(':');
          if (dmy.length === 3 && hms.length >= 2) {
            let year = parseInt(dmy[2]);
            if (year > 2400) year -= 543;
            const month = parseInt(dmy[1]) - 1;
            const day = parseInt(dmy[0]);
            return new Date(year, month, day, parseInt(hms[0]), parseInt(hms[1])).getTime();
          }
        }
        return new Date(str).getTime();
      } catch (e) { return null; }
    };

    cases = data.filter(row => {
      const rawCreatedEffective = row[2];
      const assignee = row[6] ? row[6].trim() : 'Unassigned';
      
      let systemCreatedStr = rawCreatedEffective;
      try {
          const logs = JSON.parse(row[12] || "[]");
          const createLog = logs.find(l => l.action === 'Create');
          if (createLog && createLog.timestamp) systemCreatedStr = createLog.timestamp;
      } catch (e) {}

      const tCreatedSystem = parseDate(systemCreatedStr);
      if (!tCreatedSystem) return false;

      const d = new Date(tCreatedSystem);
      const rowMonthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      const rowDateKey = `${rowMonthKey}-${String(d.getDate()).padStart(2, '0')}`;

      let includeRow = true;
      if (viewMode === 'day' && filterValue && rowDateKey !== filterValue) includeRow = false;
      if (viewMode === 'month' && filterValue && filterValue !== 'all' && rowMonthKey !== filterValue) includeRow = false;
      if (filterAssignee && filterAssignee !== 'all' && assignee !== filterAssignee) includeRow = false;
      
      return includeRow;
    }).map(row => ({
      id: row[0], status: row[1],
      assignee: row[6], topic: row[9],
      actionDetails: row[11]
    }));
  }
  
  return { stats: stats, cases: cases };
}

// ------------------------------------------
// --- AI INTEGRATION (GEMINI) ---
// ------------------------------------------

function analyzeTopicCasesWithAI(promptData) {
  const API_KEY = 'AIzaSyAb8ZQdZPeuajdIOtWQlPtO0RV3_H-G8iQ'; 
  
  if (!API_KEY || API_KEY === 'YOUR_API_KEY_HERE') {
    return { success: false, message: 'กรุณาตั้งค่า API Key ของ Gemini ในไฟล์ Code.gs ก่อนใช้งานครับ' };
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
  
  const systemPrompt = `คุณคือระบบ AI อัจฉริยะสำหรับวิเคราะห์ข้อมูล Customer Success และประเมินคุณภาพการให้บริการ (QA) สำหรับระบบแอดมินที่ดูแลคุณแม่บ้าน
หน้าที่ของคุณคือการอ่านข้อมูลเคส โดยต้อง **โฟกัสการวิเคราะห์ไปที่ข้อมูลในคอลัมน์ "สิ่งที่จนท.ทำ" (การดำเนินการล่าสุด)** และ **วิเคราะห์การเลือก "หัวข้อย่อย"** อย่างละเอียด

กรุณาวิเคราะห์และสรุปผลตามหัวข้อต่อไปนี้:

1. 🕵️‍♂️ ประเมินคุณภาพการลงข้อมูลและการแก้ปัญหา (Action Quality Analysis):
   - จากข้อความที่แอดมินบันทึกไว้ พวกเขามีวิธีการจัดการปัญหาของคุณแม่บ้านอย่างไร? การลงบันทึกชัดเจนและครบถ้วนไหม?
   - เป็นการแก้ปัญหาจบในครั้งเดียว (First Contact Resolution) หรือแค่รับเรื่อง/ผลัดผ่อน/ประสานงานหลายทอดไปมา?
   - มีจุดไหนที่ดูล่าช้า (Bottleneck) หรือทำซ้ำซ้อนจากสิ่งที่แอดมินพิมพ์ไว้ไหม?

2. 🔍 วิเคราะห์สาเหตุรากฐาน (Root Cause Insights):
   - ปัญหาอะไรของคุณแม่บ้านที่โผล่มาบ่อยที่สุด และคิดว่าสาเหตุลึกๆ ของเรื่องพวกนี้น่าจะมาจากอะไร?

3. 🗂️ วิเคราะห์การจัดหมวดหมู่ "อื่นๆ" (Category Matching):
   - ตรวจสอบเคสที่ระบุหัวข้อย่อยว่า "อื่นๆ: [ข้อความ]" ให้วิเคราะห์ว่าสิ่งที่แอดมินพิมพ์มานั้น แท้จริงแล้วสามารถจัดเข้ากลุ่มหัวข้อมาตรฐานที่มีอยู่แล้วตามตรรกะได้หรือไม่? 
   - หากพบว่ามีเรื่องใหม่ที่คนกรอกลงใน "อื่นๆ" ซ้ำกันบ่อยๆ ให้เสนอแนะการสร้างหัวข้อย่อยใหม่ (New Topic Suggestion)

4. ⚠️ จุดวิกฤตหรือความเสี่ยง (Red Flags & Risks):
   - มีเคสไหนที่สะท้อนถึงช่องโหว่ของระบบการดูแลคุณแม่บ้าน หรือดูแล้วเสี่ยงจะทำให้คุณแม่บ้านหรือลูกค้าหัวเสียหนักกว่าเดิมไหม?

5. 🛠️ ข้อเสนอแนะเชิงระบบและฟีเจอร์ (System & Feature Improvements):
   - จากปัญหาที่พบ เสนอแนะการสร้างหรือปรับปรุงระบบ/แพลตฟอร์ม เช่น ควรมีฟีเจอร์รูปแบบไหน หรือระบบอัตโนมัติแบบใด เพื่อแก้ปัญหาเหล่านี้ให้คุณแม่บ้านได้อย่างยั่งยืนในระยะยาว

6. 📝 คำแนะนำสำหรับการลงข้อมูลของแอดมิน (Data Entry Coaching):
   - แนะนำวิธีเขียนบันทึก 'การดำเนินการล่าสุด' ให้ดีขึ้น เพื่อให้สามารถนำข้อมูลไปวิเคราะห์หาสาเหตุ (Root Cause) ในระยะยาวได้ง่ายขึ้น
   - เสนอตัวอย่างการลงข้อมูลที่ดี เช่น ควรระบุข้อตกลงหรือเวลาที่ต้องติดตามผล
   - เสนอไอเดียปรับสเต็ปการทำงาน (SOP) ที่ช่วยให้ทีมจบงานได้ไวขึ้น

ข้อกำหนดรูปแบบการตอบ (STRICT FORMATTING): 
- **ห้าม** พิมพ์คำเกริ่นนำหรือคำทักทาย (เช่น ห้ามพิมพ์ "เรียน...", "สวัสดีครับ")
- **ห้าม** พิมพ์คำลงท้าย (เช่น ห้ามพิมพ์ "ขอแสดงความนับถือ", "สรุปได้ว่า...")
- บังคับให้ประโยคแรกสุดเริ่มต้นด้วยคำว่า "1. 🕵️‍♂️ ประเมินคุณภาพการลงข้อมูลและการแก้ปัญหา" ทันที
- **ใช้ภาษาแบบกึ่งทางการ สบายๆ อ่านง่าย เหมือนเพื่อนร่วมงานสรุปประเด็นให้ฟัง แต่ยังคงความน่าเชื่อถือ**
- ใช้ Bullet points จัดหน้าให้สวยงามและอ่านง่าย
- **บังคับ:** เวลาอ้างอิงเคสตัวอย่าง ให้ระบุเป็นรหัสเคส เช่น [CASE-26-0195] เสมอ ห้ามใช้คำว่า Case 1, เคสที่ 1 เด็ดขาด
- **บังคับ:** ต้องดึง "คำพูด" หรือ "ข้อความ" จากคอลัมน์ 'สิ่งที่จนท.ทำ' หรือข้อความที่กรอกใน 'อื่นๆ' มาเป็นตัวอย่างอ้างอิงประกอบการวิเคราะห์ให้เห็นภาพชัดเจน`;

  const payload = {
    "system_instruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{ "parts": [{ "text": "นี่คือข้อมูลเคสที่เกิดขึ้น:\n\n" + promptData }] }],
    "generationConfig": { "temperature": 0.2 }
  };

  const options = {
    'method': 'post', 'contentType': 'application/json',
    'payload': JSON.stringify(payload), 'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, message: 'API Error: ' + json.error.message };
    return { success: true, text: json.candidates[0].content.parts[0].text };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function analyzeCombinedStatsWithAI(promptData) {
  const API_KEY = 'AIzaSyAb8ZQdZPeuajdIOtWQlPtO0RV3_H-G8iQ'; 
  
  if (!API_KEY || API_KEY === 'YOUR_API_KEY_HERE') {
    return { success: false, message: 'กรุณาตั้งค่า API Key ของ Gemini ในไฟล์ Code.gs ก่อนใช้งานครับ' };
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
  
  const systemPrompt = `คุณคือระบบ AI อัจฉริยะสำหรับผู้บริหาร (Executive Dashboard Analyst) 
หน้าที่ของคุณคือการวิเคราะห์ 'ข้อมูลสถิติภาพรวม' ร่วมกับ 'ตัวอย่างบันทึกของแอดมิน' ที่ได้รับมา

กรุณาวิเคราะห์และสรุปผลตามหัวข้อต่อไปนี้:

1. 📊 สรุปประสิทธิภาพการทำงานของทีม (Team Performance & SLA):
   - ภาพรวมการตอบสนองและเวลาในการปิดเคสอยู่ในเกณฑ์ดีหรือไม่? 
   - มีใครที่ทำงานเร็วที่สุด หรือมีใครที่ดูเหมือนจะมีงานสะสมในมือเยอะเกินไปไหม? (อ้างอิงจากตัวเลขสถิติ)

2. 🕵️‍♂️ ประเมินคุณภาพการลงข้อมูลและการแก้ปัญหา (Action Quality Analysis):
   - จาก 'ตัวอย่างบันทึกการดำเนินการล่าสุด' แอดมินมีวิธีการจัดการปัญหาอย่างไร? อธิบายชัดเจนไหม? เป็นการแก้ปัญหาเบ็ดเสร็จหรือแค่ส่งเรื่องต่อ?
   - ดึง "คำพูด" หรือข้อความตัวอย่างที่แอดมินพิมพ์ มาเป็นหลักอ้างอิงในข้อนี้ให้ชัดเจน

3. 🗂️ วิเคราะห์สาเหตุปัญหา (Root Cause & Trends):
   - ปัญหาอะไรที่คุณแม่บ้านแจ้งเข้ามาบ่อยที่สุด สัดส่วนมากแค่ไหน?
   - จากข้อความตัวอย่าง มีเคสไหนที่ระบุ "อื่นๆ" แล้วน่าจะจัดเข้าหมวดหมู่ที่มีอยู่แล้วได้บ้าง?

4. 🛠️ ข้อเสนอแนะเชิงระบบและฟีเจอร์ (System & Feature Improvements):
   - เสนอแนะการสร้างฟีเจอร์ใหม่ ปรับ UX/UI หรือเพิ่มระบบอัตโนมัติบนแพลตฟอร์ม เพื่อช่วยลดปัญหาที่แม่บ้านเจอซ้ำๆ หรือช่วยให้แอดมินทำงานได้เร็วขึ้น

5. 💡 ข้อเสนอแนะเชิงกลยุทธ์และกระบวนการทำงาน (Strategic Recommendations):
   - เสนอแนวทาง 2-3 ข้อ เพื่อช่วยลดเวลาทำงาน กระจายงาน หรือปรับปรุงวิธีการเขียนบันทึกของทีมให้เป็นระบบมากขึ้น

ข้อกำหนดรูปแบบการตอบ (STRICT FORMATTING): 
- **ห้าม** พิมพ์คำเกริ่นนำหรือคำทักทาย (เช่น ห้ามพิมพ์ "เรียน...", "ขอแสดงความนับถือ")
- บังคับประโยคแรกเริ่มที่ "1. 📊 สรุปประสิทธิภาพการทำงานของทีม" ทันที
- ใช้ภาษาแบบกึ่งทางการ สบายๆ อ่านง่าย แต่อ้างอิงตัวเลข/ข้อมูลจริงเสนอเสมอ
- **บังคับ:** เวลายกตัวอย่างเคส ให้ระบุรหัสเคส เช่น [CASE-26-0195] เสมอ (ห้ามใช้คำว่า Case 1, 2)`;

  const payload = {
    "system_instruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{ "parts": [{ "text": promptData }] }],
    "generationConfig": { "temperature": 0.2 }
  };

  const options = {
    'method': 'post', 'contentType': 'application/json',
    'payload': JSON.stringify(payload), 'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, message: 'API Error: ' + json.error.message };
    return { success: true, text: json.candidates[0].content.parts[0].text };
  } catch (e) { return { success: false, message: e.toString() }; }
}
