// =======================================================
// ตั้งค่าระบบ (ตั้งค่าตรงนี้ให้ตรงกับของคุณ)
// =======================================================
const DRIVE_FOLDER_ID = '14G2JuUW4OQIL2ts2QdIHZ8-REYKDr5z4'; // <--- เปลี่ยนตรงนี้!

// =======================================================
// === ฟังก์ชัน Helper สำหรับแปลง Data เป็น TSV String (Hybrid Data Transfer) ===
function convertToTSV(dataArray, keys) {
  if (!dataArray || dataArray.length === 0) return '';
  const header = keys.join('\t');
  const rows = dataArray.map(obj => {
    return keys.map(k => {
      let val = (obj[k] !== undefined && obj[k] !== null) ? String(obj[k]) : '';
      return val.replace(/\t|\n|\r/g, ' '); // กันพัง: แทนที่ Tab/Newline ในข้อมูลด้วยช่องว่าง
    }).join('\t');
  });
  return [header, ...rows].join('\n');
}
// =======================================================

// =======================================================
// 0. Auto-Infrastructure (จัดการโฟลเดอร์และตาราง)
// =======================================================
function getDriveFolder() {
  try {
    if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID !== '14G2JuUW4OQIL2ts2QdIHZ8-REYKDr5z4') return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {}
  const folderName = "ระบบคลังเอกสารการสอน_Uploads";
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName); 
}

function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Email', 'Role', 'Name', 'Subject', 'ProfilePic', 'CoverPic', 'Password', 'IDCard', 'Phone']);
    userSheet.appendRow([Session.getActiveUser().getEmail() || 'admin@example.com', 'super_admin', 'ผู้อำนวยการระบบ', 'ส่วนกลาง', '', '', 'admin123', '0000000000000', '0000000000']);
  }
  
  let fileSheet = ss.getSheetByName('Files');
  if (!fileSheet) {
    fileSheet = ss.insertSheet('Files');
    fileSheet.appendRow(['FileID', 'UserEmail', 'FileName', 'FileType', 'Subject', 'FileURL', 'Note', 'Timestamp', 'Level', 'Term', 'Year']);
  } else {
    if(fileSheet.getMaxColumns() < 11) {
        fileSheet.insertColumnsAfter(fileSheet.getMaxColumns(), 11 - fileSheet.getMaxColumns());
    }
    if(fileSheet.getRange(1, 9).getValue() !== 'Level') {
        fileSheet.getRange(1, 9, 1, 3).setValues([['Level', 'Term', 'Year']]);
    }
  }

  let idCardSheet = ss.getSheetByName('ID_Cards');
  if(!idCardSheet) {
    idCardSheet = ss.insertSheet('ID_Cards');
    idCardSheet.appendRow(['IDCard', 'TeacherName']);
  }

  let settingSheet = ss.getSheetByName('Settings');
  if(!settingSheet) {
    settingSheet = ss.insertSheet('Settings');
    settingSheet.appendRow(['Key', 'Value']);
    settingSheet.appendRow(['teacherCanEdit', 'true']);
  }
  
  return { ss, userSheet, fileSheet, idCardSheet, settingSheet };
}

// =======================================================
// 1. Core & Settings
// =======================================================
function doGet(e) {
  ensureSheetsExist();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบคลังเอกสารการสอน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSystemSettings() {
  const { settingSheet } = ensureSheetsExist();
  const rows = settingSheet.getDataRange().getValues();
  let settings = { teacherCanEdit: true, subjects: [], levels: [], years: [] };
  
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][0] === 'teacherCanEdit') settings.teacherCanEdit = (rows[i][1] === true || rows[i][1] === 'true');
    if(rows[i][0] === 'subjects') settings.subjects = rows[i][1].toString().split(',');
    if(rows[i][0] === 'levels') settings.levels = rows[i][1].toString().split(',');
    if(rows[i][0] === 'years') settings.years = rows[i][1].toString().split(',');
  }
  return settings;
}

function updateSystemSettings(newSettings) {
  const { settingSheet } = ensureSheetsExist();
  const rows = settingSheet.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][0] === 'teacherCanEdit' && newSettings.teacherCanEdit !== undefined) settingSheet.getRange(i+1, 2).setValue(newSettings.teacherCanEdit.toString());
  }
  return { success: true };
}

// =======================================================
// 2. Registration & Login (ด้วยเลขบัตร)
// =======================================================
function registerUser(data) {
  try {
    const { userSheet, idCardSheet } = ensureSheetsExist();
    
    // 1. เช็คเลขบัตรประชาชนในฐานข้อมูล Whitelist
    const idRows = idCardSheet.getDataRange().getValues();
    let isValidID = false;
    for(let i = 1; i < idRows.length; i++) {
      if(idRows[i][0].toString() === data.idcard) { isValidID = true; break; }
    }
    if(!isValidID) return { success: false, error: 'ไม่พบเลขบัตรประชาชนนี้ในฐานข้อมูล กรุณาติดต่อผู้ดูแลระบบ' };

    // 2. เช็คความซ้ำซ้อน
    const userRows = userSheet.getDataRange().getValues();
    for(let i = 1; i < userRows.length; i++) {
      if(userRows[i][0] === data.email) return { success: false, error: 'อีเมลนี้ถูกใช้งานแล้ว' };
      if(userRows[i][7] && userRows[i][7].toString() === data.idcard) return { success: false, error: 'เลขบัตรประชาชนนี้ลงทะเบียนไปแล้ว' };
    }

    // 3. บันทึกข้อมูล
    userSheet.appendRow([data.email, 'user', data.name, '', '', '', data.password, data.idcard, data.phone]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function loginUser(idcard, password) {
  const { userSheet } = ensureSheetsExist();
  const userRows = userSheet.getDataRange().getValues();
  for (let i = 1; i < userRows.length; i++) {
    if (userRows[i][7] && userRows[i][7].toString().trim() === idcard.toString().trim()) {
      const storedPassword = userRows[i][6]; 
      
      if (storedPassword && password.toString().trim() === storedPassword.toString().trim()) {
        return { success: true, role: userRows[i][1], name: userRows[i][2], email: userRows[i][0] };
      } else return { success: false, error: 'รหัสผ่านไม่ถูกต้อง' };
    }
  }
  return { success: false, error: 'ไม่พบเลขบัตรประชาชนนี้ในระบบ' };
}

// =======================================================
// 3. User Data & File Ops (ส่งเป็น TSV String กลับไป)
// =======================================================
function getUserData(idcard) {
  const { userSheet, fileSheet } = ensureSheetsExist();
  let userProfile = { email: '', name: '', subject: '', role: 'user', profilePic: '', coverPic: '' };
  
  const userRows = userSheet.getDataRange().getValues();
  for (let i = 1; i < userRows.length; i++) {
    if (userRows[i][7] && userRows[i][7].toString() === idcard.toString()) {
      userProfile = { email: userRows[i][0], role: userRows[i][1], name: userRows[i][2], subject: userRows[i][3], profilePic: userRows[i][4], coverPic: userRows[i][5] };
      break;
    }
  }

  let userFiles = [];
  if (userProfile.email !== '') {
      const fileRows = fileSheet.getDataRange().getValues();
      for (let i = 1; i < fileRows.length; i++) {
        if (fileRows[i][1] === userProfile.email) {
          userFiles.push({
            id: fileRows[i][0], name: fileRows[i][2], type: fileRows[i][3], subject: fileRows[i][4], url: fileRows[i][5],
            note: fileRows[i][6] || '', date: Utilities.formatDate(new Date(fileRows[i][7] || new Date()), Session.getScriptTimeZone(), "dd/MM/yyyy"),
            level: fileRows[i][8] || '', term: fileRows[i][9] || '', year: fileRows[i][10] || ''
          });
        }
      }
  }
  
  // แปลง Array of Objects เป็น TSV String
  const filesTSV = convertToTSV(userFiles.reverse(), ['id', 'name', 'type', 'subject', 'url', 'note', 'date', 'level', 'term', 'year']);
  
  return { profile: userProfile, files: filesTSV };
}

function getPublicRemedialFiles() {
  const { userSheet, fileSheet } = ensureSheetsExist();
  const userRows = userSheet.getDataRange().getValues();
  const fileRows = fileSheet.getDataRange().getValues();
  let publicFiles = [];
  for (let i = 1; i < fileRows.length; i++) {
    if (fileRows[i][3] === 'ใบงานแก้ผลการเรียน') {
      let teacherName = 'ไม่ระบุชื่อ';
      for (let j = 1; j < userRows.length; j++) { if(userRows[j][0] === fileRows[i][1]) { teacherName = userRows[j][2]; break; } }
      publicFiles.push({
        id: fileRows[i][0], teacherName: teacherName, name: fileRows[i][2], type: fileRows[i][3], subject: fileRows[i][4], 
        url: fileRows[i][5], note: fileRows[i][6] || '', date: Utilities.formatDate(new Date(fileRows[i][7] || new Date()), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        level: fileRows[i][8] || '', term: fileRows[i][9] || '', year: fileRows[i][10] || ''
      });
    }
  }
  
  // แปลง Array of Objects เป็น TSV String
  return convertToTSV(publicFiles.reverse(), ['id', 'teacherName', 'name', 'type', 'subject', 'url', 'note', 'date', 'level', 'term', 'year']);
}

function uploadFileToSystem(data) {
  try {
    const folder = getDriveFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType || 'application/octet-stream', data.fileName);
    const file = folder.createFile(blob);
    const { fileSheet } = ensureSheetsExist();
    fileSheet.appendRow([
        file.getId(), data.userEmail, data.docName, data.docType, data.docSubject, file.getUrl(), data.docNote || '', new Date(),
        data.docLevel || '', data.docTerm || '', data.docYear || ''
    ]);
    return { success: true, url: file.getUrl() };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function editFileInSystem(data) {
  try {
    const { fileSheet } = ensureSheetsExist();
    const fileRows = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileRows.length; i++) {
      if (fileRows[i][0] === data.fileId) {
        fileSheet.getRange(i + 1, 3).setValue(data.docName);
        fileSheet.getRange(i + 1, 4).setValue(data.docType);
        fileSheet.getRange(i + 1, 5).setValue(data.docSubject);
        fileSheet.getRange(i + 1, 7).setValue(data.docNote || '');
        fileSheet.getRange(i + 1, 8).setValue(new Date());
        fileSheet.getRange(i + 1, 9).setValue(data.docLevel || '');
        fileSheet.getRange(i + 1, 10).setValue(data.docTerm || '');
        fileSheet.getRange(i + 1, 11).setValue(data.docYear || '');

        if (data.base64) {
          const folder = getDriveFolder();
          const newFile = folder.createFile(Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType || 'application/octet-stream', data.fileName));
          fileSheet.getRange(i + 1, 1).setValue(newFile.getId()); fileSheet.getRange(i + 1, 6).setValue(newFile.getUrl());
          try { DriveApp.getFileById(data.fileId).setTrashed(true); } catch(e){}
        }
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบไฟล์' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function deleteFileFromSystem(fileId) {
  try {
    const { fileSheet } = ensureSheetsExist();
    const fileRows = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileRows.length; i++) {
      if (fileRows[i][0] === fileId) {
        fileSheet.deleteRow(i + 1);
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบไฟล์' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function updateUserProfile(data) {
  try {
    const { userSheet } = ensureSheetsExist();
    const userRows = userSheet.getDataRange().getValues();
    let pPic = data.profilePicUrl || '', cPic = data.coverPicUrl || '';

    if (data.profilePicBase64) pPic = getDriveFolder().createFile(Utilities.newBlob(Utilities.base64Decode(data.profilePicBase64), data.profilePicMimeType, "profile_" + data.userEmail)).getUrl();
    if (data.coverPicBase64) cPic = getDriveFolder().createFile(Utilities.newBlob(Utilities.base64Decode(data.coverPicBase64), data.coverPicMimeType, "cover_" + data.userEmail)).getUrl();
    
    for (let i = 1; i < userRows.length; i++) {
      if (userRows[i][0] === data.userEmail) {
        userSheet.getRange(i + 1, 3).setValue(data.name); userSheet.getRange(i + 1, 4).setValue(data.subject);
        if (pPic) userSheet.getRange(i + 1, 5).setValue(pPic);
        if (cPic) userSheet.getRange(i + 1, 6).setValue(cPic);
        return { success: true, profilePic: pPic, coverPic: cPic };
      }
    }
    return { success: false, error: 'ไม่พบผู้ใช้' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// =======================================================
// 4. Admin & Super Admin APIs (ส่งเป็น TSV String กลับไป)
// =======================================================
function getAdminDashboard() {
  const { userSheet, fileSheet } = ensureSheetsExist();
  const userRows = userSheet.getDataRange().getValues();
  const fileRows = fileSheet.getDataRange().getValues();
  
  let reports = [];
  for (let i = fileRows.length - 1; i >= 1; i--) {
    let uName = fileRows[i][1];
    for (let j = 1; j < userRows.length; j++) { if(userRows[j][0] === fileRows[i][1]) uName = userRows[j][2]; }
    reports.push({
      id: fileRows[i][0], email: fileRows[i][1], userName: uName, name: fileRows[i][2], docType: fileRows[i][3], subject: fileRows[i][4],
      url: fileRows[i][5], note: fileRows[i][6] || '', date: Utilities.formatDate(new Date(fileRows[i][7] || new Date()), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      level: fileRows[i][8] || '', term: fileRows[i][9] || '', year: fileRows[i][10] || ''
    });
  }
  let allUsers = [];
  for (let j = 1; j < userRows.length; j++) { 
    if(userRows[j][0]) allUsers.push({ email: userRows[j][0], name: userRows[j][2] }); 
  }
  
  // แปลง Array ทั้ง 2 ชุดให้เป็น TSV String
  return { 
      totalUsers: userRows.length-1, 
      totalFiles: fileRows.length-1, 
      reports: convertToTSV(reports, ['id', 'email', 'userName', 'name', 'docType', 'subject', 'url', 'note', 'date', 'level', 'term', 'year']), 
      allUsers: convertToTSV(allUsers, ['email', 'name']) 
  };
}

// ---- จัดการผู้ใช้ (Super Admin) ----
function getAllSystemUsers() {
  const { userSheet } = ensureSheetsExist();
  const rows = userSheet.getDataRange().getValues();
  let users = [];
  for(let i=1; i<rows.length; i++) {
    users.push({ email: rows[i][0], role: rows[i][1], name: rows[i][2], subject: rows[i][3], password: rows[i][6], idcard: rows[i][7], phone: rows[i][8] });
  }
  return convertToTSV(users, ['email', 'role', 'name', 'subject', 'password', 'idcard', 'phone']);
}

function addSystemUser(data) {
  try {
    const { userSheet } = ensureSheetsExist();
    const rows = userSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === data.email) return { success: false, error: 'อีเมลนี้ซ้ำ' }; }
    userSheet.appendRow([data.email, data.role || 'user', data.name, data.subject, '', '', data.password, '', '']);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function updateSystemUser(data) {
  try {
    const { userSheet } = ensureSheetsExist();
    const rows = userSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { 
      if (rows[i][0] === data.oldEmail) {
        userSheet.getRange(i+1, 1).setValue(data.email);
        userSheet.getRange(i+1, 2).setValue(data.role);
        userSheet.getRange(i+1, 3).setValue(data.name);
        userSheet.getRange(i+1, 4).setValue(data.subject);
        userSheet.getRange(i+1, 7).setValue(data.password);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบผู้ใช้' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function deleteSystemUser(email) {
  try {
    const { userSheet } = ensureSheetsExist();
    const rows = userSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === email) { userSheet.deleteRow(i+1); return { success: true }; } }
    return { success: false, error: 'ไม่พบผู้ใช้' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- จัดการเลขบัตร (Super Admin) ----
function getIDCards() {
  const { idCardSheet } = ensureSheetsExist();
  const rows = idCardSheet.getDataRange().getValues();
  let cards = [];
  for(let i = 1; i < rows.length; i++) { cards.push({ id: rows[i][0], name: rows[i][1] }); }
  return convertToTSV(cards, ['id', 'name']);
}

function addIDCard(id, name) {
  const { idCardSheet } = ensureSheetsExist();
  const rows = idCardSheet.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++) { if(rows[i][0].toString() === id) return {success:false, error:'มีเลขบัตรนี้อยู่แล้ว'}; }
  idCardSheet.appendRow([id, name]);
  return {success:true};
}

function editIDCard(oldId, newId, newName) {
  const { idCardSheet } = ensureSheetsExist();
  const rows = idCardSheet.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++) { 
    if(rows[i][0].toString() === oldId) {
      idCardSheet.getRange(i+1, 1).setValue(newId);
      idCardSheet.getRange(i+1, 2).setValue(newName);
      return {success:true};
    }
  }
  return {success:false, error:'ไม่พบข้อมูล'};
}

function deleteIDCard(id) {
  const { idCardSheet } = ensureSheetsExist();
  const rows = idCardSheet.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][0].toString() === id) { idCardSheet.deleteRow(i+1); return {success:true}; }
  }
  return {success:false, error:'ไม่พบข้อมูล'};
}