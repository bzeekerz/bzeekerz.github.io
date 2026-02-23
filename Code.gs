// --- CONFIGURATION ---
const SPREADSHEET_ID = '1u8OaGgDcpgWdtaqTXpwWm8PX2b4I2Ovq93aKRuXol18'; 
const TEMPLATE_SLIDE_ID = '1FEVxooVLLEmxUscy6dXiPZHPjqMn8Bu7NEAXdQ19k-w';
const DESTINATION_FOLDER_ID = '1u1LpLsCDaUgwWYJIXn5L9D_a1sBhKoU7';

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.urlParams = JSON.stringify(e.parameter || {});
  template.serverMessage = ""; 
  template.serverStatus = "";

  if (e.parameter && e.parameter.page === 'verify' && e.parameter.token) {
    const result = verifyUserToken(e.parameter.token);
    template.serverMessage = result.message;
    template.serverStatus = result.status;
  }

  return template.evaluate()
    .setTitle('ระบบคำร้องออนไลน์ (JC Form)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }
function generateToken() { return Utilities.getUuid(); }

function hashPassword(password, salt) {
  if (salt == null) salt = "";
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  let txtHash = '';
  for (let i = 0; i < rawBytes.length; i++) {
    let hashVal = rawBytes[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function generateSalt() { return Utilities.getUuid(); }

function sendEmail(to, subject, body) {
  try { MailApp.sendEmail({ to: to, subject: subject, htmlBody: body });
  } catch(e) { console.log("Email Error: " + e.toString()); }
}

function getMOTD() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MOTD');
    if (!sheet) return "";
    return sheet.getRange(1, 1).getValue(); 
  } catch (e) { return ""; }
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { status: 'error', message: 'ระบบเพิ่งเริ่มต้น กรุณาสมัครสมาชิกใหม่' };

  const data = userSheet.getDataRange().getValues();
  const userRow = data.find(row => row[0] == username);
  if (userRow) {
    if (String(userRow[9]).toUpperCase() !== 'TRUE') return { status: 'error', message: 'กรุณายืนยันตัวตนทาง Email ก่อน' };
    let role = (userRow.length > 12 && userRow[12]) ? userRow[12] : 'user';
    let status = (userRow.length > 13 && userRow[13]) ? userRow[13] : 'active';
    if (String(status).toLowerCase() === 'banned') return { status: 'error', message: 'บัญชีของคุณถูกระงับการใช้งาน' };

    const storedHash = userRow[1];
    const storedSalt = userRow[14] || "";
    if (hashPassword(password, storedSalt) === storedHash) {
        return { 
          status: 'success', username: String(userRow[0]), name: userRow[2], std_id: userRow[3],
          email: userRow[4], tel: userRow[5], year: userRow[6], gender: userRow[7], role: role
        };
    }
  } 
  return { status: 'error', message: 'Username หรือ Password ไม่ถูกต้อง' };
}

function registerUser(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Name', 'Std_ID', 'Email', 'Tel', 'Year', 'Gender', 'Token', 'Verified', 'Reset_Token', 'Reset_Exp', 'Role', 'Status', 'Salt']);
  }
  
  const data = userSheet.getDataRange().getValues();
  if (data.some(row => String(row[0]) === String(formObject.reg_username))) return { status: 'error', message: 'Username นี้ถูกใช้ไปแล้ว' };
  if (data.some(row => row[4] === formObject.reg_email)) return { status: 'error', message: 'Email นี้ถูกใช้ไปแล้ว' };

  const salt = generateSalt();
  const hashedPassword = hashPassword(formObject.reg_password, salt);
  const verifyToken = generateToken();
  const verifyLink = `${getScriptUrl()}?page=verify&token=${verifyToken}`;
  userSheet.appendRow([formObject.reg_username, hashedPassword, formObject.reg_name, formObject.reg_std_id, formObject.reg_email, "'" + formObject.reg_tel, formObject.reg_year, formObject.reg_gender, verifyToken, 'FALSE', '', '', 'user', 'active', salt]);
  sendEmail(formObject.reg_email, 'ยืนยันการสมัคร', `<p><a href="${verifyLink}">คลิกยืนยันตัวตน</a></p>`);
  return { status: 'success', message: 'สมัครสำเร็จ! กรุณาตรวจสอบ Email' };
}

function verifyUserToken(token) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[8] === token);
  if (rowIndex > 0) {
    userSheet.getRange(rowIndex + 1, 9).setValue('');
    userSheet.getRange(rowIndex + 1, 10).setValue('TRUE'); 
    return { status: 'success', message: 'ยืนยันตัวตนสำเร็จ!' };
  }
  return { status: 'error', message: 'ลิงก์ไม่ถูกต้อง' };
}

function requestPasswordReset(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[4] === email);
  if (rowIndex > 0) {
    const token = generateToken();
    const link = `${getScriptUrl()}?page=reset&token=${token}`;
    userSheet.getRange(rowIndex + 1, 11).setValue(token);
    userSheet.getRange(rowIndex + 1, 12).setValue(new Date().getTime() + 3600000);
    sendEmail(email, 'Reset Password', `<a href="${link}">Reset Password</a>`);
  }
  return { status: 'success', message: 'ส่งลิงก์แล้ว' };
}

function submitResetPassword(token, newPass) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[10] === token);
  if (rowIndex > 0) {
    if (new Date().getTime() > data[rowIndex][11]) return { status: 'error', message: 'ลิงก์หมดอายุ' };
    const newSalt = generateSalt();
    const newHash = hashPassword(newPass, newSalt);
    userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
    userSheet.getRange(rowIndex + 1, 11).setValue('');
    userSheet.getRange(rowIndex + 1, 15).setValue(newSalt);
    return { status: 'success', message: 'เปลี่ยนรหัสสำเร็จ' };
  }
  return { status: 'error', message: 'Token ผิด' };
}

function changePassword(user, oldPass, newPass) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] == user);
  if(rowIndex > 0) {
    const userData = data[rowIndex];
    if (hashPassword(oldPass, userData[14] || "") === userData[1]) {
        const newSalt = generateSalt();
        const newHash = hashPassword(newPass, newSalt);
        userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
        userSheet.getRange(rowIndex + 1, 15).setValue(newSalt);
        return { status: 'success', message: 'เปลี่ยนรหัสเรียบร้อย' };
    }
  }
  return { status: 'error', message: 'รหัสเดิมผิด' };
}

async function processForm(formData, userInfo) {
  try {
    const destFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const templateFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
    
    let fileName = formData.custom_filename || `Request_${userInfo.std_id}_${new Date().getTime()}`;
    const copyFile = templateFile.makeCopy(fileName, destFolder);
    const copyId = copyFile.getId();
    const slide = SlidesApp.openById(copyId);

    if (formData.signature_data) {
      const firstSlide = slide.getSlides()[0];
      replaceTextWithImage(firstSlide, '{{signature}}', formData.signature_data);
    }

    let reqType = formData.request_type;
    const val = (topic, value) => (reqType === topic || (Array.isArray(topic) && topic.includes(reqType))) ? value : "";
    const replace = (key, value) => slide.replaceAllText(`{{${key}}}`, value || " ");
    const tick = "✓";
    
    replace('male', userInfo.gender === 'male' ? tick : "");
    replace('female', userInfo.gender === 'female' ? tick : "");
    replace('BJM', formData.program === 'BJM' ? tick : "");
    replace('Thai', formData.program === 'Thai' ? tick : "");
    for(let i=1; i<=10; i++) replace(`t${i}`, (reqType === `t${i}`) ? tick : "");

    // ไม่ต้องมี Truncate ฝั่ง Backend แล้วครับ เพราะหน้าเว็บใช้ Canvas วัด Pixel มาเป๊ะ 100% แล้ว
    replace('name', formData.reg_name || userInfo.name || "");
    replace('std_id', userInfo.std_id || "");
    replace('Year', formData.year || "");
    replace('advisor', formData.advisor || "");
    replace('major', formData.major || ""); 
    replace('address', formData.address || ""); 
    replace('tel', (formData.tel || "").replace(/\D/g,''));
    replace('email', formData.email || ""); 
    
    let specificData = "";
    if (reqType === 't1') specificData = formData.major_sel || "";
    else if (reqType === 't2') specificData = `${formData.major_from || ""} ไปยัง ${formData.major_to || ""}`;
    else if (reqType === 't3') specificData = `${formData.prof_rec || ""} (${formData.r_no || ""})`;
    else if (reqType === 't5') specificData = `${formData.reg_sem || ""}/${formData.reg_year || ""} ${formData.reg_reasson || ""}`;
    else if (reqType === 't6') specificData = `${formData.re_ad || ""}/${formData.re_ad_year || ""}`;
    else if (reqType === 't7' || reqType === 't8') specificData = formData.location || "";
    else if (reqType === 't9') specificData = formData.items || "";
    else if (reqType === 't10') specificData = formData.other || "";
    
    replace('major_sel',  val('t1', formData.major_sel));
    replace('major_from', val('t2', formData.major_from));
    replace('major_to',   val('t2', formData.major_to));
    replace('prof_rec',   val('t3', formData.prof_rec));
    replace('r_no',       val('t3', formData.r_no));
    replace('reg_sem',    val('t5', formData.reg_sem));
    replace('reg_year',   val('t5', formData.reg_year));
    replace('reg_reasson',val('t5', formData.reg_reasson)); 
    replace('re_ad',      val('t6', formData.re_ad));
    replace('re_ad_year', val('t6', formData.re_ad_year));
    replace('location',   val(['t7', 't8'], formData.location)); 
    replace('items',      val('t9', formData.items)); 
    replace('other',      val('t10', formData.other)); 

    // ดึงค่า res_1, res_2, res_3 ไปใช้ตรงๆ 
    replace('res_1', formData.res_1 || "");
    replace('res_2', formData.res_2 || "");
    replace('res_3', formData.res_3 || "");

    slide.saveAndClose();

    let mainPdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
    DriveApp.getFileById(copyId).setTrashed(true);

    let finalPdfBlob = mainPdfBlob;
    if (formData.fileAttachment) {
      try {
        const attachmentBlob = Utilities.newBlob(Utilities.base64Decode(formData.fileAttachment.content), formData.fileAttachment.mimeType, formData.fileAttachment.name);
        let mergedBlob;
        if (typeof PDFApp !== 'undefined' && PDFApp.mergePDFs) { mergedBlob = await PDFApp.mergePDFs([mainPdfBlob, attachmentBlob]); } 
        else if (typeof mergePDFs === 'function') { mergedBlob = await mergePDFs([mainPdfBlob, attachmentBlob]); }
        if (mergedBlob && typeof mergedBlob.setName === 'function') { finalPdfBlob = mergedBlob; }
      } catch (mergeErr) { console.log("Merge Error: " + mergeErr); }
    }

    finalPdfBlob.setName(fileName + ".pdf");
    const finalFile = destFolder.createFile(finalPdfBlob);
    const pdfUrl = finalFile.getUrl();
    const fileId = finalFile.getId();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('Logs');
    if(!logSheet) { 
      logSheet = ss.insertSheet('Logs');
      logSheet.appendRow(['Timestamp', 'Username', 'Name', 'File Name', 'Type', 'URL', 'File ID', 'Program', 'Gender', 'Year', 'Tel', 'Major', 'Advisor', 'Email', 'Address', 'Topic Data', 'Reason', 'Status', 'Student_File', 'Admin_File', 'Raw_Data']);
    }
    if (logSheet.getLastColumn() < 21) logSheet.insertColumnsAfter(logSheet.getLastColumn(), 21 - logSheet.getLastColumn());

    let displayName = userInfo.name;
    if (displayName && userInfo.gender && !displayName.startsWith('นาย') && !displayName.startsWith('นาง')) {
        displayName = (userInfo.gender === 'male' ? 'นาย' : 'นางสาว') + displayName;
    }

    const rawDataJSON = JSON.stringify(formData);
    // รวมข้อความสำหรับเก็บบันทึก (Logging) เท่านั้น
    const combinedReasonLog = [formData.res_1, formData.res_2, formData.res_3].filter(x => x).join(' ');

    logSheet.appendRow([
      new Date(), String(userInfo.username), displayName, fileName, reqType, pdfUrl, fileId, 
      formData.program, userInfo.gender, formData.year, "'" + formData.tel, formData.major, 
      formData.advisor, formData.email, formData.address, specificData, combinedReasonLog,
      'รอ', '', '', rawDataJSON 
    ]);

    try {
        const topicMap = {
          't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา', 't3': 'ขอหนังสือรับรองความประพฤติ', 't4': 'ขออนุมัติย้ายคณะ', 't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา', 't7': 'ขอใช้สถานที่', 't8': 'ขออนุญาตใช้ห้อง', 't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ'
        };
        const topicName = topicMap[reqType] || reqType;
        const lineMsg = `🔔 มีคำร้องใหม่ (พร้อมไฟล์แนบ)!\n👤 ชื่อ: ${displayName} (${userInfo.std_id})\n📝 เรื่อง: ${topicName}\n📂 PDF: ${pdfUrl}`;
        sendLinePushMessage(lineMsg);
    } catch(err) {}

    return { status: 'success', url: pdfUrl };
  } catch (e) { return { status: 'error', message: 'Error: ' + e.toString() }; }
}

function getRequestsData(user) {
  if (!user || !user.username) return [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  let userMap = {};
  if (userSheet) {
     const uData = userSheet.getDataRange().getValues();
     uData.forEach(r => { if(r[0]) userMap[String(r[0])] = { name: r[2], std_id: r[3], gender: r[7] }; });
  }

  let sheet = ss.getSheetByName('Logs');
  if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  let requests = data.map(r => {
    try {
        let ts = r[0];
        let timeStr = (ts instanceof Date) ? Utilities.formatDate(ts, "GMT+7", "dd/MM/yyyy HH:mm") : String(ts || "-");
        let username = String(r[1] || ""); let logName = String(r[2] || "");
        let userInfo = userMap[username] || { name: "-", std_id: "-", gender: "" };
        let finalName = logName !== "-" && logName !== "" ? logName : userInfo.name;
        let rawData = {};
        try { if(r[20] && r[20] !== "") rawData = JSON.parse(r[20]); } catch(e) {}
        
        return {
            timestamp: timeStr, username: username, name: finalName, std_id: String(userInfo.std_id),
            fileName: String(r[3] || "ไม่มีชื่อไฟล์"), type: String(r[4] || ""), pdfUrl: String(r[5] || "#"),  
            fileId: String(r[6] || ""), program: String(r[7] || ""), year: String(r[9] || ""),     
            tel: String(r[10] || "").replace(/'/g, ''), major: String(r[11] || ""), advisor: String(r[12] || ""),
            email: String(r[13] || ""), address: String(r[14] || ""), reason: String(r[16] || ""), 
            status: String(r[17] || "รอ"), studentFile: String(r[18] || ""), adminFile: String(r[19] || ""),
            ...rawData
        };
    } catch (err) { return null; }
  }).filter(item => item !== null);
  if (user.role !== 'admin') requests = requests.filter(r => r.username === String(user.username));
  return requests.reverse();
}

async function uploadFile(base64Data, fileType, relatedFileId, uploaderRole, username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Logs');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[6] === relatedFileId);
    if (rowIndex <= 0) return { status: 'error', message: 'ไม่พบรายการ' };

    const splitBase = base64Data.split(','); const decoded = Utilities.base64Decode(splitBase[1]);
    let uploadBlob = Utilities.newBlob(decoded, fileType, `Upload_${new Date().getTime()}.pdf`);
    let timestampText = "Received: " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    if (fileType === 'application/pdf' && typeof PDFApp !== 'undefined') {
       try {
         const newPdfBlob = await PDFApp.setPDFBlob(uploadBlob).insertHeaderFooter({ header: { left: { text: timestampText, size: 3, x: 20, yOffset: 10 } } });
         if (newPdfBlob) { uploadBlob = newPdfBlob; uploadBlob.setName(`Upload_${new Date().getTime()}.pdf`); }
       } catch (e) { console.log("PDFApp Stamp Error: " + e.toString()); }
    }

    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const file = folder.createFile(uploadBlob);
    const fileUrl = file.getUrl();

    if (sheet.getLastColumn() < 21) sheet.insertColumnsAfter(sheet.getLastColumn(), 21 - sheet.getLastColumn());
    sheet.getRange(rowIndex + 1, 1).setValue(new Date());

    if (uploaderRole === 'admin') {
      const oldAdminUrl = data[rowIndex][19];
      if (oldAdminUrl && String(oldAdminUrl).includes('drive.google.com')) {
          try { const match = String(oldAdminUrl).match(/[-\w]{25,}/); if (match) DriveApp.getFileById(match[0]).setTrashed(true); } catch(e) {}
      }
      sheet.getRange(rowIndex + 1, 20).setValue(fileUrl); sheet.getRange(rowIndex + 1, 18).setValue('เสร็จสิ้น');
    } else {
      if (String(data[rowIndex][1]) !== String(username)) return { status: 'error', message: 'ไม่มีสิทธิ์' };
      const oldStudentUrl = data[rowIndex][18];
      if (oldStudentUrl && String(oldStudentUrl).includes('drive.google.com')) {
          try { const match = String(oldStudentUrl).match(/[-\w]{25,}/); if (match) DriveApp.getFileById(match[0]).setTrashed(true); } catch(e) {}
      }
      sheet.getRange(rowIndex + 1, 19).setValue(fileUrl); sheet.getRange(rowIndex + 1, 18).setValue('รอเจ้าหน้าที่ตรวจสอบ'); 
      try {
        const topicMap = { 't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา', 't3': 'ขอหนังสือรับรอง', 't4': 'ขออนุมัติย้ายคณะ', 't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา', 't7': 'ขอใช้สถานที่', 't8': 'ขออนุญาตใช้ห้อง', 't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ' };
        const reqType = data[rowIndex][4]; const topicName = topicMap[reqType] || reqType; const nameShow = data[rowIndex][2] || username;
        const lineMsg = `🔄 Updated นักศึกษาส่งไฟล์ใหม่แล้ว!\nกำลังรอรับเจ้าหน้าที่รับเรื่อง\n👤 จาก: ${nameShow}\n📝 เรื่อง: ${topicName}\n⏱️ ส่งเมื่อ: ${timestampText}\n📂 ไฟล์แนบ: ${fileUrl}`;
        sendLinePushMessage(lineMsg);
      } catch(err) {}
    }
    return { status: 'success', message: 'อัปโหลดสำเร็จ (ไฟล์เก่าถูกลบแล้ว)' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function adminUpdateStatus(fileId, newStatus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[6] === fileId);
  if (rowIndex > 0) {
    if (sheet.getLastColumn() < 18) sheet.insertColumnsAfter(sheet.getLastColumn(), 18 - sheet.getLastColumn());
    sheet.getRange(rowIndex + 1, 18).setValue(newStatus);
    return 'อัปเดตสถานะเรียบร้อย';
  }
  return 'ไม่พบรายการ';
}

function adminBanUser(targetEmail) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[4] === targetEmail);
  if (rowIndex > 0) {
    if (sheet.getLastColumn() < 14) sheet.insertColumnsAfter(sheet.getLastColumn(), 14 - sheet.getLastColumn());
    sheet.getRange(rowIndex + 1, 14).setValue('banned');
    return `ระงับการใช้งาน ${targetEmail} แล้ว`;
  }
  return 'ไม่พบ Email นี้';
}

function deleteHistory(fileId, username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const userRows = userSheet.getDataRange().getValues();
  const currentUser = userRows.find(row => String(row[0]) === String(username));
  const isAdmin = currentUser && currentUser[12] === 'admin'; 

  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[6] === fileId && (String(r[1]) === String(username) || isAdmin));
  if(rowIndex > 0) { 
      const rowData = data[rowIndex];
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
      if (rowData[18] && String(rowData[18]).includes('drive.google.com')) { try { const match = String(rowData[18]).match(/[-\w]{25,}/); if (match) DriveApp.getFileById(match[0]).setTrashed(true); } catch(e) {} }
      if (rowData[19] && String(rowData[19]).includes('drive.google.com')) { try { const match = String(rowData[19]).match(/[-\w]{25,}/); if (match) DriveApp.getFileById(match[0]).setTrashed(true); } catch(e) {} }
      sheet.deleteRow(rowIndex + 1);
      return { status: 'success', message: 'ลบข้อมูลและไฟล์แนบทั้งหมดเรียบร้อย' };
  }
  return { status: 'error', message: 'เกิดข้อผิดพลาด: คุณไม่มีสิทธิ์ลบไฟล์นี้ หรือไม่พบไฟล์' };
}

function renameHistory(fileId, newName, username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const userSheet = ss.getSheetByName('Users');
  const userRows = userSheet.getDataRange().getValues();
  const userObj = userRows.find(row => String(row[0]) === String(username));
  const isAdmin = userObj && userObj[12] === 'admin';

  const rowIndex = data.findIndex(r => r[6] === fileId && (String(r[1]) === String(username) || isAdmin));
  if(rowIndex > 0) {
      try { DriveApp.getFileById(fileId).setName(newName); } catch(e){}
      sheet.getRange(rowIndex + 1, 4).setValue(newName);
      return { status: 'success', message: 'แก้ไขเรียบร้อย' };
  }
  return { status: 'error', message: 'Error' };
}

function replaceTextWithImage(slide, searchText, base64Data) {
  if (!base64Data) return;
  const encodedImage = base64Data.split(',')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(encodedImage), 'image/png', 'signature.png');
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    if (shape.getText().asString().includes(searchText)) {
      slide.insertImage(blob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
      shape.remove();
      break;
    }
  }
}

function getTemplateData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) { configSheet = ss.insertSheet('Config'); configSheet.appendRow(['Major', 'Advisor']); }

  let majors = [], teachers = [];
  if (configSheet.getLastRow() > 1) {
    const d = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
    majors = d.map(r => r[0]).filter(String);
    teachers = d.map(r => r[1]).filter(String);
  }

  let tempSheet = ss.getSheetByName('Templates');
  if (!tempSheet) { tempSheet = ss.insertSheet('Templates'); tempSheet.appendRow(['Name', 'Topic', 'Data', 'Reason']); }
  let templates = [];
  if (tempSheet.getLastRow() > 1) {
    const d = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, 4).getValues();
    templates = d.map(r => ({ name: r[0], topic: r[1], data: r[2], reason: r[3] })).filter(t => t.name);
  }
  return { majors, teachers, templates };
}

function sendLinePushMessage(message) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = sheet.getSheetByName("Config_Line"); 
    if (!configSheet) return;
    var accessToken = configSheet.getRange("B1").getValue();
    var targetId = configSheet.getRange("B2").getValue();
    if (!accessToken || !targetId) return;

    var url = "https://api.line.me/v2/bot/message/push";
    var payload = JSON.stringify({ "to": targetId, "messages": [{ "type": "text", "text": message }] });
    UrlFetchApp.fetch(url, { "method": "post", "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + accessToken }, "payload": payload });
  } catch (e) { console.log("Line Error: " + e.toString()); }
}

function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    if (json.events.length === 0) return;
    var event = json.events[0];
    var msg = event.message.text || "";
    var type = event.source.type;
    var id = (type === "group") ? event.source.groupId : event.source.userId;

    if (msg.toLowerCase().includes("check")) { 
       MailApp.sendEmail({ to: "nitichan@tu.ac.th", subject: "📌 ได้ Group ID แล้วครับ!", htmlBody: "<h3>ข้อมูลจาก LINE (" + type + ")</h3><p><b>Group ID / User ID คือ:</b></p><h2>" + id + "</h2>" });
    }
  } catch (error) {
    MailApp.sendEmail({ to: "nitichan@tu.ac.th", subject: "❌ ระบบ Error", body: "Error: " + error.toString() });
  }
}
