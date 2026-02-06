// --- CONFIGURATION ---
const SPREADSHEET_ID = '1u8OaGgDcpgWdtaqTXpwWm8PX2b4I2Ovq93aKRuXol18'; 
const TEMPLATE_SLIDE_ID = '1FEVxooVLLEmxUscy6dXiPZHPjqMn8Bu7NEAXdQ19k-w';
const DESTINATION_FOLDER_ID = '1u1LpLsCDaUgwWYJIXn5L9D_a1sBhKoU7';

// --- ROUTING & INIT ---
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
  // กรณีไม่มี salt (เผื่อข้อมูลเก่า) ให้เป็นค่าว่าง
  if (salt == null) salt = ""; 
  
  // ผสม password กับ salt ก่อนเข้า SHA-256
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

// เพิ่มฟังก์ชันสร้าง Salt (ใช้ UUID เพราะไม่ซ้ำแน่นอน)
function generateSalt() {
  return Utilities.getUuid();
}

function sendEmail(to, subject, body) {
  try {
    MailApp.sendEmail({ to: to, subject: subject, htmlBody: body });
  } catch(e) { console.log("Email Error: " + e.toString()); }
}

function getMOTD() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MOTD');
    if (!sheet) return "";
    return sheet.getRange(1, 1).getValue(); 
  } catch (e) {
    return "";
  }
}

// --- USER MANAGEMENT ---
function loginUser(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let userSheet = ss.getSheetByName('Users');
  
  // ถ้ายังไม่มี Sheet ให้สร้างใหม่พร้อม Header ที่มี Salt
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Name', 'Std_ID', 'Email', 'Tel', 'Year', 'Gender', 'Token', 'Verified', 'Reset_Token', 'Reset_Exp', 'Role', 'Status', 'Salt']);
    return { status: 'error', message: 'ระบบเพิ่งเริ่มต้น กรุณาสมัครสมาชิกใหม่' };
  }

  const data = userSheet.getDataRange().getValues();
  
  // 1. ค้นหาเฉพาะ Username ก่อน (ยังไม่เช็ค Password ตรงนี้)
  const userRow = data.find(row => row[0] == username);
  
  if (userRow) {
    // ตรวจสอบสถานะการยืนยันตัวตนและ Ban
    if (String(userRow[9]).toUpperCase() !== 'TRUE') {
      return { status: 'error', message: 'กรุณายืนยันตัวตนทาง Email ก่อน' };
    }
    
    let role = (userRow.length > 12 && userRow[12]) ? userRow[12] : 'user';
    let status = (userRow.length > 13 && userRow[13]) ? userRow[13] : 'active';
    
    if (String(status).toLowerCase() === 'banned') {
      return { status: 'error', message: 'บัญชีของคุณถูกระงับการใช้งาน' };
    }

    // --- ส่วนตรวจสอบรหัสผ่านแบบ Salted ---
    const storedHash = userRow[1];      // รหัสผ่านที่ถูกแฮชใน DB
    const storedSalt = userRow[14] || ""; // ดึง Salt จากคอลัมน์ที่ 15 (Index 14)
    
    // เอาพาสที่กรอก + Salt ใน DB มาแฮชใหม่ แล้วเทียบกัน
    if (hashPassword(password, storedSalt) === storedHash) {
        return { 
          status: 'success', 
          username: userRow[0], 
          name: userRow[2], 
          std_id: userRow[3],
          email: userRow[4], 
          tel: userRow[5],
          year: userRow[6],
          gender: userRow[7],
          role: role
        };
    }
  } 
  
  // ถ้าหา User ไม่เจอ หรือ Password ผิด
  return { status: 'error', message: 'Username หรือ Password ไม่ถูกต้อง' };
}

function registerUser(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let userSheet = ss.getSheetByName('Users');
  
  // อัปเดต Header ให้มี Salt หากสร้าง Sheet ใหม่
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Name', 'Std_ID', 'Email', 'Tel', 'Year', 'Gender', 'Token', 'Verified', 'Reset_Token', 'Reset_Exp', 'Role', 'Status', 'Salt']);
  }
  
  const data = userSheet.getDataRange().getValues();
  if (data.some(row => row[0] === formObject.reg_username)) return { status: 'error', message: 'Username นี้ถูกใช้ไปแล้ว' };
  if (data.some(row => row[4] === formObject.reg_email)) return { status: 'error', message: 'Email นี้ถูกใช้ไปแล้ว' };

  // --- สร้าง Salt และ Hash ---
  const salt = generateSalt(); 
  const hashedPassword = hashPassword(formObject.reg_password, salt);
  
  const verifyToken = generateToken();
  const verifyLink = `${getScriptUrl()}?page=verify&token=${verifyToken}`;

  // บันทึกข้อมูล โดยเพิ่ม salt ต่อท้ายสุด (คอลัมน์ที่ 15)
  userSheet.appendRow([
    formObject.reg_username, 
    hashedPassword, 
    formObject.reg_name, 
    formObject.reg_std_id,
    formObject.reg_email, 
    "'" + formObject.reg_tel, 
    formObject.reg_year, 
    formObject.reg_gender,
    verifyToken, 
    'FALSE', 
    '', 
    '', 
    'user', 
    'active',
    salt // <--- เพิ่ม Salt ตรงนี้
  ]);
  
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
  
  // ค้นหาจาก Token
  const rowIndex = data.findIndex(row => row[10] === token); // Index 10 = Reset_Token
  
  if (rowIndex > 0) {
    if (new Date().getTime() > data[rowIndex][11]) return { status: 'error', message: 'ลิงก์หมดอายุ' };
    
    // --- สร้าง Salt ใหม่ ---
    const newSalt = generateSalt();
    const newHash = hashPassword(newPass, newSalt);

    // อัปเดต Password (คอลัมน์ 2)
    userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
    // ล้าง Token
    userSheet.getRange(rowIndex + 1, 11).setValue('');
    // อัปเดต Salt ใหม่ (คอลัมน์ 15)
    userSheet.getRange(rowIndex + 1, 15).setValue(newSalt);
    
    return { status: 'success', message: 'เปลี่ยนรหัสสำเร็จ' };
  }
  return { status: 'error', message: 'Token ผิด' };
}

function changePassword(user, oldPass, newPass) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  
  // ค้นหา User (Row Index) จาก Username
  const rowIndex = data.findIndex(row => row[0] == user);
  
  if(rowIndex > 0) {
    const userData = data[rowIndex];
    const storedHash = userData[1];
    const storedSalt = userData[14] || ""; // ดึง Salt เดิม
    
    // ตรวจสอบรหัสเก่า
    if (hashPassword(oldPass, storedSalt) === storedHash) {
        // --- ถ้ารหัสเก่าถูก ให้สร้าง Salt ใหม่สำหรับรหัสใหม่ ---
        const newSalt = generateSalt();
        const newHash = hashPassword(newPass, newSalt);
        
        // อัปเดต Password
        userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
        // อัปเดต Salt
        userSheet.getRange(rowIndex + 1, 15).setValue(newSalt);
        
        return { status: 'success', message: 'เปลี่ยนรหัสเรียบร้อย' };
    }
  }
  return { status: 'error', message: 'รหัสเดิมผิด' };
}

function processForm(formData, userInfo) {
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

    let fullText = formData.reason_full || "";
    let res_1 = truncate(fullText, 40);
    fullText = fullText.substring(res_1.length);
    let res_2 = truncate(fullText, 120);
    fullText = fullText.substring(res_2.length);
    let res_3 = truncate(fullText, 120);

    let reqType = formData.request_type;
    const val = (topic, value) => (reqType === topic || (Array.isArray(topic) && topic.includes(reqType))) ? value : "";
    const replace = (key, value) => slide.replaceAllText(`{{${key}}}`, value || " ");
    const tick = "✓";
    replace('male', userInfo.gender === 'male' ? tick : "");
    replace('female', userInfo.gender === 'female' ? tick : "");
    replace('BJM', formData.program === 'BJM' ? tick : "");
    replace('Thai', formData.program === 'Thai' ? tick : "");
    for(let i=1; i<=10; i++) replace(`t${i}`, (reqType === `t${i}`) ? tick : "");

    replace('name', truncate(userInfo.name, 30));
    replace('std_id', truncate(userInfo.std_id, 10));
    replace('Year', truncate(formData.year, 1));
    replace('advisor', formData.advisor);
    replace('major', truncate(formData.major, 30)); 
    replace('address', truncate(formData.address, 95));
    replace('tel', truncate((formData.tel || "").replace(/\D/g,''), 10));
    replace('email', truncate(formData.email, 60));
    
    let specificData = "";
    specificData += truncate(val('t1', formData.major_sel), 40);
    specificData += truncate(val('t2', formData.major_from), 40) + " " + truncate(val('t2', formData.major_to), 40);
    specificData += truncate(val('t3', formData.prof_rec), 30) + " (" + truncate(val('t3', formData.r_no), 1) + ")";
    specificData += truncate(val('t5', formData.reg_sem), 1) + "/" + truncate(val('t5', formData.reg_year), 4) + " " + truncate(val('t5', formData.reg_reasson), 30);
    specificData += truncate(val('t6', formData.re_ad), 1) + "/" + truncate(val('t6', formData.re_ad_year), 4);
    specificData += truncate(val(['t7', 't8'], formData.location), 80);
    specificData += truncate(val('t9', formData.items), 80);
    specificData += truncate(val('t10', formData.other), 90);

    replace('major_sel',  truncate(val('t1', formData.major_sel), 40));
    replace('major_from', truncate(val('t2', formData.major_from), 40));
    replace('major_to',   truncate(val('t2', formData.major_to), 40));
    replace('prof_rec',   truncate(val('t3', formData.prof_rec), 30));
    replace('r_no',       truncate(val('t3', formData.r_no), 1));
    replace('reg_sem',    truncate(val('t5', formData.reg_sem), 1));
    replace('reg_year',   truncate(val('t5', formData.reg_year), 4));
    replace('reg_reasson',truncate(val('t5', formData.reg_reasson), 30));
    replace('re_ad',      truncate(val('t6', formData.re_ad), 1));
    replace('re_ad_year', truncate(val('t6', formData.re_ad_year), 4));
    replace('location',   truncate(val(['t7', 't8'], formData.location), 80));
    replace('items',      truncate(val('t9', formData.items), 80));
    replace('other',      truncate(val('t10', formData.other), 90));

    replace('res_1', res_1);
    replace('res_2', res_2);
    replace('res_3', res_3);

    slide.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
    const pdfFile = destFolder.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();
    const fileId = pdfFile.getId();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('Logs');
    if(!logSheet) { 
      logSheet = ss.insertSheet('Logs');
      logSheet.appendRow(['Timestamp', 'Username', 'File Name', 'Type', 'URL', 'File ID', 'Program', 'Gender', 'Year', 'Tel', 'Major', 'Advisor', 'Email', 'Address', 'Topic Data', 'Reason', 'Status', 'Student_File', 'Admin_File']);
    }
    
    if (logSheet.getLastColumn() < 19) {
       logSheet.insertColumnsAfter(logSheet.getLastColumn(), 19 - logSheet.getLastColumn());
    }

    logSheet.appendRow([
      new Date(), userInfo.username, fileName, reqType, pdfUrl, fileId, 
      formData.program, userInfo.gender, formData.year, "'" + formData.tel, formData.major, 
      formData.advisor, formData.email, formData.address, specificData, formData.reason_full,
      'รอ', '', '' 
    ]);
    
    // แจ้งเตือน LINE (สร้างใหม่)
    try {
        const topicMap = {
          't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา',
          't3': 'ขอหนังสือรับรองความประพฤติ', 't4': 'ขออนุมัติย้ายคณะ',
          't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา',
          't7': 'ขอใช้ห้องเรียน / สถานที่', 't8': 'ขออนุญาตใช้ห้อง',
          't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ'
        };
        const topicName = topicMap[reqType] || reqType;
        const lineMsg = `🔔 มีคำร้องใหม่!\n` +
                        `👤 ชื่อ: ${userInfo.name} (${userInfo.std_id})\n` +
                        `📝 เรื่อง: ${topicName}\n` +
                        `📂 PDF: ${pdfUrl}`;
        sendLinePushMessage(lineMsg);

    } catch(err) {
        console.log("LINE Alert Error: " + err);
    }

    return { status: 'success', url: pdfUrl };
  } catch (e) { return { status: 'error', message: 'Error: ' + e.toString() };
  }
}

function getRequestsData(user) {
  if (!user || !user.username) return [];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. ดึงข้อมูล Users มาทำ Map เพื่อหาชื่อและรหัสนักศึกษา
  const userSheet = ss.getSheetByName('Users');
  let userMap = {};
  if (userSheet) {
     const uData = userSheet.getDataRange().getValues();
     // สมมติ: col 0=User, col 2=Name, col 3=Std_ID
     uData.forEach(r => {
        if(r[0]) userMap[r[0]] = { name: r[2], std_id: r[3] };
     });
  }

  // 2. ดึงข้อมูล Logs
  let sheet = ss.getSheetByName('Logs');
  if(!sheet || sheet.getLastRow() < 2) return [];
  
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  
  let requests = data.map(r => {
    try {
        let ts = r[0];
        let timeStr = "-";
        if (ts instanceof Date) {
            timeStr = Utilities.formatDate(ts, "GMT+7", "dd/MM/yyyy HH:mm");
        } else {
            timeStr = String(ts || "-");
        }
        
        let username = String(r[1] || "");
        let userInfo = userMap[username] || { name: "-", std_id: "-" };

        return {
            timestamp: timeStr,
            username: username,
            name: String(userInfo.name),       // เพิ่มชื่อ
            std_id: String(userInfo.std_id),   // เพิ่มรหัสนักศึกษา
            fileName: String(r[2] || "ไม่มีชื่อไฟล์"),
            type: String(r[3] || ""),
            pdfUrl: String(r[4] || "#"),
            fileId: String(r[5] || ""),
            status: (r.length > 16) ? String(r[16] || "รอ") : "รอ",
            studentFile: (r.length > 17) ? String(r[17] || "") : "",
            adminFile: (r.length > 18) ? String(r[18] || "") : ""
        };
    } catch (err) {
        return null;
    }
  }).filter(item => item !== null);
  
  if (user.role !== 'admin') {
    requests = requests.filter(r => r.username === user.username);
  }
  
  return requests.reverse();
}

function uploadFile(base64Data, fileType, relatedFileId, uploaderRole, username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Logs');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[5] === relatedFileId);
    if (rowIndex <= 0) return { status: 'error', message: 'ไม่พบรายการ' };

    const splitBase = base64Data.split(',');
    const blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), fileType, `Upload_${new Date().getTime()}`);
    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    if (sheet.getLastColumn() < 19) sheet.insertColumnsAfter(sheet.getLastColumn(), 19 - sheet.getLastColumn());

    if (uploaderRole === 'admin') {
      // Admin อัปโหลดไฟล์ (เช่น ไฟล์ที่เซ็นแล้ว)
      sheet.getRange(rowIndex + 1, 19).setValue(fileUrl);
      sheet.getRange(rowIndex + 1, 17).setValue('เสร็จสิ้น');
    } else {
      // นักศึกษาอัปโหลดไฟล์
      if (data[rowIndex][1] !== username) return { status: 'error', message: 'ไม่มีสิทธิ์' };
      
      sheet.getRange(rowIndex + 1, 18).setValue(fileUrl);
      sheet.getRange(rowIndex + 1, 17).setValue('รอเจ้าหน้าที่ตรวจสอบ'); // เปลี่ยนสถานะตาม Requirement

      // --- แจ้งเตือน LINE เมื่อนักศึกษาส่งไฟล์ ---
      try {
        const topicMap = {
          't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา',
          't3': 'ขอหนังสือรับรองความประพฤติ', 't4': 'ขออนุมัติย้ายคณะ',
          't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา',
          't7': 'ขอใช้ห้องเรียน / สถานที่', 't8': 'ขออนุญาตใช้ห้อง',
          't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ'
        };
        const r = data[rowIndex];
        const reqType = r[3];
        const topicName = topicMap[reqType] || reqType;
        
        // ดึงชื่อนักศึกษา (ต้องหาจาก Users เหมือนเดิม หรือถ้าไม่มีก็ใช้ username ไปก่อน)
        const userSheet = ss.getSheetByName('Users');
        const userData = userSheet.getDataRange().getValues();
        const userObj = userData.find(u => u[0] === username);
        const nameShow = userObj ? `${userObj[2]} (${userObj[3]})` : username;

        const lineMsg = `🔄 Updated นักศึกษาส่งไฟล์ใหม่แล้ว!\n` +
                        `กำลังรอรับเจ้าหน้าที่รับเรื่อง\n` +
                        `👤 จาก: ${nameShow}\n` +
                        `📝 เรื่อง: ${topicName}\n` +
                        `📂 ไฟล์แนบ: ${fileUrl}`;
        sendLinePushMessage(lineMsg);
      } catch(err) {
        console.log("LINE Update Error: " + err);
      }
    }

    return { status: 'success', message: 'อัปโหลดสำเร็จ' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function adminUpdateStatus(fileId, newStatus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[5] === fileId);
  if (rowIndex > 0) {
    if (sheet.getLastColumn() < 17) sheet.insertColumnsAfter(sheet.getLastColumn(), 17 - sheet.getLastColumn());
    sheet.getRange(rowIndex + 1, 17).setValue(newStatus);
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
  const currentUser = userRows.find(row => row[0] === username);
  const isAdmin = currentUser && currentUser[12] === 'admin'; 

  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[5] === fileId && (r[1] === username || isAdmin));
  if(rowIndex > 0) { 
      try { DriveApp.getFileById(fileId).setTrashed(true);
      } catch(e){}
      sheet.deleteRow(rowIndex + 1); 
      return { status: 'success', message: 'ลบเรียบร้อย' };
  }
  
  return { status: 'error', message: 'เกิดข้อผิดพลาด: คุณไม่มีสิทธิ์ลบไฟล์นี้ หรือไม่พบไฟล์' };
}

function renameHistory(fileId, newName, username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[5] === fileId && r[1] === username);
  if(rowIndex > 0) {
      try { DriveApp.getFileById(fileId).setName(newName);
      } catch(e){}
      sheet.getRange(rowIndex + 1, 3).setValue(newName);
      return { status: 'success', message: 'แก้ไขเรียบร้อย' };
  }
  return { status: 'error', message: 'Error' };
}

function truncate(text, limit) {
  if (!text) return "";
  text = String(text);
  const getVisualLen = (t) => t.replace(/[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/g, '').length;
  let current = "";
  for (let char of text) {
    if (getVisualLen(current + char) > limit) break;
    current += char;
  }
  return current;
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
  if (!configSheet) { configSheet = ss.insertSheet('Config');
  configSheet.appendRow(['Major', 'Advisor']); }

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
    templates = d.map(r => ({
      name: r[0], topic: r[1], data: r[2], reason: r[3]
    })).filter(t => t.name);
  }
  return { majors, teachers, templates };
}

function sendLinePushMessage(message) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = sheet.getSheetByName("Config_Line"); 

    if (!configSheet) {
      console.log("❌ ไม่พบ Sheet 'Config_Line'");
      return;
    }

    var accessToken = configSheet.getRange("B1").getValue();
    var targetId = configSheet.getRange("B2").getValue();
    if (!accessToken || !targetId) {
      console.log("❌ ข้อมูล Token (B1) หรือ ID (B2) ไม่ครบ");
      return;
    }

    var url = "https://api.line.me/v2/bot/message/push";
    var payload = JSON.stringify({
      "to": targetId,
      "messages": [{ "type": "text", "text": message }]
    });
    UrlFetchApp.fetch(url, {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + accessToken
      },
      "payload": payload
    });
    console.log("✅ ส่ง LINE สำเร็จ");

  } catch (e) {
    console.log("❌ Error sendLine: " + e.toString());
  }
}

function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    if (json.events.length === 0) return;
    var event = json.events[0];
    var msg = event.message.text || "";
    var type = event.source.type;
    var id = "";
    if (type === "group") {
      id = event.source.groupId;
    } else {
      id = event.source.userId;
    }

    if (msg.toLowerCase().includes("check")) { 
       MailApp.sendEmail({
         to: "nitichan@tu.ac.th",
         subject: "📌 ได้ Group ID แล้วครับ!",
         htmlBody: "<h3>ข้อมูลจาก LINE (" + type + ")</h3>" +
                   "<p><b>Group ID / User ID คือ:</b></p>" +
                   "<h2>" 
                   + id + "</h2>" +
                   "<hr>" +
                   "<p>ก๊อปปี้รหัสนี้ไปใส่ในตัวแปร <b>ADMIN_USER_ID</b> ในไฟล์ Code.gs ได้เลยครับ</p>"
       });
    }

  } catch (error) {
    MailApp.sendEmail({
       to: "nitichan@tu.ac.th",
       subject: "❌ ระบบ Error",
       body: "Error: " + error.toString()
    });
  }
}

function replyLineMessage(replyToken, id, typeText, token) {
  var url = "https://api.line.me/v2/bot/message/reply";
  var payload = JSON.stringify({
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": typeText + " ของคุณคือ:\n" + id + "\n\n(นำรหัสนี้ไปใส่ในช่อง B2 ของ Sheet 'Config_Line')"
    }]
  });
  UrlFetchApp.fetch(url, {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    "payload": payload
  });
}

function testPushSystem() {
  console.log("🚀 กำลังทดสอบส่งข้อความ...");
  sendLinePushMessage("🟢 ทดสอบระบบ: การเชื่อมต่อสำเร็จ! (จาก Admin)");
}
