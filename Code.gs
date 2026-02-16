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
function generateToken() { return Utilities.getUuid();
}

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
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Name', 'Std_ID', 'Email', 'Tel', 'Year', 'Gender', 'Token', 'Verified', 'Reset_Token', 'Reset_Exp', 'Role', 'Status', 'Salt']);
    return { status: 'error', message: 'ระบบเพิ่งเริ่มต้น กรุณาสมัครสมาชิกใหม่' };
  }

  const data = userSheet.getDataRange().getValues();
  const userRow = data.find(row => row[0] == username);
  if (userRow) {
    if (String(userRow[9]).toUpperCase() !== 'TRUE') {
      return { status: 'error', message: 'กรุณายืนยันตัวตนทาง Email ก่อน' };
    }
    
    let role = (userRow.length > 12 && userRow[12]) ? userRow[12] : 'user';
    let status = (userRow.length > 13 && userRow[13]) ? userRow[13] : 'active';
    if (String(status).toLowerCase() === 'banned') {
      return { status: 'error', message: 'บัญชีของคุณถูกระงับการใช้งาน' };
    }

    const storedHash = userRow[1];
    const storedSalt = userRow[14] || "";
    if (hashPassword(password, storedSalt) === storedHash) {
        return { 
          status: 'success', 
          username: String(userRow[0]), 
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
    salt
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
    const storedHash = userData[1];
    const storedSalt = userData[14] || ""; 
    
    if (hashPassword(oldPass, storedSalt) === storedHash) {
        const newSalt = generateSalt();
        const newHash = hashPassword(newPass, newSalt);
        userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
        userSheet.getRange(rowIndex + 1, 15).setValue(newSalt);
        return { status: 'success', message: 'เปลี่ยนรหัสเรียบร้อย' };
    }
  }
  return { status: 'error', message: 'รหัสเดิมผิด' };
}

// --- 🔥 Modified processForm to Support PDF Attachment 🔥 ---
async function processForm(formData, userInfo) {
  try {
    const destFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const templateFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
    
    // ตั้งชื่อไฟล์
    let fileName = formData.custom_filename || `Request_${userInfo.std_id}_${new Date().getTime()}`;
    // 1. สร้างไฟล์ PDF หลักจาก Template
    const copyFile = templateFile.makeCopy(fileName, destFolder);
    const copyId = copyFile.getId();
    const slide = SlidesApp.openById(copyId);

    // ใส่ลายเซ็น
    if (formData.signature_data) {
      const firstSlide = slide.getSlides()[0];
      replaceTextWithImage(firstSlide, '{{signature}}', formData.signature_data);
    }

    // แทนที่ข้อมูลต่างๆ
    let fullText = formData.reason_full || "";
    // 🔥 ปรับลด Limit ลง เพื่อแก้ปัญหาตกขอบ 🔥
    // บรรทัดแรกเหลือ 35, บรรทัดที่เหลือเหลือ 105
    let res_1 = truncate(fullText, 35);
    fullText = fullText.substring(res_1.length);
    let res_2 = truncate(fullText, 105);
    fullText = fullText.substring(res_2.length);
    let res_3 = truncate(fullText, 105);

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
    // Fix specificData
    let specificData = "";
    if (reqType === 't1') specificData = truncate(formData.major_sel, 40);
    else if (reqType === 't2') specificData = `${truncate(formData.major_from, 40)} ไปยัง ${truncate(formData.major_to, 40)}`;
    else if (reqType === 't3') specificData = `${truncate(formData.prof_rec, 30)} (${truncate(formData.r_no, 1)})`;
    else if (reqType === 't5') specificData = `${truncate(formData.reg_sem, 1)}/${truncate(formData.reg_year, 4)} ${truncate(formData.reg_reasson, 30)}`;
    else if (reqType === 't6') specificData = `${truncate(formData.re_ad, 1)}/${truncate(formData.re_ad_year, 4)}`;
    else if (reqType === 't7' || reqType === 't8') specificData = truncate(formData.location, 80);
    else if (reqType === 't9') specificData = truncate(formData.items, 80);
    else if (reqType === 't10') specificData = truncate(formData.other, 90);
    
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

    // แปลงไฟล์หลักเป็น PDF Blob
    let mainPdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
    // ลบไฟล์ Slide ชั่วคราวทิ้ง
    DriveApp.getFileById(copyId).setTrashed(true);

    // 2. ตรวจสอบและรวมไฟล์แนบ (ถ้ามี)
    let finalPdfBlob = mainPdfBlob;
    if (formData.fileAttachment) {
      try {
        const attachmentBlob = Utilities.newBlob(
          Utilities.base64Decode(formData.fileAttachment.content),
          formData.fileAttachment.mimeType,
          formData.fileAttachment.name
        );
        // 🔥 FIX: ตรวจสอบและเลือกวิธีการเรียกใช้ PDFApp ให้รองรับทั้ง Library และ Local File 🔥
        let mergedBlob;
        if (typeof PDFApp !== 'undefined' && PDFApp.mergePDFs) {
            // กรณีใช้ Library
            mergedBlob = await PDFApp.mergePDFs([mainPdfBlob, attachmentBlob]);
        } else if (typeof mergePDFs === 'function') {
            // กรณี Copy โค้ดลงไฟล์ (Local)
            mergedBlob = await mergePDFs([mainPdfBlob, attachmentBlob]);
        } else {
            throw new Error("ไม่พบฟังก์ชัน PDFApp.mergePDFs หรือ mergePDFs");
        }

        // ตรวจสอบว่าผลลัพธ์เป็น Blob จริงหรือไม่
        if (mergedBlob && typeof mergedBlob.setName === 'function') {
            finalPdfBlob = mergedBlob;
        } else {
            console.error("Merge returned invalid object:", mergedBlob);
        }
        
      } catch (mergeErr) {
        console.log("Merge Error: " + mergeErr);
      }
    }

    // 3. บันทึกไฟล์สุดท้ายลง Drive
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
    
    if (logSheet.getLastColumn() < 21) {
       logSheet.insertColumnsAfter(logSheet.getLastColumn(), 21 - logSheet.getLastColumn());
    }

    let displayName = userInfo.name;
    if (displayName && userInfo.gender && !displayName.startsWith('นาย') && !displayName.startsWith('นาง')) {
        displayName = (userInfo.gender === 'male' ? 'นาย' : 'นางสาว') + displayName;
    }

    const rawDataJSON = JSON.stringify(formData);
    
    logSheet.appendRow([
      new Date(), 
      String(userInfo.username), 
      displayName, 
      fileName, reqType, pdfUrl, fileId, 
      formData.program, userInfo.gender, formData.year, "'" + formData.tel, formData.major, 
      formData.advisor, formData.email, formData.address, specificData, formData.reason_full,
      'รอ', '', '', rawDataJSON 
    ]);
    try {
        const topicMap = {
          't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา',
          't3': 'ขอหนังสือรับรองความประพฤติ', 't4': 'ขออนุมัติย้ายคณะ',
          't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา',
          't7': 'ขอใช้ห้องเรียน / สถานที่', 't8': 'ขออนุญาตใช้ห้อง',
          't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ'
        };
        const topicName = topicMap[reqType] || reqType;
        const lineMsg = `🔔 มีคำร้องใหม่ (พร้อมไฟล์แนบ)!\n` +
                        `👤 ชื่อ: ${displayName} (${userInfo.std_id})\n` +
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
  
  const userSheet = ss.getSheetByName('Users');
  let userMap = {};
  if (userSheet) {
     const uData = userSheet.getDataRange().getValues();
     uData.forEach(r => {
        if(r[0]) userMap[String(r[0])] = { name: r[2], std_id: r[3], gender: r[7] };
     });
  }

  let sheet = ss.getSheetByName('Logs');
  if(!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  
  let requests = data.map(r => {
    try {
        let ts = r[0];
        let timeStr = (ts instanceof Date) ? Utilities.formatDate(ts, "GMT+7", "dd/MM/yyyy HH:mm") : String(ts || "-");
        let username = String(r[1] || ""); 
        let logName = String(r[2] || "");
        let userInfo = userMap[username] || { name: "-", std_id: "-", gender: "" };
 
        let finalName = logName;
        if (!finalName || finalName === "" || finalName === "-") {
            finalName = String(userInfo.name);
            if (finalName !== "-" && userInfo.gender && !finalName.startsWith('นาย') && !finalName.startsWith('นาง')) {
                finalName = (String(userInfo.gender).toLowerCase() === 'male' ? 'นาย' : 'นางสาว') + finalName;
          
             }
        }
   
        let rawData = {};
        try {
            if(r[20] && r[20] !== "") rawData = JSON.parse(r[20]);
        } catch(e) {}

        return {
            timestamp: timeStr,
            username: username,
            name: finalName, 
            std_id: String(userInfo.std_id),
            fileName: String(r[3] || "ไม่มีชื่อไฟล์"), 
            type: String(r[4] || ""),   
            pdfUrl: String(r[5] || "#"),  
            fileId: String(r[6] || ""),   
            program: String(r[7] || ""),
            year: String(r[9] || ""),     
            tel: String(r[10] || "").replace(/'/g, ''),
            major: String(r[11] || ""),
            advisor: String(r[12] || ""),
            email: String(r[13] || ""),
            address: String(r[14] || ""),
            reason: String(r[16] || ""), 
            status: (r.length > 17) ? String(r[17] || "รอ") : "รอ", 
            studentFile: (r.length > 18) ? String(r[18] || "") : "", 
            adminFile: (r.length > 19) ? String(r[19] || "") : "",   
            ...rawData
        };
    } catch (err) {
        return null;
    }
  }).filter(item => item !== null);
  if (user.role !== 'admin') {
    requests = requests.filter(r => r.username === String(user.username));
  }
  
  return requests.reverse();
}

async function uploadFile(base64Data, fileType, relatedFileId, uploaderRole, username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Logs');
    const data = sheet.getDataRange().getValues();
    
    // 1. หาแถวของรายการที่เกี่ยวข้อง
    const rowIndex = data.findIndex(row => row[6] === relatedFileId);
    if (rowIndex <= 0) return { status: 'error', message: 'ไม่พบรายการ' };
    // 2. เตรียมไฟล์ PDF Blob
    const splitBase = base64Data.split(',');
    const decoded = Utilities.base64Decode(splitBase[1]);
    let uploadBlob = Utilities.newBlob(decoded, fileType, `Upload_${new Date().getTime()}.pdf`);
    
    // ประกาศตัวแปร Timestamp ให้มองเห็นทั่วฟังก์ชัน
    let timestampText = "Received: " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    // 3. ใช้ PDFApp ประทับเวลา (ถ้าเป็น PDF)
    if (fileType === 'application/pdf' && typeof PDFApp !== 'undefined') {
       try {
         const newPdfBlob = await PDFApp.setPDFBlob(uploadBlob)
           .insertHeaderFooter({
              header: {
                left: { 
                  text: timestampText,   
                  size: 3,
                  x: 20,                 
                  yOffset: 10            
                }
              }
           });
           if (newPdfBlob) {
            uploadBlob = newPdfBlob;
            uploadBlob.setName(`Upload_${new Date().getTime()}.pdf`);
           }
       } catch (e) {
         console.log("PDFApp Stamp Error: " + e.toString());
       }
    }

    // 4. บันทึกไฟล์ "ใหม่" ลง Google Drive
    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const file = folder.createFile(uploadBlob);
    const fileUrl = file.getUrl();

    // ขยายตารางถ้าคอลัมน์ไม่พอ
    if (sheet.getLastColumn() < 21) sheet.insertColumnsAfter(sheet.getLastColumn(), 21 - sheet.getLastColumn());
    // 5. อัปเดตวันที่และเวลา
    sheet.getRange(rowIndex + 1, 1).setValue(new Date());
    // 6. อัปเดตสถานะและแจ้งเตือน (พร้อมลบไฟล์เก่า)
    if (uploaderRole === 'admin') {
      // --- ลบไฟล์เก่าของ Admin (ถ้ามี) ---
      const oldAdminUrl = data[rowIndex][19];
      // Col 20 (Index 19)
      if (oldAdminUrl && String(oldAdminUrl).includes('drive.google.com')) {
          try {
             const match = String(oldAdminUrl).match(/[-\w]{25,}/);
             if (match) DriveApp.getFileById(match[0]).setTrashed(true);
          } catch(e) { console.log("Failed to delete old Admin file: " + e);
          }
      }
      // -------------------------------

      sheet.getRange(rowIndex + 1, 20).setValue(fileUrl);
      sheet.getRange(rowIndex + 1, 18).setValue('เสร็จสิ้น');

    } else {
      if (String(data[rowIndex][1]) !== String(username)) return { status: 'error', message: 'ไม่มีสิทธิ์' };
      // --- ลบไฟล์เก่าของนักศึกษา (ถ้ามี) ---
      const oldStudentUrl = data[rowIndex][18];
      // Col 19 (Index 18)
      if (oldStudentUrl && String(oldStudentUrl).includes('drive.google.com')) {
          try {
             const match = String(oldStudentUrl).match(/[-\w]{25,}/);
             if (match) DriveApp.getFileById(match[0]).setTrashed(true);
          } catch(e) { console.log("Failed to delete old Student file: " + e);
          }
      }
      // ---------------------------------

      sheet.getRange(rowIndex + 1, 19).setValue(fileUrl);
      sheet.getRange(rowIndex + 1, 18).setValue('รอเจ้าหน้าที่ตรวจสอบ'); 

      // ส่ง LINE Notify
      try {
        const topicMap = {
          't1': 'เลือกเรียนกลุ่มวิชา', 't2': 'ขอเปลี่ยนกลุ่มวิชา',
          't3': 'ขอหนังสือรับรองความประพฤติ', 't4': 'ขออนุมัติย้ายคณะ',
          't5': 'ขอลาออก', 't6': 'ขอคืนสภาพนักศึกษา',
          't7': 'ขอใช้ห้องเรียน / สถานที่', 't8': 'ขออนุญาตใช้ห้อง',
          't9': 'ขอยืมอุปกรณ์', 't10': 'อื่นๆ'
        };
        const r = data[rowIndex];
        const reqType = r[4]; 
        const topicName = topicMap[reqType] || reqType;
        const nameShow = r[2] || username;
        
        const lineMsg = `🔄 Updated นักศึกษาส่งไฟล์ใหม่แล้ว!\n` +
                        `กำลังรอรับเจ้าหน้าที่รับเรื่อง\n` +
                        `👤 จาก: ${nameShow}\n` +
                        `📝 เรื่อง: ${topicName}\n` +
                        `⏱️ ส่งเมื่อ: ${timestampText}\n` +
                        `📂 ไฟล์แนบ: ${fileUrl}`;
        sendLinePushMessage(lineMsg);
      } catch(err) {
        console.log("LINE Update Error: " + err);
      }
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
  // ตรวจสอบสิทธิ์ (User หรือ Admin)
  const userSheet = ss.getSheetByName('Users');
  const userRows = userSheet.getDataRange().getValues();
  const currentUser = userRows.find(row => String(row[0]) === String(username));
  const isAdmin = currentUser && currentUser[12] === 'admin'; 

  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  
  // ค้นหาแถวที่ต้องการลบ
  const rowIndex = data.findIndex(r => r[6] === fileId && (String(r[1]) === String(username) || isAdmin));
  if(rowIndex > 0) { 
      const rowData = data[rowIndex];
      // --- 1. ลบไฟล์ต้นฉบับ (Main File) ---
      try { 
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch(e) { console.log("Delete Main File Error: " + e);
      }

      // --- 2. ลบไฟล์แนบเพิ่มของนักศึกษา (Student_File: Col 19 / Index 18) ---
      if (rowData[18] && String(rowData[18]).includes('drive.google.com')) {
          try {
             // ดึง ID ออกมาจาก URL
             const match = String(rowData[18]).match(/[-\w]{25,}/);
             if (match) DriveApp.getFileById(match[0]).setTrashed(true);
          } catch(e) { console.log("Delete Student File Error: " + e);
          }
      }

      // --- 3. ลบไฟล์แนบของ Admin (Admin_File: Col 20 / Index 19) ---
      if (rowData[19] && String(rowData[19]).includes('drive.google.com')) {
          try {
             const match = String(rowData[19]).match(/[-\w]{25,}/);
             if (match) DriveApp.getFileById(match[0]).setTrashed(true);
          } catch(e) { console.log("Delete Admin File Error: " + e);
          }
      }

      // --- 4. ลบแถวใน Google Sheets ---
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
      try { DriveApp.getFileById(fileId).setName(newName);
      } catch(e){}
      sheet.getRange(rowIndex + 1, 4).setValue(newName);
      return { status: 'success', message: 'แก้ไขเรียบร้อย' };
  }
  return { status: 'error', message: 'Error' };
}

// 🔥 MODIFIED: Weighted Truncate Function (Tuned for Safety) 🔥
function truncate(text, limit) {
  if (!text) return "";
  text = String(text);
  
  let currentWidth = 0;
  let result = "";
  
  for (let char of text) {
    let w = 1.0;
    let c = char;
    let code = c.charCodeAt(0);
    
    // Logic การคำนวณความกว้าง (Updated Weights)
    if ((code >= 0x0E31 && code <= 0x0E3A) || (code >= 0x0E47 && code <= 0x0E4E)) {
      w = 0.0; // สระบนล่าง (ไม่กินที่)
    } else if (["ณ", "ญ", "ฌ", "ฒ", "ฑ", "ฤ", "ฦ", "W", "M", "m", "w"].includes(c)) {
      w = 2.0; // 🔥 เพิ่มจาก 1.8 เป็น 2.0 (ให้กินที่เยอะขึ้น จะได้ตัดเร็วขึ้น)
    } else if (["เ", "แ", "ไ", "ใ", "โ", "า", "i", "l", "I", "1", "j", "f", "|", ".", ","].includes(c)) {
      w = 0.7; // 🔥 เพิ่มจาก 0.6 เป็น 0.7
    }
    
    if (currentWidth + w > limit) break;
    result += c;
    currentWidth += w;
  }
  
  return result;
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
                   "<h2>" + id + "</h2>" +
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
