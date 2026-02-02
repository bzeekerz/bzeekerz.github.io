// --- CONFIGURATION ---
const SPREADSHEET_ID = '1u8OaGgDcpgWdtaqTXpwWm8PX2b4I2Ovq93aKRuXol18';
const TEMPLATE_SLIDE_ID = '1FEVxooVLLEmxUscy6dXiPZHPjqMn8Bu7NEAXdQ19k-w';
const DESTINATION_FOLDER_ID = '1u1LpLsCDaUgwWYJIXn5L9D_a1sBhKoU7';

// --- ROUTING & INIT ---
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.urlParams = JSON.stringify(e.parameter);
  template.serverMessage = "";
  template.serverStatus = "";

  if (e.parameter.page === 'verify' && e.parameter.token) {
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

function hashPassword(password) {
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawBytes.length; i++) {
    let hashVal = rawBytes[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function sendEmail(to, subject, body) {
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: body });
}

// --- USER MANAGEMENT ---
function registerUser(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { status: 'error', message: 'Database Error: ไม่พบ Sheet Users' };
  
  const data = userSheet.getDataRange().getValues();
  if (data.some(row => row[0] === formObject.reg_username)) return { status: 'error', message: 'Username นี้ถูกใช้ไปแล้ว' };
  if (data.some(row => row[4] === formObject.reg_email)) return { status: 'error', message: 'Email นี้ถูกใช้ไปแล้ว' };

  const hashedPassword = hashPassword(formObject.reg_password);
  const verifyToken = generateToken();
  const verifyLink = `${getScriptUrl()}?page=verify&token=${verifyToken}`;

  // บันทึกข้อมูล (เพิ่ม Gender ที่ index 7 / Col H)
  userSheet.appendRow([
    formObject.reg_username, 
    hashedPassword, 
    formObject.reg_name, 
    formObject.reg_std_id,
    formObject.reg_email, 
    formObject.reg_tel,   
    formObject.reg_year,
    formObject.reg_gender, // <-- เก็บเพศตรงนี้
    verifyToken, // Col I (Index 8)
    'FALSE',     // Col J (Index 9)
    '',          // Col K (Index 10)
    ''           // Col L (Index 11)
  ]);
  
  try {
    sendEmail(
      formObject.reg_email,
      'ยืนยันการสมัครสมาชิก JC Request Form',
      `<h2>สวัสดีคุณ ${formObject.reg_name}</h2>
       <p>กรุณาคลิกลิงก์ด้านล่างเพื่อยืนยันบัญชีของคุณ:</p>
       <p><a href="${verifyLink}">คลิกเพื่อยืนยันตัวตน</a></p>`
    );
    return { status: 'success', message: 'สมัครสำเร็จ! กรุณาตรวจสอบ Email เพื่อยืนยันตัวตน' };
  } catch (e) {
    return { status: 'error', message: 'ส่งเมลไม่ผ่าน: ' + e.toString() };
  }
}

function verifyUserToken(token) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[8] === token); // Token อยู่ Col I (index 8)
  
  if (rowIndex > 0) {
    userSheet.getRange(rowIndex + 1, 9).setValue(''); // Clear Token (Col I)
    userSheet.getRange(rowIndex + 1, 10).setValue('TRUE'); // Set Verified (Col J)
    return { status: 'success', message: 'ยืนยันตัวตนสำเร็จ! เข้าสู่ระบบได้เลย' };
  }
  return { status: 'error', message: 'ลิงก์ยืนยันไม่ถูกต้อง หรือถูกใช้งานไปแล้ว' };
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const inputHash = hashPassword(password);
  
  const user = data.find(row => row[0] == username && row[1] == inputHash);
  
  if (user) {
    if (String(user[9]).toUpperCase() !== 'TRUE') {
      return { status: 'error', message: 'กรุณายืนยันตัวตนทาง Email ก่อน' };
    }
    return { 
      status: 'success', 
      username: user[0], 
      name: user[2], 
      std_id: user[3],
      email: user[4], 
      tel: user[5],
      year: user[6],
      gender: user[7] // <-- ส่งเพศกลับไปให้หน้าเว็บใช้งาน
    };
  } else {
    return { status: 'error', message: 'Username หรือ Password ไม่ถูกต้อง' };
  }
}

function requestPasswordReset(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[4] === email);
  
  if (rowIndex > 0) {
    const token = generateToken();
    const expiry = new Date().getTime() + (3600 * 1000); 
    const resetLink = `${getScriptUrl()}?page=reset&token=${token}`;
    
    userSheet.getRange(rowIndex + 1, 11).setValue(token);
    userSheet.getRange(rowIndex + 1, 12).setValue(expiry);
    
    try {
      sendEmail(email, 'แจ้งลืมรหัสผ่าน', `<p><a href="${resetLink}">ตั้งรหัสผ่านใหม่คลิกที่นี่</a></p>`);
    } catch(e) {}
  }
  return { status: 'success', message: 'หากอีเมลถูกต้อง ระบบส่งลิงก์เปลี่ยนรหัสให้แล้ว' };
}

function submitResetPassword(token, newPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[10] === token);
  
  if (rowIndex > 0) {
    const expiry = data[rowIndex][11]; 
    if (new Date().getTime() > expiry) return { status: 'error', message: 'ลิงก์หมดอายุ' };
    
    const newHash = hashPassword(newPassword);
    userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
    userSheet.getRange(rowIndex + 1, 11).setValue('');
    userSheet.getRange(rowIndex + 1, 12).setValue('');
    return { status: 'success', message: 'เปลี่ยนรหัสสำเร็จ!' };
  }
  return { status: 'error', message: 'Token ไม่ถูกต้อง' };
}

function changePassword(username, oldPass, newPass) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const oldHash = hashPassword(oldPass);
  const rowIndex = data.findIndex(row => row[0] == username && row[1] == oldHash);
  
  if (rowIndex > 0) {
    const newHash = hashPassword(newPass);
    userSheet.getRange(rowIndex + 1, 2).setValue(newHash);
    return { status: 'success', message: 'เปลี่ยนรหัสผ่านเรียบร้อย' };
  }
  return { status: 'error', message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
}

// --- FORM PROCESSING & PDF ---
function processForm(formData, userInfo) {
  try {
    const destFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const templateFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
    
    let fileName = formData.custom_filename;
    if (!fileName || fileName.trim() === "") {
      fileName = `Request_${userInfo.std_id}_${new Date().getTime()}`;
    }

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

    // Log to Sheet 'Logs'
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('Logs');
    if(!logSheet) { 
      logSheet = ss.insertSheet('Logs'); 
      logSheet.appendRow([
        'Timestamp', 'Username', 'File Name', 'Type', 'URL', 'File ID', 
        'Program', 'Gender', 'Year', 'Tel', 'Major', 'Advisor', 'Email', 'Address', 'Topic Data', 'Reason'
      ]); 
    }
    
    logSheet.appendRow([
      new Date(),           
      userInfo.username,    
      fileName,             
      reqType,              
      pdfUrl,               
      fileId,               
      formData.program,     
      userInfo.gender,      
      formData.year,        
      formData.tel,         
      formData.major,       
      formData.advisor,     
      formData.email,       
      formData.address,     
      specificData,         
      formData.reason_full  
    ]);

    return { status: 'success', url: pdfUrl };
  } catch (e) {
    return { status: 'error', message: 'PDF Error: ' + e.toString() };
  }
}

// --- UTILS ---
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

// --- DATA HELPERS ---
function getUserHistory(username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  
  if(!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues(); 

  return data
    .filter(r => r[1] == username)
    .map(r => {
      let timeStr = "-";
      try {
        if (r[0] instanceof Date) {
          timeStr = Utilities.formatDate(r[0], "GMT+7", "dd/MM/yyyy HH:mm");
        } else {
          timeStr = String(r[0]); 
        }
      } catch (e) { console.log(e); }

      return { 
        timestamp: timeStr, 
        fileName: r[2], 
        type: r[3], 
        url: r[4],
        fileId: r[5]    
      };
    }).reverse();
}

function deleteHistory(fileId, username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[5] === fileId && r[1] === username);

  if (rowIndex > 0) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
    sheet.deleteRow(rowIndex + 1);
    return { status: 'success', message: 'ลบข้อมูลเรียบร้อยแล้ว' };
  }
  return { status: 'error', message: 'ไม่พบข้อมูล หรือคุณไม่มีสิทธิ์ลบ' };
}

function renameHistory(fileId, newName, username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[5] === fileId && r[1] === username);

  if (rowIndex > 0) {
    let finalName = newName;
    try {
      const file = DriveApp.getFileById(fileId);
      file.setName(finalName);
    } catch(e) { return { status: 'error', message: 'ไม่พบไฟล์ใน Drive' }; }

    sheet.getRange(rowIndex + 1, 3).setValue(finalName);
    return { status: 'success', message: 'เปลี่ยนชื่อเรียบร้อยแล้ว' };
  }
  return { status: 'error', message: 'ไม่พบข้อมูล หรือคุณไม่มีสิทธิ์แก้ไข' };
}

function getTemplateData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  let majors = [], teachers = [];
  if (configSheet && configSheet.getLastRow() > 1) {
    const d = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
    majors = d.map(r => r[0]).filter(String);
    teachers = d.map(r => r[1]).filter(String);
  } else {
     majors = ["ไม่พบข้อมูล Config"]; teachers = ["ไม่พบข้อมูล Config"];
  }

  const tempSheet = ss.getSheetByName('Templates');
  let templates = [];
  if (tempSheet && tempSheet.getLastRow() > 1) {
    const d = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, 4).getValues();
    templates = d.map(r => ({
      name: r[0], topic: r[1], data: r[2], reason: r[3]
    })).filter(t => t.name);
  }
  return { majors, teachers, templates };
}
