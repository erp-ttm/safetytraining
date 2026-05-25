// ==========================================
// TTM Safety Training System - Code.gs
// Google Apps Script Backend
// ==========================================

const SPREADSHEET_ID = '1flL8UacixiOags5vufPuWK--z6waWtTXdM3WBeJ7r6c'; // <-- ใส่ ID ของ Google Sheet
const SLIDES_TEMPLATE_ID = '1gO7YG67uSU1Koud5xhsVGNht8_hcIET0HAH_cURfUcA'; // <-- ใส่ ID ของ Google Slides Template
const DRIVE_FOLDER_ID = '15YaDqlNPExyM4wCnNQtfwXCAWaytuNcu'; // <-- ใส่ ID ของ Google Drive Folder

// ==========================================
// MAIN ENTRY POINT
// ==========================================
function doGet(e) {
  const page = e.parameter.page || 'index';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle('Admin - TTM Safety Training')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'user') {
    return HtmlService.createTemplateFromFile('User')
      .evaluate()
      .setTitle('Training - TTM Safety Training')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TTM Safety Training System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// AUTHENTICATION
// ==========================================
function login(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === username && row[1] === password && row[3] === true) {
        return {
          success: true,
          role: row[2], // 'admin' or 'user'
          username: username,
          displayName: row[4] || username
        };
      }
    }
    return { success: false, message: 'Invalid username or password' };
  } catch (e) {
    return { success: false, message: 'System error: ' + e.message };
  }
}

// ==========================================
// USER MANAGEMENT (Admin)
// ==========================================
function getUsers() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        users.push({
          row: i + 1,
          username: data[i][0],
          password: data[i][1],
          role: data[i][2],
          active: data[i][3],
          displayName: data[i][4]
        });
      }
    }
    return { success: true, data: users };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addUser(userData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');

    // Check duplicate
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userData.username) {
        return { success: false, message: 'Username already exists' };
      }
    }

    sheet.appendRow([
      userData.username,
      userData.password,
      userData.role || 'user',
      true,
      userData.displayName || userData.username,
      new Date()
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateUser(rowNum, userData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    sheet.getRange(rowNum, 1, 1, 6).setValues([[
      userData.username,
      userData.password,
      userData.role,
      userData.active,
      userData.displayName,
      new Date()
    ]]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteUser(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function generateRandomPassword(length) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789@#$';
  let pass = '';
  for (let i = 0; i < (length || 8); i++) {
    pass += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return pass;
}

// ==========================================
// QUESTIONS MANAGEMENT (Admin)
// ==========================================
function getQuestions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');
    const data = sheet.getDataRange().getValues();
    const questions = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        questions.push({
          row: i + 1,
          id: data[i][0],
          questionTH: data[i][1],
          questionEN: data[i][2],
          opt1TH: data[i][3], opt1EN: data[i][4],
          opt2TH: data[i][5], opt2EN: data[i][6],
          opt3TH: data[i][7], opt3EN: data[i][8],
          opt4TH: data[i][9], opt4EN: data[i][10],
          answer: data[i][11] // 1,2,3,4
        });
      }
    }
    return { success: true, data: questions };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addQuestion(qData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');
    const lastRow = sheet.getLastRow();
    const newId = lastRow; // auto id
    sheet.appendRow([
      newId,
      qData.questionTH, qData.questionEN,
      qData.opt1TH, qData.opt1EN,
      qData.opt2TH, qData.opt2EN,
      qData.opt3TH, qData.opt3EN,
      qData.opt4TH, qData.opt4EN,
      qData.answer
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateQuestion(rowNum, qData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');
    sheet.getRange(rowNum, 1, 1, 12).setValues([[
      qData.id,
      qData.questionTH, qData.questionEN,
      qData.opt1TH, qData.opt1EN,
      qData.opt2TH, qData.opt2EN,
      qData.opt3TH, qData.opt3EN,
      qData.opt4TH, qData.opt4EN,
      qData.answer
    ]]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteQuestion(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// SETTINGS (Admin)
// ==========================================
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }
    return { success: true, data: settings };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveSettings(key, value) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return { success: true };
      }
    }
    sheet.appendRow([key, value]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// VIDEOS (Admin)
// ==========================================
function getVideos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Videos');
    const data = sheet.getDataRange().getValues();
    const videos = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        videos.push({
          row: i + 1,
          title: data[i][0],
          url: data[i][1],
          active: data[i][2]
        });
      }
    }
    return { success: true, data: videos };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addVideo(videoData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Videos');
    sheet.appendRow([videoData.title, videoData.url, true, new Date()]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateVideo(rowNum, videoData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Videos');
    sheet.getRange(rowNum, 1, 1, 3).setValues([[videoData.title, videoData.url, videoData.active]]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteVideo(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Videos');
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// SAVE TRAINING RESULT
// ==========================================
function saveTrainingResult(traineeData, answers, score, lang) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Results');
    const now = new Date();

    // Build answer row: Ans1..Ans30
    const answerValues = [];
    for (let i = 1; i <= 30; i++) {
      const a = answers[i] || '';
      answerValues.push(a);
    }

    const row = [
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      traineeData.name,
      traineeData.age,
      traineeData.blood,
      traineeData.noAddress,
      traineeData.moo,
      traineeData.road,
      traineeData.subdistrict,
      traineeData.district,
      traineeData.city,
      traineeData.postCode,
      traineeData.tel,
      traineeData.company,
      traineeData.email,
      score,
      lang,
      ...answerValues
    ];
    sheet.appendRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// PDF GENERATION
// ==========================================
function generateAnswerSheetPDF(traineeData, answers, score, lang) {
  try {
    const now = new Date();
    const months_th = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                       'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const months_en = ['January','February','March','April','May','June',
                       'July','August','September','October','November','December'];

    const day = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd');
    const monthIndex = now.getMonth();
    const year_th = (now.getFullYear() + 543).toString();
    const year_en = now.getFullYear().toString();
    const month = lang === 'th' ? months_th[monthIndex] : months_en[monthIndex];
    const year = lang === 'th' ? year_th : year_en;

    // Labels for options
    const optLabels_th = { '1': 'ก', '2': 'ข', '3': 'ค', '4': 'ง' };
    const optLabels_en = { '1': 'A', '2': 'B', '3': 'C', '4': 'D' };
    const optLabels = lang === 'th' ? optLabels_th : optLabels_en;

    // Copy template
    const templateFile = DriveApp.getFileById(SLIDES_TEMPLATE_ID);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const copyName = 'AnswerSheet_' + traineeData.name + '_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const copy = templateFile.makeCopy(copyName, folder);
    const copyId = copy.getId();

    // Open copied slides
    const presentation = SlidesApp.openById(copyId);
    const slides = presentation.getSlides();

    // Build replacement map
    const replacements = {
      '{{Date}}': day,
      '{{Month}}': month,
      '{{Year}}': year,
      '{{Name}}': traineeData.name || '',
      '{{Age}}': traineeData.age || '',
      '{{Blood}}': traineeData.blood || '',
      '{{NoAddress}}': traineeData.noAddress || '',
      '{{Moo}}': traineeData.moo || '',
      '{{Road}}': traineeData.road || '',
      '{{Subdistrict}}': traineeData.subdistrict || '',
      '{{District}}': traineeData.district || '',
      '{{City}}': traineeData.city || '',
      '{{PostCode}}': traineeData.postCode || '',
      '{{Company}}': traineeData.company || '',
      '{{Tel}}': traineeData.tel || '',
      '{{TotalScore}}': score.toString() + '/30'
    };

    // Add answer placeholders
    for (let i = 1; i <= 30; i++) {
      const raw = answers[i] ? answers[i].toString() : '';
      replacements['{{Ans' + i + '}}'] = optLabels[raw] || '-';
    }

    // Replace in all slides
    slides.forEach(function(slide) {
      slide.getShapes().forEach(function(shape) {
        if (shape.getText) {
          const textRange = shape.getText();
          Object.keys(replacements).forEach(function(placeholder) {
            textRange.replaceAllText(placeholder, replacements[placeholder]);
          });
        }
      });
    });

    presentation.saveAndClose();

    // Export as PDF
    const pdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
    pdfBlob.setName(copyName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);

    // Delete the slides copy (keep only PDF)
    DriveApp.getFileById(copyId).setTrashed(true);

    // Return base64 for download
    const pdfBytes = pdfBlob.getBytes();
    const base64 = Utilities.base64Encode(pdfBytes);

    return {
      success: true,
      fileName: copyName + '.pdf',
      base64: base64,
      fileId: pdfFile.getId()
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// SETUP SHEETS (Run once to create structure)
// ==========================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Users sheet
  let users = ss.getSheetByName('Users');
  if (!users) {
    users = ss.insertSheet('Users');
    users.appendRow(['Username', 'Password', 'Role', 'Active', 'DisplayName', 'CreatedAt']);
    users.appendRow(['admin', 'Admin@1234', 'admin', true, 'Administrator', new Date()]);
    users.appendRow(['user01', 'User@1234', 'user', true, 'Test User', new Date()]);
  }

  // Questions sheet
  let questions = ss.getSheetByName('Questions');
  if (!questions) {
    questions = ss.insertSheet('Questions');
    questions.appendRow(['ID','QuestionTH','QuestionEN','Opt1TH','Opt1EN','Opt2TH','Opt2EN','Opt3TH','Opt3EN','Opt4TH','Opt4EN','Answer']);
  }

  // Results sheet
  let results = ss.getSheetByName('Results');
  if (!results) {
    results = ss.insertSheet('Results');
    const headers = ['Timestamp','Name','Age','Blood','NoAddress','Moo','Road','Subdistrict','District','City','PostCode','Tel','Company','Email','Score','Lang'];
    for (let i = 1; i <= 30; i++) headers.push('Ans' + i);
    results.appendRow(headers);
  }

  // Settings sheet
  let settings = ss.getSheetByName('Settings');
  if (!settings) {
    settings = ss.insertSheet('Settings');
    settings.appendRow(['Key', 'Value']);
    settings.appendRow(['passing_score', 27]);
    settings.appendRow(['total_questions', 30]);
  }

  // Videos sheet
  let videos = ss.getSheetByName('Videos');
  if (!videos) {
    videos = ss.insertSheet('Videos');
    videos.appendRow(['Title', 'YouTubeURL', 'Active', 'CreatedAt']);
  }

  return 'Setup complete!';
}
