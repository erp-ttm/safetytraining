// ============================================================
// TTM Training System - Code.gs (Google Apps Script Backend)
// ============================================================

const SPREADSHEET_ID = '1flL8UacixiOags5vufPuWK--z6waWtTXdM3WBeJ7r6c'; // ← ใส่ Spreadsheet ID ของคุณ
const SLIDE_TEMPLATE_ID = '1gO7YG67uSU1Koud5xhsVGNht8_hcIET0HAH_cURfUcA'; // ← ใส่ Google Slides Template ID

// Sheet names
const SHEET_USERS     = 'Users';
const SHEET_QUESTIONS = 'Questions';
const SHEET_RESULTS   = 'Results';
const SHEET_SETTINGS  = 'Settings';
const SHEET_VIDEOS    = 'Videos';

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  let template;
  if (page === 'admin') {
    template = HtmlService.createTemplateFromFile('Admin');
  } else if (page === 'user') {
    template = HtmlService.createTemplateFromFile('User');
  } else {
    template = HtmlService.createTemplateFromFile('Index');
  }
  return template.evaluate()
    .setTitle('ระบบอบรมข้อกำหนดด้านความปลอดภัย TTM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// GET SCRIPT URL — ใช้สำหรับ redirect หลัง login
// คืนค่า URL ของ Web App พร้อม ?page=xxx
// ============================================================
function getWebAppUrl(page) {
  const base = ScriptApp.getService().getUrl();
  if (!page || page === 'index') return base;
  return base + '?page=' + page;
}

// ============================================================
// AUTHENTICATION
// ============================================================
function login(username, password) {
  try {
    username = sanitize(username);
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][0]).trim() === username &&
        String(data[i][1]).trim() === String(password).trim() &&
        String(data[i][3]).trim() !== 'inactive'
      ) {
        return {
          success:     true,
          role:        data[i][2],       // 'admin' or 'user'
          username:    username,
          displayName: data[i][4] || username,
          redirectUrl: getWebAppUrl(data[i][2] === 'admin' ? 'admin' : 'user')
        };
      }
    }
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง | Invalid username or password' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function sanitize(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/[<>"'%;()&+]/g, '').trim().substring(0, 200);
}

// ============================================================
// QUESTIONS MANAGEMENT (Admin)
// ============================================================
function getQuestions() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_QUESTIONS);
  const data  = sheet.getDataRange().getValues();
  const questions = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      questions.push({
        id:         i,
        rowIndex:   i + 1,
        questionTH: data[i][0],
        questionEN: data[i][1],
        opt1TH: data[i][2],  opt1EN: data[i][3],
        opt2TH: data[i][4],  opt2EN: data[i][5],
        opt3TH: data[i][6],  opt3EN: data[i][7],
        opt4TH: data[i][8],  opt4EN: data[i][9],
        answer: data[i][10]  // 1,2,3,4
      });
    }
  }
  return questions;
}

function addQuestion(q) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_QUESTIONS);
    sheet.appendRow([
      sanitize(q.questionTH), sanitize(q.questionEN),
      sanitize(q.opt1TH), sanitize(q.opt1EN),
      sanitize(q.opt2TH), sanitize(q.opt2EN),
      sanitize(q.opt3TH), sanitize(q.opt3EN),
      sanitize(q.opt4TH), sanitize(q.opt4EN),
      parseInt(q.answer)
    ]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function updateQuestion(rowIndex, q) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_QUESTIONS);
    sheet.getRange(rowIndex, 1, 1, 11).setValues([[
      sanitize(q.questionTH), sanitize(q.questionEN),
      sanitize(q.opt1TH), sanitize(q.opt1EN),
      sanitize(q.opt2TH), sanitize(q.opt2EN),
      sanitize(q.opt3TH), sanitize(q.opt3EN),
      sanitize(q.opt4TH), sanitize(q.opt4EN),
      parseInt(q.answer)
    ]]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function deleteQuestion(rowIndex) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_QUESTIONS).deleteRow(rowIndex);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================
// SETTINGS (Passing Score)
// ============================================================
function getSettings() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const data  = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function saveSetting(key, value) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SETTINGS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return { success: true };
      }
    }
    sheet.appendRow([key, value]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================
// VIDEO MANAGEMENT (Admin)
// ============================================================
function getVideos() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_VIDEOS);
  const data  = sheet.getDataRange().getValues();
  const videos = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      videos.push({
        id: i, rowIndex: i + 1,
        title: data[i][0], url: data[i][1], order: data[i][2]
      });
    }
  }
  return videos;
}

function addVideo(title, url, order) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_VIDEOS)
      .appendRow([sanitize(title), sanitize(url), parseInt(order) || 1]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function updateVideo(rowIndex, title, url, order) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_VIDEOS)
      .getRange(rowIndex, 1, 1, 3).setValues([[sanitize(title), sanitize(url), parseInt(order) || 1]]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function deleteVideo(rowIndex) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_VIDEOS).deleteRow(rowIndex);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================
// USER MANAGEMENT (Admin)
// ============================================================
function getUsers() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data  = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      users.push({
        rowIndex:    i + 1,
        username:    data[i][0],
        password:    data[i][1],
        role:        data[i][2],
        status:      data[i][3],
        displayName: data[i][4]
      });
    }
  }
  return users;
}

function addUser(username, password, role, displayName) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS)
      .appendRow([sanitize(username), password, sanitize(role), 'active', sanitize(displayName)]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function updateUser(rowIndex, username, password, role, status, displayName) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS)
      .getRange(rowIndex, 1, 1, 5).setValues([[
        sanitize(username), password, sanitize(role), sanitize(status), sanitize(displayName)
      ]]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function deleteUser(rowIndex) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS).deleteRow(rowIndex);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================
// SUBMIT QUIZ & GENERATE PDF
// ============================================================
function submitQuiz(userData, answers, score, totalQuestions, language) {
  try {
    const settings     = getSettings();
    const passingScore = parseInt(settings['passing_score']) || 27;

    const now     = new Date();
    const day     = Utilities.formatDate(now, 'Asia/Bangkok', 'dd');
    const month   = Utilities.formatDate(now, 'Asia/Bangkok', 'MM');
    const year    = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy');
    const dateStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_RESULTS);

    const rowData = [
      dateStr, userData.username, userData.name, userData.age, userData.blood,
      userData.noAddress, userData.moo, userData.road, userData.subdistrict,
      userData.district, userData.city, userData.postCode, userData.tel,
      userData.company, score, totalQuestions,
      score >= passingScore ? 'PASS' : 'FAIL', language
    ];
    for (let i = 1; i <= 30; i++) {
      rowData.push(answers['ans' + i] || '-');
    }
    sheet.appendRow(rowData);

    if (score >= passingScore) {
      const pdfUrl = generateAnswerSheetPDF(userData, answers, score, day, month, year, language);
      return { success: true, passed: true, pdfUrl: pdfUrl, score: score, passingScore: passingScore };
    } else {
      return { success: true, passed: false, score: score, passingScore: passingScore };
    }
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// PDF GENERATION FROM GOOGLE SLIDES TEMPLATE
// ============================================================
function generateAnswerSheetPDF(userData, answers, score, day, month, year, language) {
  const optionLabels = language === 'en'
    ? { '1':'A', '2':'B', '3':'C', '4':'D' }
    : { '1':'ก', '2':'ข', '3':'ค', '4':'ง' };

  const templateFile = DriveApp.getFileById(SLIDE_TEMPLATE_ID);
  const folder       = DriveApp.getRootFolder();
  const copyName     = 'AnswerSheet_' + sanitize(userData.name) + '_' + day + month + year;
  const copy         = templateFile.makeCopy(copyName, folder);
  const copyId       = copy.getId();

  const presentation = SlidesApp.openById(copyId);
  const slides       = presentation.getSlides();

  const replacements = {
    '{{Date}}':        day,
    '{{Month}}':       month,
    '{{Year}}':        year,
    '{{Name}}':        userData.name        || '',
    '{{Age}}':         String(userData.age  || ''),
    '{{Blood}}':       userData.blood       || '',
    '{{NoAddress}}':   userData.noAddress   || '',
    '{{Moo}}':         userData.moo         || '',
    '{{Road}}':        userData.road        || '',
    '{{Subdistrict}}': userData.subdistrict || '',
    '{{District}}':    userData.district    || '',
    '{{City}}':        userData.city        || '',
    '{{PostCode}}':    userData.postCode    || '',
    '{{Company}}':     userData.company     || '',
    '{{Tel}}':         userData.tel         || '',
    '{{TotalScore}}':  score + '/30'
  };

  for (let i = 1; i <= 30; i++) {
    const raw = answers['ans' + i] || '';
    replacements['{{Ans' + i + '}}'] = optionLabels[raw] || raw;
  }

  slides.forEach(function(slide) {
    slide.getShapes().forEach(function(shape) {
      try {
        const tf  = shape.getText();
        const txt = tf.asString();
        let newTxt = txt;
        Object.keys(replacements).forEach(function(k) {
          newTxt = newTxt.split(k).join(replacements[k]);
        });
        if (newTxt !== txt) tf.setText(newTxt);
      } catch(e) {}
    });
    slide.getTables().forEach(function(table) {
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getNumColumns(); c++) {
          try {
            const cell = table.getCell(r, c);
            const txt  = cell.getText().asString();
            let newTxt = txt;
            Object.keys(replacements).forEach(function(k) {
              newTxt = newTxt.split(k).join(replacements[k]);
            });
            if (newTxt !== txt) cell.getText().setText(newTxt);
          } catch(e) {}
        }
      }
    });
  });

  presentation.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
  pdfBlob.setName(copyName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(copyId).setTrashed(true);

  return pdfFile.getDownloadUrl();
}

// ============================================================
// GET RESULTS (Admin)
// ============================================================
function getResults(limit) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RESULTS);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];
  const max = limit ? Math.min(data.length, limit + 1) : data.length;
  for (let i = 1; i < max; i++) {
    const row = {};
    headers.forEach(function(h, j) { row[h] = data[i][j]; });
    results.push(row);
  }
  return results.reverse();
}
