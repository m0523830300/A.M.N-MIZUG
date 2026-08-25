const SHEET_NAME = "קריאות שירות";
const TECH_SHEET_NAME = "טכנאים";
const PRICE_SHEET_NAME = "מחירון תיקונים";
const MEDIA_FOLDER_NAME = "קריאות שירות - קבצים מצורפים";
const OWNER_PHOTO_FOLDER_NAME = "קריאות שירות - תמונות הכוונה מהבעלים";
const CASH_SHEET_NAME = "כספים במזומן/אשראי";
const CUSTOMERS_SHEET_NAME = "לקוחות ופרוייקטים";
const WORKSPEC_SHEET_NAME = "מפרט עבודות";
const COLLECTION_REPORT_SHEET_NAME = "דוח גביה";
const COLLECTION_LOG_SHEET_NAME = "מעקב גבייה";

// פרטי חשבון SMS4Free לשליחת התרעות SMS לטכנאים כשמשייכים אליהם קריאה
const SMS4FREE_KEY = "KF0aqQ132";
const SMS4FREE_USER = "0527080177";
const SMS4FREE_PASS = "19621315";

// מנקה תווים בלתי נראים (סימוני כיווניות RTL/LTR שגיליונות Google לפעמים
// מכניסים אוטומטית לטקסט עברי) וגם רווחים מיותרים, כדי שהשוואת שמות
// טכנאים תמיד תעבוד נכון גם אם הטקסט נראה זהה לעין.
function normalizeName(str) {
  return String(str || "")
    .replace(/[\u200B\u200C\u200D\u200E\u200F\u202A-\u202E\uFEFF]/g, "")
    .trim();
}

// משווה מספרי טלפון תוך התעלמות ממקפים, רווחים, וקידומת 972+ -
// כדי שלא משנה באיזה פורמט הטלפון הוזן (עם/בלי מקפים), ההשוואה תעבוד.
function normalizePhone(str) {
  let digits = String(str || "").replace(/\D/g, "");
  if (digits.startsWith("972")) digits = "0" + digits.slice(3);
  return digits;
}

// שולח SMS דרך שירות SMS4Free. מחזיר {success, error} - לא זורק שגיאה כלפי חוץ
// כדי שכישלון בשליחת SMS לא יפיל פעולות אחרות (יצירת/עדכון קריאה).
function sendSMS(phone, message) {
  const recipient = normalizePhone(phone);
  if (!recipient) return { success:false, error:"אין מספר טלפון לטכנאי" };
  try {
    const res = UrlFetchApp.fetch("https://api.sms4free.co.il/ApiSMS/v2/SendSMS", {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({
        key: SMS4FREE_KEY,
        user: SMS4FREE_USER,
        pass: SMS4FREE_PASS,
        sender: SMS4FREE_USER,
        recipient: recipient,
        msg: message
      })
    });
    const responseText = res.getContentText();
    let parsed = null;
    try { parsed = JSON.parse(responseText); } catch (e) { /* לא JSON תקין */ }
    if (parsed && Number(parsed.status) > 0) {
      return { success:true };
    }
    return { success:false, error: (parsed && parsed.message) ? parsed.message : responseText };
  } catch (err) {
    return { success:false, error:String(err) };
  }
}

// שולף את מספר הטלפון של טכנאי לפי שם מתוך טאב "טכנאים"
function getTechnicianPhone(techName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TECH_SHEET_NAME);
  if (!sheet) return "";
  const name = normalizeName(techName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeName(data[i][0]) === name) return data[i][1];
  }
  return "";
}

// שולח ללקוח SMS שהטכנאי בדרך אליו, עם שם וטלפון הטכנאי ליצירת קשר.
// גם מעדכן את סטטוס הקריאה ל"בטיפול" כי זה מסמן בפועל שהטכנאי יצא לדרך.
function handleNotifyCustomerOnWay(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  let customerPhone = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) {
      targetRow = i + 1;
      customerPhone = data[i][4]; // E - טלפון
      break;
    }
  }
  if (targetRow === -1) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!customerPhone) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"אין מספר טלפון ללקוח בקריאה הזו"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const techPhone = getTechnicianPhone(body.technicianName);
  const message = `שלום, ${normalizeName(body.technicianName)} מטעם א.מ.ן מיזוג אוויר בדרך אליך.` +
    (techPhone ? `\nניתן ליצור קשר: ${techPhone}` : "");
  const smsResult = sendSMS(customerPhone, message);

  sheet.getRange(targetRow, 9).setValue("בטיפול"); // I - סטטוס
  sheet.getRange(targetRow, 12).setValue(new Date()); // L - עדכון אחרון

  return ContentService.createTextOutput(JSON.stringify(smsResult))
    .setMimeType(ContentService.MimeType.JSON);
}

// שולח לטכנאי הודעת SMS על שיוך/העברה של קריאה אליו
const TECHNICIAN_APP_URL = "https://m0523830300.github.io/A.M.N-MIZUG/technician.html";

function notifyTechnicianSMS(techName, customerName, address, issueType) {
  const phone = getTechnicianPhone(techName);
  if (!phone) return { success:false, error:"לא נמצא טלפון לטכנאי " + techName };
  const parts = ["קריאת שירות חדשה שויכה אליך - א.מ.ן מיזוג אוויר"];
  if (customerName) parts.push("לקוח: " + customerName);
  if (address) parts.push("כתובת: " + address);
  if (issueType) parts.push("תקלה: " + issueType);
  parts.push(TECHNICIAN_APP_URL);
  return sendSMS(phone, parts.join("\n"));
}

// פונקציית בדיקה ידנית בלבד - מריצים אותה פעם אחת מתוך עורך הסקריפט (Run)
// כדי להפעיל את בקשת ההרשאה לגישה לאינטרנט. אפשר למחוק אחר כך, לא בשימוש
// על ידי האפליקציה עצמה.
function testUrlFetchAuth() {
  const res = UrlFetchApp.fetch("https://api.sms4free.co.il/ApiSMS/v2/SendSMS", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({ key: SMS4FREE_KEY, user: SMS4FREE_USER, pass: SMS4FREE_PASS, sender: SMS4FREE_USER, recipient: SMS4FREE_USER, msg: "בדיקה" })
  });
  Logger.log(res.getContentText());
}

function doGet(e) {
  const page = e.parameter.page;
  const action = e.parameter.action;

  if (page === 'owner') {
    return HtmlService.createHtmlOutputFromFile('owner')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'technician') {
    return HtmlService.createHtmlOutputFromFile('technician')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === 'technicians') {
    return ContentService.createTextOutput(JSON.stringify(getTechnicianNames()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'priceList') {
    return ContentService.createTextOutput(JSON.stringify(getPriceList()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'collections') {
    if (!isValidCollectionToken_(e.parameter.ownerToken)) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"נדרשת כניסת בעלים"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(getCollectionCustomers()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = findLastContentRow(sheet, 1); // A - מזהה
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = headerRow.map(h => normalizeName(h));
  let rows = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
    rows = data
      .filter(row => row[0] !== "" && row[0] !== null) // מסננים שורות ריקות ("רפאים")
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
  }
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPriceList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRICE_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0])
    .map(r => ({
      name: normalizeName(r[0]),
      customerPrice: r[1] === "" ? null : Number(r[1]),
      techPrice: r[2] === "" ? null : Number(r[2])
    }));
}

function getTechnicianNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TECH_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0])
    .map(r => normalizeName(r[0]));
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "collectionLogin") {
    return handleCollectionLogin(body);
  }

  if (body.action === "login") {
    return handleLogin(body);
  }
  if (body.action === "addTechnician") {
    return handleAddTechnician(body);
  }
  if (body.action === "uploadFile") {
    return handleUploadFile(body);
  }
  if (body.action === "uploadOwnerPhoto") {
    return handleUploadOwnerPhoto(body);
  }
  if (body.action === "recordPayment") {
    return handleRecordPayment(body);
  }
  if (body.action === "notifyCustomerOnWay") {
    return handleNotifyCustomerOnWay(body);
  }
  if (body.action === "updateCollectionContact") {
    if (!isValidCollectionToken_(body.ownerToken)) return jsonOutput_({success:false, error:"הכניסה פגה"});
    return handleUpdateCollectionContact(body);
  }
  if (body.action === "logCollectionReminder") {
    if (!isValidCollectionToken_(body.ownerToken)) return jsonOutput_({success:false, error:"הכניסה פגה"});
    return handleLogCollectionReminder(body);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (body.action === "add") {
    // עבור "התקנה" - לא יוצרים קריאת שירות בכלל (לא רלוונטי לטכנאים),
    // רק מקימים/מאתרים לקוח ומקצים מספר מסמך. הלקוח מטופל ב-createCustomerDoc בנפרד.
    if (normalizeName(body.workType) === "התקנה") {
      return ContentService.createTextOutput(JSON.stringify({success:true, skippedCall:true}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const id = Utilities.getUuid();
    sheet.appendRow([
      id,                              // A - מזהה
      false,                           // B - עודכן במערכת
      new Date(),                      // C - תאריך פתיחה
      body.customerName,               // D - שם לקוח
      body.phone,                      // E - טלפון
      body.address,                    // F - כתובת
      body.issueType,                  // G - סוג תקלה
      body.notes || "",                // H - הערות פתיחה
      "חדש",                           // I - סטטוס
      "",                               // J - מה בוצע
      "",                               // K - נדרש המשך טיפול
      new Date(),                      // L - עדכון אחרון
      body.assignedTechnician ? normalizeName(body.assignedTechnician) : "",   // M - טכנאי משויך
      "",                               // N - מחיר טכנאי
      "",                               // O - כולל מעמ טכנאי
      "",                               // P - מחיר ללקוח
      "",                               // Q - נאמר ללקוח
      "",                               // R - פירוט חיוב
      "",                               // S - נתוני חיוב JSON
      "",                               // T - קבצים מצורפים (מהטכנאי)
      "",                               // U - שולם בפועל
      "",                               // V - סכום שהתקבל
      ""                                // W - תמונת הכוונה מהבעלים
    ]);
    // כופים פורמט טקסט על תא הטלפון - מונע מ-Google Sheets לפרש מספרים
    // שמתחילים ב-"+" (כמו 972+) כניסיון לנוסחה, מה שגורם ל-#ERROR!
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 5).setNumberFormat("@").setValue(String(body.phone || ""));

    // הערה: ה-SMS לטכנאי לא נשלח כאן יותר - הוא נשלח בקריאה נפרדת ברקע
    // מהאפליקציה (action: sendAssignmentSms), כדי שהמשתמש לא יחכה לתשובת
    // שרת ה-SMS (יכול לקחת כמה שניות) לפני שהקריאה בכלל מוצגת לו כנשלחה.
    return ContentService.createTextOutput(JSON.stringify({success:true, id}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "sendAssignmentSms") {
    let smsResult = null;
    try {
      smsResult = notifyTechnicianSMS(body.technicianName, body.customerName, body.address, body.issueType);
    } catch (smsErr) {
      smsResult = { success:false, error:String(smsErr) };
    }
    return ContentService.createTextOutput(JSON.stringify(smsResult))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "createCustomerDoc") {
    let docResult = { success:false };
    try {
      docResult = findOrCreateCustomerAndDocument(
        body.customerName, body.phone, body.address, body.workType
      );
    } catch (docErr) {
      docResult = { success:false, error:String(docErr) };
    }
    return ContentService.createTextOutput(JSON.stringify(docResult))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "update") {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        const row = i + 1;
        let needsSms = null;
        if (body.status) sheet.getRange(row, 9).setValue(body.status);
        if (body.workDone !== undefined) sheet.getRange(row, 10).setValue(body.workDone);
        if (body.followUpNeeded !== undefined) sheet.getRange(row, 11).setValue(body.followUpNeeded);
        if (body.assignedTechnician !== undefined) {
          const prevTech = normalizeName(data[i][12]); // M - טכנאי משויך (לפני העדכון)
          const newTech = normalizeName(body.assignedTechnician);
          sheet.getRange(row, 13).setValue(newTech);
          if (newTech && newTech !== prevTech) {
            // לא שולחים SMS כאן (זה חוסם) - רק מסמנים ללקוח שצריך לשלוח ברקע
            needsSms = { technicianName: newTech, customerName: data[i][3], address: data[i][5], issueType: data[i][6] };
          }
        }
        if (body.techPrice !== undefined) sheet.getRange(row, 14).setValue(body.techPrice);
        if (body.techVat !== undefined) sheet.getRange(row, 15).setValue(body.techVat);
        if (body.customerPrice !== undefined) sheet.getRange(row, 16).setValue(body.customerPrice);
        if (body.toldCustomer !== undefined) sheet.getRange(row, 17).setValue(body.toldCustomer);
        if (body.billingBreakdown !== undefined) sheet.getRange(row, 18).setValue(body.billingBreakdown);
        if (body.billingItemsJson !== undefined) sheet.getRange(row, 19).setValue(body.billingItemsJson);
        sheet.getRange(row, 12).setValue(new Date());
        return ContentService.createTextOutput(JSON.stringify({success:true, needsSms}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "delete") {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({success:true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "editCallDetails") {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        const row = i + 1;
        if (body.customerName !== undefined) sheet.getRange(row, 4).setValue(body.customerName);
        if (body.phone !== undefined) {
          sheet.getRange(row, 5).setNumberFormat("@").setValue(String(body.phone || ""));
        }
        if (body.address !== undefined) sheet.getRange(row, 6).setValue(body.address);
        if (body.issueType !== undefined) sheet.getRange(row, 7).setValue(body.issueType);
        sheet.getRange(row, 12).setValue(new Date());
        return ContentService.createTextOutput(JSON.stringify({success:true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===================== מעקב גבייה =====================
function handleCollectionLogin(body) {
  const configuredPin = String(PropertiesService.getScriptProperties().getProperty("COLLECTION_OWNER_PIN") || "");
  if (!configuredPin) return jsonOutput_({success:false, error:"טרם הוגדר קוד בעלים ב-Script Properties"});
  if (String(body.pin || "") !== configuredPin) return jsonOutput_({success:false, error:"קוד בעלים שגוי"});
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("collection_token_" + token, "1", 21600);
  return jsonOutput_({success:true, token});
}

function isValidCollectionToken_(token) {
  if (!token) return false;
  return CacheService.getScriptCache().get("collection_token_" + token) === "1";
}

function getCollectionCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName(COLLECTION_REPORT_SHEET_NAME);
  if (!reportSheet) return { success:false, error:"לא נמצא גיליון דוח גביה", customers:[] };

  const lastRow = findLastContentRow(reportSheet, 2); // B - שם לקוח
  if (lastRow < 2) return { success:true, customers:[], summary:{count:0,totalFinal:0,needsPricing:0} };

  const values = reportSheet.getRange(1, 1, lastRow, 18).getValues();
  const headers = values[0].map(normalizeName);
  const index = {};
  headers.forEach((h, i) => index[h] = i);

  const logInfo = getLastCollectionLogs_();
  const customers = values.slice(1).map(row => {
    const name = normalizeName(row[index["שם לקוח"]]);
    const phoneOptions = parseIsraeliPhones_(row[index["טלפון 1"]], row[index["טלפון 2"]]);
    const temporaryBalance = Number(row[index["יתרה זמנית"]]) || 0;
    const finalBalance = Number(row[index["יתרה סופית"]]) || 0;
    const unpricedCount = Number(row[index["עבודות לפני תמחור"]]) || 0;
    const amount = finalBalance > 0 ? finalBalance : temporaryBalance;
    const lastLog = logInfo[name] || null;
    return {
      name,
      phone1: phoneOptions[0] || "",
      phone2: phoneOptions[1] || "",
      phoneOptions,
      email: String(row[index["דואר אלקטרוני"]] || ""),
      address: String(row[index["כתובת"]] || ""),
      lastDocumentDate: formatDateForApi_(row[index["תאריך מסמך אחרון"]]),
      lastDocumentNumber: row[index["מסמך אחרון ללקוח"]] || "",
      temporaryBalance,
      finalBalance,
      amount,
      unpricedCount,
      amountIsFinal: unpricedCount === 0,
      lastReminderDate: lastLog ? lastLog.date : "",
      lastReminderChannel: lastLog ? lastLog.channel : "",
      lastReminderNote: lastLog ? lastLog.note : ""
    };
  }).filter(c => c.name && (c.temporaryBalance > 10 || c.finalBalance > 10));

  customers.sort((a, b) => {
    if (a.amountIsFinal !== b.amountIsFinal) return a.amountIsFinal ? 1 : -1;
    return b.amount - a.amount;
  });

  return {
    success:true,
    customers,
    summary:{
      count: customers.length,
      totalFinal: customers.filter(c => c.amountIsFinal).reduce((sum, c) => sum + c.amount, 0),
      needsPricing: customers.filter(c => !c.amountIsFinal).length
    }
  };
}

function parseIsraeliPhones_() {
  const values = Array.prototype.slice.call(arguments);
  const found = [];
  values.forEach(value => {
    let digits = String(value || "").replace(/\D/g, "");
    digits = digits.replace(/972(?=5\d{8})/g, "0");
    const matches = digits.match(/05\d{8}/g) || [];
    matches.forEach(phone => {
      if (found.indexOf(phone) === -1) found.push(phone);
    });
  });
  return found;
}

function handleUpdateCollectionContact(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!sheet) return jsonOutput_({success:false, error:"לא נמצא גיליון לקוחות ופרוייקטים"});

  const customerName = normalizeName(body.customerName);
  if (!customerName) return jsonOutput_({success:false, error:"חסר שם לקוח"});
  const lastRow = findLastContentRow(sheet, 2);
  const names = sheet.getRange(2, 2, Math.max(lastRow - 1, 1), 1).getValues();
  let targetRow = 0;
  for (let i = 0; i < names.length; i++) {
    if (normalizeName(names[i][0]) === customerName) { targetRow = i + 2; break; }
  }
  if (!targetRow) return jsonOutput_({success:false, error:"הלקוח לא נמצא"});

  if (body.phone1 !== undefined) sheet.getRange(targetRow, 3).setNumberFormat("@").setValue(String(body.phone1 || ""));
  if (body.email !== undefined) sheet.getRange(targetRow, 6).setValue(String(body.email || "").trim());
  return jsonOutput_({success:true});
}

function handleLogCollectionReminder(body) {
  const name = normalizeName(body.customerName);
  if (!name) return jsonOutput_({success:false, error:"חסר שם לקוח"});
  const sheet = getOrCreateCollectionLogSheet_();
  sheet.appendRow([
    new Date(), name, normalizeName(body.channel), Number(body.amount) || 0,
    body.amountIsFinal === true ? "כן" : "לא", String(body.phone || ""),
    String(body.email || ""), String(body.note || "")
  ]);
  return jsonOutput_({success:true, date:formatDateForApi_(new Date())});
}

function getOrCreateCollectionLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COLLECTION_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COLLECTION_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 8).setValues([[
      "תאריך ושעה", "שם לקוח", "ערוץ תזכורת", "סכום שהוצג",
      "הסכום סופי", "טלפון", "דואר אלקטרוני", "הערה"
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLastCollectionLogs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(COLLECTION_LOG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const result = {};
  data.forEach(row => {
    const name = normalizeName(row[1]);
    if (!name) return;
    result[name] = {
      date: formatDateForApi_(row[0]),
      channel: normalizeName(row[2]),
      note: String(row[7] || "")
    };
  });
  return result;
}

function formatDateForApi_(value) {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function findLastContentRow(sheet, col) {
  const values = sheet.getRange(1, col, sheet.getMaxRows(), 1).getValues();
  let last = 1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) last = i + 1;
  }
  return last;
}

// מקצה מספר מסמך חדש בלי כפילויות. שומר "מונה" משלנו (PropertiesService) כדי
// לא להסתמך רק על הנוסחה של הגיליון (=MAX('מפרט עבודות'!P:P)+1) - הנוסחה הזו
// "לא יודעת" שכבר הקצינו מספר עד שהוא מוקלד בפועל בטאב "מפרט עבודות", ולכן
// לבד היא הייתה חוזרת על אותו מספר בקריאות מהירות עוקבות. לוקחים את הגבוה
// מבין המונה השמור לנו לבין הנוסחה הרשמית, כדי גם למנוע כפילות וגם להישאר
// מסונכרן אם מישהו הוסיף מסמכים ידנית ישירות בגיליון.
function getNextDocNumber(custSheet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = Number(props.getProperty('nextDocNumber')) || 0;
    const liveNext = Number(custSheet.getRange("U1").getValue()) || 0;
    // רק מחזירים את המספר המיועד - לא שומרים עדיין! השמירה בפועל קורית
    // רק ב-confirmDocNumberUsed(), אחרי שכל הכתיבה לשורה הצליחה. כך אם
    // התהליך נכשל באמצע, אותו מספר יינתן שוב בניסיון הבא ולא "יבוזבז".
    return Math.max(stored, liveNext);
  } finally {
    lock.releaseLock();
  }
}

// נקרא רק אחרי שכל הכתיבה של שורת המסמך הסתיימה בהצלחה - "נועל" את המספר
// כדי שהניסיון הבא יתחיל מהמספר הבא, לא יחזור על אותו אחד.
function confirmDocNumberUsed(usedNumber) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = Number(props.getProperty('nextDocNumber')) || 0;
    if (usedNumber >= stored) {
      props.setProperty('nextDocNumber', String(usedNumber + 1));
    }
  } finally {
    lock.releaseLock();
  }
}

function findOrCreateCustomerAndDocument(customerName, phone, address, workType) {
  const custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!custSheet) {
    return { success:false, error:`לא נמצא טאב בשם "${CUSTOMERS_SHEET_NAME}"` };
  }

  const name = normalizeName(customerName);
  const pn = normalizePhone(phone);
  const addr = normalizeName(address);
  const workTypeVal = normalizeName(workType) || "תיקון";

  // --- שלב 1: איתור/יצירת לקוח (עמודות A-G) ---
  const lastCustRow = findLastContentRow(custSheet, 2); // B - שם לקוח
  // קוראים רק את עמודות הטלפון (C:D) לבדיקת התאמה - הרבה יותר מהיר מקריאת A:G כולה
  const phoneData = custSheet.getRange(2, 3, Math.max(lastCustRow - 1, 0), 2).getValues(); // C:D
  const nameData = custSheet.getRange(2, 2, Math.max(lastCustRow - 1, 0), 1).getValues(); // B
  let customerFound = false;
  let matchedName = name;
  for (let i = 0; i < phoneData.length; i++) {
    const rowPhone1 = normalizePhone(phoneData[i][0]); // C
    const rowPhone2 = normalizePhone(phoneData[i][1]); // D
    if (pn && (rowPhone1 === pn || rowPhone2 === pn)) {
      customerFound = true;
      matchedName = normalizeName(nameData[i][0]) || name;
      break;
    }
  }

  if (!customerFound) {
    const newCustRow = lastCustRow + 1;
    custSheet.getRange(newCustRow, 1, 1, 5).setValues([[new Date(), name, "", "", addr]]); // A,B,C,D,E
    // עמודת הטלפון נכתבת בנפרד כי היא דורשת פורמט טקסט מיוחד (מונע #ERROR! עם 972+)
    custSheet.getRange(newCustRow, 3).setNumberFormat("@").setValue(String(phone || ""));
  }

  // --- שלב 2: הקצאת מספר מסמך חדש ---
  const nextDocNumber = getNextDocNumber(custSheet);
  const lastProjRow = findLastContentRow(custSheet, 21); // U - מספר מסמך
  const targetRow = Math.max(lastCustRow, lastProjRow) + 1;

  // כתיבה תא-בתא (לא setValues מרוכז) - זו הגרסה שהוכחה כעובדת בפועל
  custSheet.getRange(targetRow, 21).setValue(nextDocNumber);  // U - מספר מסמך
  custSheet.getRange(targetRow, 22).setValue(matchedName);    // V - שם לקוח
  custSheet.getRange(targetRow, 23).setValue(addr);           // W - פרוייקט
  custSheet.getRange(targetRow, 24).setValue(addr);           // X - כתובת לפרוייקט
  custSheet.getRange(targetRow, 25).setValue(workTypeVal);    // Y - סוג עבודה
  custSheet.getRange(targetRow, 26).setValue(new Date());     // Z - תאריך שליחת הצעת מחיר

  // עמודה AK (37) - עמודה פנויה בגיליון. נוסחה חיה שמראה "חדש - טרם הוזן"
  // כל עוד מספר המסמך הזה לא מופיע בטאב "מפרט עבודות" (עמודה P) - ונעלמת
  // אוטומטית ברגע שהוא כן מוזן שם.
  custSheet.getRange(targetRow, 19).setFormula(
    `=IF(COUNTIF('${WORKSPEC_SHEET_NAME}'!P:P,U${targetRow})=0,"🆕 חדש - טרם הוזן במפרט עבודות","")`
  );

  confirmDocNumberUsed(nextDocNumber);

  return { success:true, docNumber: nextDocNumber, customerFound, targetRow };
}

function getOrCreateMediaFolder() {
  const folders = DriveApp.getFoldersByName(MEDIA_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(MEDIA_FOLDER_NAME);
}

function getOrCreateOwnerPhotoFolder() {
  const folders = DriveApp.getFoldersByName(OWNER_PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(OWNER_PHOTO_FOLDER_NAME);
}

// תמונת הכוונה שהבעלים מעלה כדי שהטכנאי יראה למה לצפות (למשל דגם מזגן/מעבה
// ספציפי) - נשמרת בתיקיית דרייב נפרדת מקבצי הטכנאי, ומוצגת רק בתוך האפליקציה
// (לא מוצגת בדוחות/בקבצים של הבעלים עצמו).
function handleUploadOwnerPhoto(body) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const folder = getOrCreateOwnerPhotoFolder();
    const bytes = Utilities.base64Decode(body.base64Data);
    const blob = Utilities.newBlob(bytes, body.mimeType, body.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();

    const richBuilder = SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(0, url.length, url);
    sheet.getRange(targetRow, 23).setRichTextValue(richBuilder.build()); // W - תמונת הכוונה מהבעלים
    sheet.getRange(targetRow, 12).setValue(new Date());

    return ContentService.createTextOutput(JSON.stringify({success:true, url}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUploadFile(body) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const folder = getOrCreateMediaFolder();
    const bytes = Utilities.base64Decode(body.base64Data);
    const blob = Utilities.newBlob(bytes, body.mimeType, body.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();

    const existingText = sheet.getRange(targetRow, 20).getValue();
    const existingUrls = existingText ? String(existingText).split("\n").filter(Boolean) : [];
    existingUrls.push(url);
    const fullText = existingUrls.join("\n");

    // בונים טקסט עשיר עם קישור אמיתי לכל URL - כך שהתא יהיה לחיץ בפועל
    // גם כשמסתכלים ישירות בגיליון (לא רק דרך האפליקציה)
    const richBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);
    let pos = 0;
    existingUrls.forEach((u) => {
      richBuilder.setLinkUrl(pos, pos + u.length, u);
      pos += u.length + 1; // +1 עבור ה-\n
    });
    sheet.getRange(targetRow, 20).setRichTextValue(richBuilder.build());
    sheet.getRange(targetRow, 12).setValue(new Date());

    return ContentService.createTextOutput(JSON.stringify({success:true, url}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleRecordPayment(body) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let customerName = "";
    let techPrice = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        targetRow = i + 1;
        customerName = data[i][3];             // D - שם לקוח
        techPrice = Number(data[i][13]) || 0;  // N - מחיר טכנאי
        break;
      }
    }
    if (targetRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"קריאה לא נמצאה"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const amountReceived = Number(body.amountReceived) || 0;
    const diff = amountReceived - techPrice;
    let diffNote;
    if (diff > 0) {
      diffNote = `טכנאי קיבל ${amountReceived} ש"ח, מגיע לו ${techPrice} ש"ח - יתרה לזכות העסק: ${diff} ש"ח`;
    } else if (diff < 0) {
      diffNote = `טכנאי קיבל ${amountReceived} ש"ח, מגיע לו ${techPrice} ש"ח - חוסר של ${-diff} ש"ח שיש להשלים לטכנאי`;
    } else {
      diffNote = `טכנאי קיבל ${amountReceived} ש"ח - תואם בדיוק את המחיר שנקבע לו (${techPrice} ש"ח)`;
    }

    // סימון בשורת הקריאה עצמה ששולם בפועל, וכמה
    sheet.getRange(targetRow, 21).setValue("כן");          // U - שולם בפועל
    sheet.getRange(targetRow, 22).setValue(amountReceived); // V - סכום שהתקבל
    sheet.getRange(targetRow, 12).setValue(new Date());     // L - עדכון אחרון

    // רישום שורה בטאב "כספים במזומן/אשראי" הקיים - התאמה לפי שמות כותרות
    // כדי לא לפגוע במבנה הקיים גם אם יש בו עוד עמודות שלא נוגעים בהן.
    const cashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_SHEET_NAME);
    if (!cashSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success:false,
        error:`לא נמצא טאב בשם "${CASH_SHEET_NAME}". טאבים קיימים בגיליון: ${SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s=>s.getName()).join(", ")}`
      })).setMimeType(ContentService.MimeType.JSON);
    }
    {
      const headers = cashSheet.getRange(1, 1, 1, cashSheet.getLastColumn()).getValues()[0]
        .map(h => normalizeName(h));
      const row = new Array(headers.length).fill("");
      const setByHeader = (headerName, value) => {
        const idx = headers.indexOf(headerName);
        if (idx !== -1) row[idx] = value;
      };
      setByHeader("חותמת זמן", new Date());
      setByHeader("עבור מה ההוצאה/הכנסה?", "קבלת תשלום מלקוח");
      setByHeader("סה\"כ בש\"ח", amountReceived);
      setByHeader("הערות:", diffNote);
      setByHeader("שם הלקוח", customerName);
      setByHeader("למי שולם", "ישירות לטכנאי");
      setByHeader("אמצעי תשלום", body.paymentMethod || "מזומן");
      setByHeader("טיפול בתשלום ישירות לטכנאי", "מקדמה לעובד");
      setByHeader("הערה לגבי טיפול בתשלום", diffNote);
      setByHeader("מקדמה לעובד", body.technicianName || "");
      // לא משתמשים ב-appendRow הרגיל כי בגיליון הזה יש עיצוב/רשימות נפתחות
      // מוגדרות מראש על אלפי שורות קדימה - זה גורם ל-appendRow "לחשוב" שהגיליון
      // מלא ולהוסיף שורה חדשה הרחק מתחת לתוכן הנראה. במקום זה, מוצאים בעצמנו
      // את השורה האחרונה שיש בה תוכן אמיתי בעמודה A (חותמת זמן) ומוסיפים מיד אחריה.
      const colA = cashSheet.getRange("A:A").getValues();
      let lastContentRow = 1;
      for (let i = 0; i < colA.length; i++) {
        if (colA[i][0] !== "") lastContentRow = i + 1;
      }
      cashSheet.getRange(lastContentRow + 1, 1, 1, row.length).setValues([row]);
    }

    return ContentService.createTextOutput(JSON.stringify({success:true, diff, diffNote}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleLogin(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TECH_SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"לא נמצא טאב טכנאים בגיליון"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const pin = normalizePhone(body.pin);
  if (!pin) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"יש להזין קוד"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowPin = normalizePhone(data[i][2]);
    if (rowPin && rowPin === pin) {
      return ContentService.createTextOutput(JSON.stringify({success:true, name: normalizeName(data[i][0])}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({success:false, error:"קוד שגוי"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAddTechnician(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TECH_SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"לא נמצא טאב טכנאים בגיליון"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const name = normalizeName(body.name);
  const phone = normalizeName(body.phone);
  const pin = normalizePhone(body.pin);
  if (!name) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"חובה למלא שם"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!pin || pin.length !== 4) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:"קוד הכניסה חייב להיות בדיוק 4 ספרות"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeName(data[i][0]) === name) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"טכנאי בשם הזה כבר קיים"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (normalizePhone(data[i][2]) === pin) {
      return ContentService.createTextOutput(JSON.stringify({success:false, error:"קוד הכניסה הזה כבר תפוס אצל טכנאי אחר"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  sheet.appendRow([name, phone, pin]);
  return ContentService.createTextOutput(JSON.stringify({success:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
