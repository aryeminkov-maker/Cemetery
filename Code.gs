/**
 * בית העלמין בית אל - Backend מבוסס Google Sheets
 * ------------------------------------------------
 * קובץ זה מודבק לתוך Google Apps Script (Extensions > Apps Script
 * מתוך הגיליון עצמו). הוא חושף שני נתיבים:
 *   GET  -> מחזיר את כל רשימת הקברים כ-JSON (פתוח לכולם, לצורך החיפוש הציבורי)
 *   POST -> מוסיף / מעדכן / מוחק רשומה (מוגן לפי רשימת מיילים מורשים)
 *
 * חשוב: זהו מנגנון הרשאה "רך" (בודק את המייל שנשלח בבקשה מול רשימה),
 * לא אימות Google אמיתי. מתאים לכלי ניהול פנימי בין אנשי צוות מוכרים,
 * לא לאבטחה ברמה גבוהה. אם בעתיד יידרש רמת אבטחה גבוהה יותר -
 * אפשר לשדרג לאימות Google Sign-In אמיתי.
 */

// ==== הגדרות ====
const SHEET_NAME = 'Graves';                    // שם הלשונית בגיליון
const ALLOWED_EDITORS = [                        // רשימת מיילים מורשים לעריכה
  'arye.minkov@gmail.com'
];

// סדר העמודות בגיליון (השורה הראשונה חייבת להיות בדיוק כך):
// id | block | col | rowi | num | status | name | death | description | depth
const COLUMNS = ['id','block','col','rowi','num','status','name','death','description','depth'];

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function sheetToObjects() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .filter(r => r[0] !== '' && r[0] !== null) // מדלג על שורות ריקות
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      // נירמול טיפוסים
      obj.id = Number(obj.id);
      obj.col = Number(obj.col);
      obj.rowi = Number(obj.rowi);
      obj.num = obj.num === '' ? null : Number(obj.num);
      obj.death = obj.death ? formatDate_(obj.death) : null;
      obj.name = obj.name === '' ? null : String(obj.name);
      obj.description = obj.description === '' ? null : String(obj.description);
      obj.depth = obj.depth === '' || obj.depth === null || obj.depth === undefined ? 1 : Number(obj.depth);
      return obj;
    });
}

function formatDate_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==== GET - קריאת הנתונים (פתוח לכולם) ====
function doGet(e) {
  try {
    const graves = sheetToObjects();
    return jsonResponse_({ success: true, graves: graves });
  } catch (err) {
    return jsonResponse_({ success: false, message: String(err) });
  }
}

// ==== POST - הוספה / עריכה / מחיקה (מוגן) ====
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const email = (body.email || '').trim().toLowerCase();

    if (ALLOWED_EDITORS.indexOf(email) === -1) {
      return jsonResponse_({ success: false, message: 'אין לך הרשאת עריכה עם המייל הזה.' });
    }

    const action = body.action;
    const grave = body.grave || {};

    if (action === 'add') {
      return handleAdd_(grave);
    } else if (action === 'update') {
      return handleUpdate_(grave);
    } else if (action === 'delete') {
      return handleDelete_(grave.id);
    } else {
      return jsonResponse_({ success: false, message: 'פעולה לא מוכרת.' });
    }
  } catch (err) {
    return jsonResponse_({ success: false, message: String(err) });
  }
}

function handleAdd_(grave) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const ids = data.slice(1).map(r => Number(r[0])).filter(n => !isNaN(n));
  const newId = ids.length ? Math.max.apply(null, ids) + 1 : 1;

  const row = COLUMNS.map(col => {
    if (col === 'id') return newId;
    return grave[col] !== undefined && grave[col] !== null ? grave[col] : '';
  });
  sheet.appendRow(row);
  return jsonResponse_({ success: true, id: newId });
}

function handleUpdate_(grave) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idColIndex = COLUMNS.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idColIndex]) === Number(grave.id)) {
      const row = COLUMNS.map(col =>
        grave[col] !== undefined && grave[col] !== null ? grave[col] : ''
      );
      sheet.getRange(i + 1, 1, 1, COLUMNS.length).setValues([row]);
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, message: 'לא נמצאה רשומה עם המזהה הזה.' });
}

function handleDelete_(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idColIndex = COLUMNS.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idColIndex]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, message: 'לא נמצאה רשומה עם המזהה הזה.' });
}
