/**
 * בית העלמין בית אל - Backend מבוסס Google Sheets
 * ------------------------------------------------
 * קובץ זה מודבק לתוך Google Apps Script (Extensions > Apps Script
 * מתוך הגיליון עצמו). הוא חושף שני נתיבים:
 *   GET  -> מחזיר את כל רשימת הקברים כ-JSON (פתוח לכולם, לצורך החיפוש הציבורי)
 *   POST -> פעולות מוגנות: הוספה/עריכה/מחיקה של קברים (לעורכים ומנהלים),
 *           וניהול רשימת המורשים עצמה (למנהלים בלבד)
 *
 * ההרשאה נבדקת כך: הצד הלקוח (האפליקציה) שולח idToken (JWT) שהתקבל
 * מ-Google Identity Services לאחר התחברות אמיתית. השרת כאן מאמת את
 * ה-token מול שרתי גוגל (לא סומך על שום דבר שהלקוח טוען), ואז בודק
 * את התפקיד של המייל המאומת מול לשונית ה-Permissions בגיליון.
 *
 * שתי רמות הרשאה:
 *   עורך  - יכול להוסיף / לערוך / למחוק קברים, וגם להוסיף/להסיר תמונות
 *   מנהל  - הכל שעורך יכול, ובנוסף יכול להוסיף/להסיר מורשים אחרים
 */

// ==== הגדרות ====
const SHEET_NAME = 'Graves';           // לשונית נתוני הקברים
const PERMISSIONS_SHEET_NAME = 'Permissions'; // לשונית רשימת המורשים
const GOOGLE_CLIENT_ID = '656595279986-48a5mirlpia1rb1r032h7vgbn6optek7.apps.googleusercontent.com';

// סדר העמודות בגיליון Graves לפי מיקום (A,B,C...) - קבוע, לא תלוי בטקסט
// שכתוב בשורת הכותרות. אפשר לכתוב בשורה הראשונה כל טקסט שרוצים (גם עברית).
// A=id, B=block, C=col, D=rowi, E=num, F=status ("מלא"/"שמור"/"פנוי"),
// G=firstName, H=lastName, I=fatherName, J=death, K=reservedFor, L=depth,
// M=description, N=photo, O=linkUrl, P=linkText, Q=photoX, R=photoY
// (photoX/photoY = מיקום הקבר על גבי תמונת האוויר של החלקה, אחוזים 0-100, מכויל בממשק הניהול)
//
// *** חשוב: יש להתאים בפועל את עמודות הגיליון לסדר הזה (כותרת חופשית בכל עמודה,
// למשל "שם פרטי", "שם משפחה", "שם האב", "שמור עבור מי" וכו') ***
const COLUMNS = ['id','block','col','rowi','num','status','firstName','lastName','fatherName','death','reservedFor','depth','description','photo','linkUrl','linkText','photoX','photoY'];

// עמודות לשונית Permissions: A=מייל, B=תפקיד ("עורך" או "מנהל")
const PERM_COLUMNS = ['email','role'];

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}
function getPermissionsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PERMISSIONS_SHEET_NAME);
}

function sheetToObjects() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // מדלג על שורת הכותרות, יהיה אשר יהיה כתוב בה
  return rows
    .filter(r => r[0] !== '' && r[0] !== null) // מדלג על שורות ריקות
    .map(r => {
      const obj = {};
      COLUMNS.forEach((key, i) => { obj[key] = r[i]; });
      // נירמול טיפוסים
      obj.id = Number(obj.id);
      obj.col = Number(obj.col);
      obj.rowi = Number(obj.rowi);
      obj.num = obj.num === '' ? null : Number(obj.num);
      obj.death = obj.death ? formatDate_(obj.death) : null;
      obj.status = obj.status === '' || obj.status === null || obj.status === undefined ? null : String(obj.status).trim();
      obj.firstName = obj.firstName === '' || obj.firstName === null || obj.firstName === undefined ? null : String(obj.firstName);
      obj.lastName = obj.lastName === '' || obj.lastName === null || obj.lastName === undefined ? null : String(obj.lastName);
      obj.fatherName = obj.fatherName === '' || obj.fatherName === null || obj.fatherName === undefined ? null : String(obj.fatherName);
      obj.reservedFor = obj.reservedFor === '' || obj.reservedFor === null || obj.reservedFor === undefined ? null : String(obj.reservedFor);
      obj.description = obj.description === '' ? null : String(obj.description);
      obj.depth = obj.depth === '' || obj.depth === null || obj.depth === undefined ? 1 : Number(obj.depth);
      obj.photo = obj.photo === '' ? null : String(obj.photo);
      obj.linkUrl = obj.linkUrl === '' || obj.linkUrl === null || obj.linkUrl === undefined ? null : String(obj.linkUrl);
      obj.linkText = obj.linkText === '' || obj.linkText === null || obj.linkText === undefined ? null : String(obj.linkText);
      obj.photoX = obj.photoX === '' || obj.photoX === null || obj.photoX === undefined ? null : Number(obj.photoX);
      obj.photoY = obj.photoY === '' || obj.photoY === null || obj.photoY === undefined ? null : Number(obj.photoY);
      return obj;
    });
}

function formatDate_(val) {
  if (val instanceof Date || (val && typeof val.getFullYear === 'function')) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// מסנן שדות רגישים לפני שהנתונים יוצאים מהשרת לבקשה ציבורית לא-מאומתת (GET).
// זה קורה בשרת, לא בדפדפן - כלומר קברים שאינם "מלא" פשוט לא כוללים בכלל שם,
// שם אב, תאריך פטירה, "שמור עבור מי", תיאור, תמונה או קישור בתשובת ה-JSON,
// ולא רק מוסתרים ע"י קוד ה-HTML/JS בצד הלקוח.
function sanitizeForPublic_(grave) {
  if (grave.status === 'מלא') return grave;
  return {
    id: grave.id, block: grave.block, col: grave.col, rowi: grave.rowi,
    num: grave.num, status: grave.status, depth: grave.depth,
    photoX: grave.photoX, photoY: grave.photoY
  };
}

// ==== GET - קריאת הנתונים (פתוח לכולם, ללא אימות - הנתונים כבר מסוננים בשרת) ====
function doGet(e) {
  try {
    const graves = sheetToObjects().map(sanitizeForPublic_);
    return jsonResponse_({ success: true, graves: graves });
  } catch (err) {
    return jsonResponse_({ success: false, message: String(err) });
  }
}

// ==== אימות Google ותפקידים ====
function verifyGoogleToken_(idToken) {
  if (!idToken) return null;
  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const data = JSON.parse(res.getContentText());
  if (data.aud !== GOOGLE_CLIENT_ID) return null;
  if (data.email_verified !== 'true' && data.email_verified !== true) return null;
  if (!data.email) return null;
  return String(data.email).trim().toLowerCase();
}

function getPermissionsList_() {
  const sheet = getPermissionsSheet();
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      email: String(r[0]).trim().toLowerCase(),
      role: String(r[1]).trim() === 'מנהל' ? 'admin' : 'editor'
    }));
}

function getRole_(email) {
  const list = getPermissionsList_();
  const found = list.find(u => u.email === email);
  return found ? found.role : null; // 'admin' | 'editor' | null
}

// ==== POST - כל הפעולות המוגנות ====
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const email = verifyGoogleToken_(body.idToken);

    if (!email) {
      return jsonResponse_({ success: false, message: 'אימות ההתחברות נכשל או פג תוקף. יש להתחבר שוב עם Google.' });
    }
    const role = getRole_(email);
    if (!role) {
      return jsonResponse_({ success: false, message: 'החשבון ' + email + ' אינו מורשה.' });
    }

    const action = body.action;

    // מי אני - כל משתמש מורשה (עורך או מנהל) יכול לבדוק את התפקיד שלו
    if (action === 'whoami') {
      return jsonResponse_({ success: true, email: email, role: role });
    }

    // רשימת קברים מלאה (לא מסוננת) - לעורך ולמנהל בלבד, לצורך ממשק הניהול.
    // בניגוד ל-doGet הציבורי, כאן מגיעים כל השדות כולל "שמור עבור מי" וכו',
    // כי המשתמש כבר עבר אימות Google ואומת מול לשונית ההרשאות.
    if (action === 'graves') {
      return jsonResponse_({ success: true, graves: sheetToObjects() });
    }

    // פעולות על קברים - עורך ומנהל כאחד
    if (action === 'add' || action === 'update' || action === 'delete') {
      const grave = body.grave || {};
      if (action === 'add') return handleAdd_(grave);
      if (action === 'update') return handleUpdate_(grave);
      if (action === 'delete') return handleDelete_(grave.id);
    }

    // תמונת קבר - עורך ומנהל כאחד (כל משתמש מורשה; role כבר אומת למעלה)
    if (action === 'uploadPhoto' || action === 'removePhoto') {
      if (action === 'uploadPhoto') return handleUploadPhoto_(body.id, body.imageBase64, body.mimeType);
      if (action === 'removePhoto') return handleRemovePhoto_(body.id);
    }

    // ניהול הרשאות - מנהל בלבד
    if (action === 'listUsers' || action === 'addUser' || action === 'removeUser') {
      if (role !== 'admin') {
        return jsonResponse_({ success: false, message: 'רק משתמש בתפקיד "מנהל" יכול לנהל הרשאות.' });
      }
      if (action === 'listUsers') return jsonResponse_({ success: true, users: getPermissionsList_() });
      if (action === 'addUser') return handleAddUser_(body.newEmail, body.newRole, email);
      if (action === 'removeUser') return handleRemoveUser_(body.targetEmail, email);
    }

    return jsonResponse_({ success: false, message: 'פעולה לא מוכרת.' });
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
      // שדה שלא נשלח מהלקוח בכלל (undefined) - שומרים את הערך הקיים בשורה,
      // כדי שעדכון חלקי (למשל שמירת טופס בלי שדה התמונה) לא ימחק שדות אחרים.
      const row = COLUMNS.map((col, idx) =>
        grave[col] !== undefined ? (grave[col] === null ? '' : grave[col]) : data[i][idx]
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

// ==== ניהול הרשאות (מנהל בלבד, נבדק כבר ב-doPost) ====
function handleAddUser_(newEmail, newRole, actorEmail) {
  newEmail = (newEmail || '').trim().toLowerCase();
  if (!newEmail || newEmail.indexOf('@') === -1) {
    return jsonResponse_({ success: false, message: 'כתובת מייל לא תקינה.' });
  }
  const roleHe = (newRole === 'admin') ? 'מנהל' : 'עורך';
  const sheet = getPermissionsSheet();
  const data = sheet.getDataRange().getValues();

  // אם המייל כבר קיים - נעדכן את התפקיד שלו במקום להוסיף כפילות
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === newEmail) {
      sheet.getRange(i + 1, 2).setValue(roleHe);
      return jsonResponse_({ success: true, message: 'התפקיד עודכן.' });
    }
  }
  sheet.appendRow([newEmail, roleHe]);
  return jsonResponse_({ success: true });
}

function handleRemoveUser_(targetEmail, actorEmail) {
  targetEmail = (targetEmail || '').trim().toLowerCase();
  if (targetEmail === actorEmail) {
    return jsonResponse_({ success: false, message: 'לא ניתן להסיר את ההרשאה של עצמך.' });
  }
  const sheet = getPermissionsSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetEmail) {
      sheet.deleteRow(i + 1);
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, message: 'המשתמש לא נמצא ברשימה.' });
}

// ==== תמונות קבר (עורך ומנהל כאחד, נבדק כבר ב-doPost) ====
const PHOTOS_FOLDER_NAME = 'Cemetery Grave Photos';

function getOrCreatePhotosFolder_() {
  const it = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTOS_FOLDER_NAME);
}

function setPhotoUrlOnRow_(id, url) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idColIndex = COLUMNS.indexOf('id');
  const photoColIndex = COLUMNS.indexOf('photo');
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idColIndex]) === Number(id)) {
      sheet.getRange(i + 1, photoColIndex + 1).setValue(url);
      return true;
    }
  }
  return false;
}

function handleUploadPhoto_(id, imageBase64, mimeType) {
  if (!id || !imageBase64) {
    return jsonResponse_({ success: false, message: 'חסר מזהה קבר או תמונה.' });
  }
  try {
    const folder = getOrCreatePhotosFolder_();
    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', 'grave_' + id + '_' + new Date().getTime() + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';

    const updated = setPhotoUrlOnRow_(id, url);
    if (!updated) {
      return jsonResponse_({ success: false, message: 'לא נמצא קבר עם המזהה הזה לעדכון.' });
    }
    return jsonResponse_({ success: true, url: url });
  } catch (err) {
    return jsonResponse_({ success: false, message: 'שגיאה בהעלאת התמונה: ' + String(err) });
  }
}

function handleRemovePhoto_(id) {
  if (!id) return jsonResponse_({ success: false, message: 'חסר מזהה קבר.' });
  const updated = setPhotoUrlOnRow_(id, '');
  if (!updated) return jsonResponse_({ success: false, message: 'לא נמצא קבר עם המזהה הזה.' });
  return jsonResponse_({ success: true });
}
