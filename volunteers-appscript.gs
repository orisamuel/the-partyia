// ============================================================
// רישום ערב מתנדבים — קוד להדבקה ב-Google Apps Script הקיים
// ============================================================
//
// איך מוסיפים:
// 1. פותחים את פרויקט ה-Apps Script הקיים (זה שמאחורי GOOGLE_SCRIPT_URL)
// 2. מדביקים את שתי הפונקציות למטה בסוף הקובץ
// 3. בתוך ה-doPost, איפה שכל שאר ה-actions מטופלים, מוסיפים:
//
//    if (action === 'registerVolunteerEvent') {
//      const result = registerVolunteerEvent(e.parameter.name, e.parameter.guests);
//      return ContentService.createTextOutput(JSON.stringify(result))
//        .setMimeType(ContentService.MimeType.JSON);
//    }
//
// 4. חשוב!!! פריסה מחדש: Deploy → Manage deployments → עריכה (עיפרון)
//    → Version: New version → Deploy. בלי זה השינוי לא עולה לאוויר!
//
// הגיליון "ערב מתנדבים" ייווצר אוטומטית ברישום הראשון.
// עמודות: תאריך | שעה | שם | כמות אנשים
// ============================================================

// Schema: תאריך(0), שעה(1), שם(2), כמות אנשים(3)
function registerVolunteerEvent(name, guests) {
  try {
    if (!name || !String(name).trim()) {
      return { success: false, message: 'חסר שם' };
    }

    const sheet = ensureVolunteerEventSheet_();
    const now = new Date();
    const date = Utilities.formatDate(now, 'Asia/Jerusalem', 'dd/MM/yyyy');
    const time = Utilities.formatDate(now, 'Asia/Jerusalem', 'HH:mm');
    const guestCount = Math.max(1, Math.min(20, parseInt(guests, 10) || 1));

    sheet.appendRow([date, time, String(name).trim(), guestCount]);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function ensureVolunteerEventSheet_() {
  // אם הסקריפט שלכם עצמאי (לא מחובר לגיליון) — החליפו את השורה הבאה ב:
  // const ss = SpreadsheetApp.openById('SHEET_ID_כאן');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('ערב מתנדבים');
  if (!sheet) {
    sheet = ss.insertSheet('ערב מתנדבים');
    sheet.appendRow(['תאריך', 'שעה', 'שם', 'כמות אנשים']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setRightToLeft(true);
  }
  return sheet;
}
