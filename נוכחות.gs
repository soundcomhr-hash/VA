/**
 * נוכחות — סאונדקום
 * ===================
 * הסקריפט קורא את טאבי הטפסים (שמות שמסתיימים ב-[P] או [A]) וכותב טאב שטוח בשם "נתונים":
 * תאריך | מקום | קבוצה | שם | סוג | מדריך | חותמת זמן   — שורה אחת לכל אדם בכל מפגש.
 *
 * [P] = השמות שסומנו נכחו.
 * [A] = השמות שסומנו נעדרו; הנוכחים = רשימת הקבוצה (מהטאב "רשימות") פחות הנעדרים.
 *       בטאב [A] כל עמודת קבוצה נחשבת כמפגש שהתקיים בכל אחד מהתאריכים בשורה.
 * מקום = שם הטאב בלי הסיומת. קבוצה = כותרת העמודה בטופס, בדיוק כפי שהיא.
 *
 * בנוסף, פעולת "מילוי נוסחאות בטאב 3" כותבת נוסחאות בעמודות "תאריכים" ו"מספר מפגשים"
 * של טאב 3 בלבד (לעולם לא בשמות/מחירים/סכומים), שמושכות מ"נתונים" לפי החודש שבתא B1.
 * הטאבים 1/2 וטאבי הטפסים לא נגעים לעולם.
 */

var TAB_DATA   = 'נתונים';
var TAB_ROSTER = 'רשימות';   // מקום | קבוצה | שמות (מופרדים בפסיק) — נדרש לטאבי [A]
var TAB_PLACES = 'מקומות';   // מקום | קישור לטופס | אימייל המדריך — נדרש לתזכורות
var REMINDER_HOUR = 20;      // שעת שליחת התזכורות בערב (0-23)

var DATA_HEADERS = ['תאריך', 'מקום', 'קבוצה', 'שם', 'סוג', 'מדריך', 'חותמת זמן'];

var TAB_REPORT = '3'; // טאב הדוח החודשי; החודש נבחר בתא B1 שלו
var TAB3_LAST = 35;   // שורות 2-35 מנוהלות ע"י המערכת; שורה 36 = ספירת מקור אוטומטית
var PRICE_ROW = 37;   // שורת הסכומים המוסכמת: כאן שמים את נוסחת המחיר של כל בלוק (הדשבורד קורא מכאן). שורה 38 ומטה — חופשי לגמרי
// הדשבורד יושב בעמודה A של טאב 3 (שורות 2-9) ומתעדכן בכל ריענון

/**
 * הבלוקים בטאב 3: לכל בלוק — עמודות השם (אחת או שתיים: פרטי+משפחה), עמודת התאריכים,
 * עמודת מספר המפגשים, ואיך מסננים מ"נתונים": place = טקסט שחייב להופיע בעמודת "מקום",
 * type = ערך עמודת "סוג" ("קבוצה"/"פרטני"; ריק = בלי סינון סוג).
 * הוספתם בלוק חדש בטאב 3? הוסיפו כאן שורה מתאימה.
 */
var TAB3_BLOCKS = [
  { names: ['B', 'C'], dates: 'D',  count: 'E',  place: 'נווה האירוס', type: '' },
  { names: ['H', 'I'], dates: 'J',  count: 'K',  place: 'כלנית',       type: '' },
  { names: ['L', 'M'], dates: 'N',  count: 'O',  place: 'רוחמה',       type: '' },
  { names: ['R', 'S'], dates: 'T',  count: 'U',  place: 'בית מיכל',    type: '' },
  { names: ['X'],      dates: 'Y',  count: 'Z',  place: 'מעון הוד',    type: 'קבוצה' },
  { names: ['AB'],     dates: 'AC', count: 'AD', place: 'מעון הוד',    type: 'פרטני' }
];

// ===== תפריט =====

function onOpen() {
  SpreadsheetApp.getUi().createMenu('נוכחות')
    .addItem('התקנה (להריץ פעם אחת)', 'setupTriggers')
    .addItem('בנייה מחדש של "נתונים"', 'rebuildAll')
    .addItem('מילוי נוסחאות בטאב 3', 'fillTab3')
    .addItem('בדיקת תזכורות עכשיו', 'sendReminders')
    .addToUi();
}

/** יוצר את הטריגרים: עדכון "נתונים" בכל שליחת טופס + תזכורת יומית בערב. בטוח להריץ שוב. */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'onFormSubmitTrigger' || f === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive()).onFormSubmit().create();
  ScriptApp.newTrigger('sendReminders')
    .timeBased().everyDays(1).atHour(REMINDER_HOUR).create();
  readPlaces_(SpreadsheetApp.getActive()); // יוצר את "מקומות" אם חסר
  rebuildAll();                            // יוצר את "נתונים" ו"רשימות" אם חסרים
  fillTab3();                              // כותב את הנוסחאות בטאב 3
  SpreadsheetApp.getActive().toast('ההתקנה הושלמה. בדקו את הטאבים "רשימות" ו"מקומות".', 'נוכחות');
}

function onFormSubmitTrigger(e) {
  rebuildAll();
}

/** עריכה בטאב 3 מרעננת את הנוסחאות; עריכה ב"רשימות" בונה מחדש את "נתונים" — הכל אוטומטי. */
function onEdit(e) {
  if (!e || !e.range) return;
  var name = e.range.getSheet().getName();
  if (name === TAB_REPORT) fillTab3();
  else if (name === TAB_ROSTER) rebuildAll();
}

// ===== בניית "נתונים" =====

/** מוחק וכותב מחדש את "נתונים" מכל טאבי הטפסים. לא נוגע בטאבי הטפסים עצמם. */
function rebuildAll() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive();
    var roster = readRoster_(ss);
    var rows = [];
    var warnings = [];

    ss.getSheets().forEach(function (sheet) {
      var src = sourceInfo_(sheet.getName());
      if (!src) return; // לא טאב מקור ([P]/[A]) — מדלגים
      collectRows_(sheet, src, roster, rows, warnings);
    });

    var out = ss.getSheetByName(TAB_DATA) || ss.insertSheet(TAB_DATA);
    out.clearContents();
    out.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
    if (rows.length) {
      out.getRange(2, 1, rows.length, DATA_HEADERS.length).setValues(rows);
      out.getRange(2, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy');
      out.getRange(2, 7, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    }

    var unique = warnings.filter(function (w, i) { return warnings.indexOf(w) === i; });
    unique.forEach(function (w) { Logger.log(w); });
    if (unique.length) ss.toast(unique.join(' | '), 'נוכחות — חסר ברשימות', 15);
  } finally {
    lock.releaseLock();
  }
}

/** קורא טאב מקור אחד ומוסיף שורות ל-rows. */
function collectRows_(sheet, src, roster, rows, warnings) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // זיהוי עמודות לפי הכותרת: חותמת זמן / מדריך / תאריך... / כל השאר = קבוצות
  var head = data[0].map(function (h) { return String(h).trim(); });
  var tsCol = -1, guideCol = -1, dateCols = [], groupCols = [];
  head.forEach(function (h, i) {
    if (!h) return;
    if (h.indexOf('חותמת') === 0) tsCol = i;
    else if (h.indexOf('מדריך') === 0) guideCol = i;
    else if (h.indexOf('תאריך') === 0) dateCols.push(i);
    else groupCols.push(i);
  });

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dates = [];
    dateCols.forEach(function (c) { if (row[c] !== '') dates.push(row[c]); });
    if (!dates.length) continue;
    var guide = guideCol >= 0 ? String(row[guideCol]).trim() : '';
    var stamp = tsCol >= 0 ? row[tsCol] : '';

    groupCols.forEach(function (c) {
      var group = head[c];
      var listed = splitNames_(row[c]);
      var present;
      if (src.kind === 'P') {
        if (!listed.length) return; // עמודה ריקה בטופס נוכחים = לא התקיים מפגש
        present = listed;
      } else { // [A]: הרשימה המלאה פחות מי שסומן נעדר (ריק = כולם נכחו)
        var full = roster[src.place + '|' + group];
        if (!full) {
          warnings.push('חסרה שורה ב"' + TAB_ROSTER + '" עבור: ' + src.place + ' / ' + group);
          return;
        }
        present = full.filter(function (n) { return listed.indexOf(n) === -1; });
      }
      var kind = group.indexOf('פרט') >= 0 ? 'פרטני' : 'קבוצה';
      dates.forEach(function (d) {
        present.forEach(function (name) {
          rows.push([d, src.place, group, name, kind, guide, stamp]);
        });
      });
    });
  }
}

// ===== מילוי טאב 3 =====

/**
 * כותב נוסחאות בעמודות "תאריכים" ו"מספר מפגשים" של כל בלוק בטאב 3, בשורות שבהן יש שם.
 * הנוסחאות מושכות מ"נתונים" לפי החודש שבתא B1, ומתעדכנות לבד (אין צורך להריץ שוב
 * אחרי כל טופס). ההתאמה היא לפי תחילת השם, ולכן "מזל י" בטאב 3 יתפוס את "מזל יעקובי"
 * מהטופס. הוספתם אנשים חדשים לטאב 3? הריצו שוב מהתפריט.
 * שימו לב: בעמודות התאריכים/המפגשים, שורות 2-35 ושורה 36 (ספירת מקור) נכתבות מחדש
 * בכל ריענון. שורה 37 = שורת הסכומים שלכם (הדשבורד קורא ממנה), שורה 38 ומטה וכל
 * עמודה אחרת (מחירים, סכומים, AA, AE...) — חופשי, המערכת לא נוגעת שם לעולם.
 */
function fillTab3() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TAB_REPORT);
  if (!sh) { ss.toast('לא נמצא טאב "' + TAB_REPORT + '"', 'נוכחות'); return; }
  var last = TAB3_LAST;
  var vals = sh.getRange(1, 1, last, colIdx_('AD') + 1).getValues();

  TAB3_BLOCKS.forEach(function (b) {
    var dateFs = [], countFs = [];
    for (var r = 2; r <= last; r++) {
      var hasName = b.names.some(function (col) {
        return String(vals[r - 1][colIdx_(col)]).trim() !== '';
      });
      if (!hasName) { dateFs.push(['']); countFs.push(['']); continue; }

      // השם המלא כפי שכתוב בטאב 3 (פרטי או פרטי+משפחה)
      var n = b.names.length === 1
        ? 'TRIM($' + b.names[0] + r + ')'
        : 'TRIM($' + b.names[0] + r + '&" "&$' + b.names[1] + r + ')';
      var conds =
        'LEFT(\'' + TAB_DATA + '\'!$D$2:$D,LEN(' + n + '))=' + n +
        ',MONTH(\'' + TAB_DATA + '\'!$A$2:$A)=($B$1+0)' + // +0 מכריח את B1 למספר, גם אם הוקלד כטקסט
        ',ISNUMBER(SEARCH("' + b.place + '",\'' + TAB_DATA + '\'!$B$2:$B))' +
        (b.type ? ',\'' + TAB_DATA + '\'!$E$2:$E="' + b.type + '"' : '');
      dateFs.push(['=IFERROR(TEXTJOIN(",",TRUE,ARRAYFORMULA(TEXT(SORT(UNIQUE(FILTER(\'' +
        TAB_DATA + '\'!$A$2:$A,' + conds + '))),"dd"))),"")']);
      var dc = '$' + b.dates + r;
      countFs.push(['=IF(' + dc + '="","",LEN(' + dc + ')-LEN(SUBSTITUTE(' + dc + ',",",""))+1)']);
    }
    sh.getRange(b.dates + '2:' + b.dates + last).setFormulas(dateFs);
    sh.getRange(b.count + '2:' + b.count + last).setFormulas(countFs);
  });

  // שורה 36, מתחת לכל בלוק: כמה רישומים נכנסו מהטפסים ("נתונים") לבלוק הזה בחודש הנבחר.
  // משווים מול הסכום בשורה 37 שלכם — שווה = הכל נכנס.
  var ctlRow = TAB3_LAST + 1;
  TAB3_BLOCKS.forEach(function (b) {
    var src = 'SUMPRODUCT((MONTH(\'' + TAB_DATA + '\'!$A$2:$A)=($B$1+0))' +
      '*ISNUMBER(SEARCH("' + b.place + '",\'' + TAB_DATA + '\'!$B$2:$B))' +
      (b.type ? '*(\'' + TAB_DATA + '\'!$E$2:$E="' + b.type + '")' : '') +
      '*(\'' + TAB_DATA + '\'!$A$2:$A<>""))';
    sh.getRange(b.dates + ctlRow).setValue('מהטפסים:');
    sh.getRange(b.count + ctlRow).setFormula('=' + src);
  });

  // ניקוי חד-פעמי: אם שורת הבקרה של גרסה קודמת (41) עדיין שם — מסירים אותה
  TAB3_BLOCKS.forEach(function (b) {
    if (sh.getRange(b.dates + '41').getValue() === 'מהטפסים:') {
      sh.getRange(b.dates + '41').clearContent();
      sh.getRange(b.count + '41').clearContent();
    }
  });

  // דשבורד בעמודה A, ליד בורר החודש: הסכום של כל בלוק (משורה 37) + סה"כ.
  // עמודה A מורחבת אוטומטית כדי ששמות המקומות ייכנסו.
  sh.setColumnWidth(1, 220);
  sh.getRange('A2').setValue('סיכום החודש:');
  var dashFs = [], refs = [];
  TAB3_BLOCKS.forEach(function (b) {
    var ref = '$' + b.count + '$' + PRICE_ROW;
    var label = b.place + (b.type ? ' - ' + (b.type === 'קבוצה' ? 'קבוצתי' : b.type) : '');
    dashFs.push(['="' + label + ': "&N(' + ref + ')']);
    refs.push('N(' + ref + ')');
  });
  dashFs.push(['="סה""כ: "&(' + refs.join('+') + ')']);
  sh.getRange(3, 1, dashFs.length, 1).setFormulas(dashFs);
}

/** 'A'->0, 'B'->1 ... 'AB'->27 */
function colIdx_(letters) {
  var idx = 0;
  for (var i = 0; i < letters.length; i++) idx = idx * 26 + (letters.charCodeAt(i) - 64);
  return idx - 1;
}

// ===== תזכורות במייל =====

/** רץ כל ערב: לכל אירוע של אתמול ביומן שהכותרת שלו מכילה מקום מ"מקומות" —
 *  אם אין ל"נתונים" אף שורה למקום+תאריך, נשלחת תזכורת למדריך. */
function sendReminders() {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone();
  var places = readPlaces_(ss);
  if (!places.length) return;

  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1); // אתמול 00:00
  var end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var dateStr = Utilities.formatDate(start, tz, 'dd/MM/yyyy');

  var reported = {}; // מקומות שכבר יש להם דיווח לאתמול
  var dataSheet = ss.getSheetByName(TAB_DATA);
  if (dataSheet && dataSheet.getLastRow() > 1) {
    dataSheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (fmtDate_(r[0], tz) === dateStr) reported[String(r[1]).trim()] = true;
    });
  }

  var titles = CalendarApp.getDefaultCalendar().getEvents(start, end)
    .map(function (ev) { return ev.getTitle(); });

  places.forEach(function (p) {
    var hadEvent = titles.some(function (t) { return t.indexOf(p.place) !== -1; });
    if (!hadEvent || reported[p.place]) return;
    if (!p.email) { Logger.log('חסר אימייל בטאב "' + TAB_PLACES + '" עבור: ' + p.place); return; }
    MailApp.sendEmail({
      to: p.email,
      subject: 'תזכורת: חסר דיווח נוכחות — ' + p.place + ' ' + dateStr,
      htmlBody:
        '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:16px">' +
        'שלום,<br><br>אתמול (' + dateStr + ') התקיימה פעילות ב<b>' + p.place +
        '</b> ועדיין לא התקבל דיווח נוכחות.<br><br>' +
        '<a href="' + p.url + '" style="font-size:22px;font-weight:bold">למילוי הטופס — לחצו כאן</a>' +
        '<br><br>תודה!</div>'
    });
  });
}

// ===== עזר =====

/** האם השם הוא טאב מקור? מחזיר {place, kind} עבור "מקום [P]" / "מקום [A]" (גם בלי סוגריים). */
function sourceInfo_(name) {
  var m = String(name).trim().match(/^(.+?)\s*(?:\[([PA])\]|\s([PA]))$/);
  return m ? { place: m[1].trim(), kind: m[2] || m[3] } : null;
}

/** "שם א, שם ב" -> ['שם א','שם ב'] */
function splitNames_(v) {
  return String(v || '').split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s; });
}

/** תאריך (Date או טקסט) -> 'dd/MM/yyyy' להשוואה. */
function fmtDate_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'dd/MM/yyyy') : String(v).trim();
}

/** קורא את "רשימות" למפה של מקום|קבוצה -> [שמות]. יוצר את הטאב אם חסר. */
function readRoster_(ss) {
  var sh = ss.getSheetByName(TAB_ROSTER);
  if (!sh) {
    sh = ss.insertSheet(TAB_ROSTER);
    sh.getRange(1, 1, 1, 3).setValues([['מקום', 'קבוצה', 'שמות']]);
    sh.getRange('A1').setNote(
      'שורה לכל קבוצה:\n' +
      'מקום = שם טאב הטופס בלי הסיומת [A] (למשל: מעון הוד-טירה)\n' +
      'קבוצה = בדיוק כמו כותרת העמודה בטופס (למשל: מעון הוד-קבוצה)\n' +
      'שמות = כל חברי הקבוצה, מופרדים בפסיק, באותו איות כמו בטופס');
    Logger.log('נוצר טאב "' + TAB_ROSTER + '" — יש למלא: מקום | קבוצה | שמות (מופרדים בפסיק).');
    ss.toast('נוצר טאב "' + TAB_ROSTER + '" — יש למלא אותו (ראו הערה בתא A1).', 'נוכחות', 15);
    return {};
  }
  var map = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var place = String(r[0]).trim(), group = String(r[1]).trim();
    if (place && group) map[place + '|' + group] = splitNames_(r[2]);
  });
  return map;
}

/** קורא את "מקומות" לרשימת {place, url, email}. יוצר את הטאב אם חסר. */
function readPlaces_(ss) {
  var sh = ss.getSheetByName(TAB_PLACES);
  if (!sh) {
    sh = ss.insertSheet(TAB_PLACES);
    sh.getRange(1, 1, 2, 3).setValues([
      ['מקום', 'קישור לטופס', 'אימייל המדריך'],
      ['בית מיכל', 'https://docs.google.com/forms/d/e/1FAIpQLSc5MUbDqLBiVSIjkf4tftYc6rUtWJWXpmeAVXmmhYs9SQIAfA/viewform', '']
    ]);
    sh.getRange('A1').setNote(
      'שורה לכל מקום:\n' +
      'מקום = כפי שמופיע בכותרת האירוע ביומן ובשם טאב הטופס\n' +
      'קישור לטופס = כתובת המילוי של הטופס\n' +
      'אימייל המדריך = לאן לשלוח את התזכורת');
    Logger.log('נוצר טאב "' + TAB_PLACES + '" — יש להשלים קישור ואימייל לכל מקום.');
    ss.toast('נוצר טאב "' + TAB_PLACES + '" — יש להשלים אימייל לכל מקום (ראו הערה בתא A1).', 'נוכחות', 15);
  }
  return sh.getDataRange().getValues().slice(1).map(function (r) {
    return { place: String(r[0]).trim(), url: String(r[1]).trim(), email: String(r[2]).trim() };
  }).filter(function (p) { return p.place; });
}
