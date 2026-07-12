/**
 * העוזרת — backend (Google Apps Script web app)
 *
 * What this does:
 *  - Receives captures from the PWA (text and/or audio), logs EVERYTHING to
 *    the "קלט" sheet (audit trail — nothing is ever dropped).
 *  - Runs a deterministic Hebrew parser (closed vocabulary from the "מילון"
 *    sheet). The parser NEVER guesses missing info — it asks instead.
 *  - Creates inbox cards in the "ממתין" sheet. Nothing touches the calendar
 *    until Ziv taps אשר in the app (inbox_confirm).
 *
 * First-time setup: run the setup() function once from the editor (▶),
 * approve the permissions, then deploy as a web app.
 */

var TZ = 'Asia/Jerusalem';
var SPREADSHEET_NAME = 'העוזרת - נתונים';
var AUDIO_FOLDER_NAME = 'העוזרת - הקלטות';

// ---------------------------------------------------------------- setup ----

function setup() {
  var props = PropertiesService.getScriptProperties();

  var token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('TOKEN', token);
  }

  var ss = getSpreadsheet_();

  getOrCreateSheet_(ss, 'קלט', ['תאריך', 'טקסט', 'קישור אודיו', 'לתמלול', 'מזהה']);
  getOrCreateSheet_(ss, 'ממתין', ['מזהה', 'נוצר', 'סטטוס', 'סוג', 'תיאור', 'שאלה', 'תשובה', 'נתונים', 'טקסט מקורי']);
  getOrCreateSheet_(ss, 'משימות', TASKS_HEADERS);
  seedDictionary_(ss);

  var settings = getOrCreateSheet_(ss, 'הגדרות', ['מפתח', 'ערך']);
  upsertSetting_(settings, 'TOKEN (להעתיק לאפליקציה)', token);
  upsertSetting_(settings, 'אזור זמן של הסקריפט', Session.getScriptTimeZone());

  Logger.log('===============================');
  Logger.log('הקוד הסודי (token) שלך: ' + token);
  Logger.log('גיליון הנתונים: ' + ss.getUrl());
  if (Session.getScriptTimeZone() !== TZ) {
    Logger.log('אזהרה! אזור הזמן של הפרויקט הוא ' + Session.getScriptTimeZone() +
      ' — יש לשנות ל-Jerusalem בהגדרות הפרויקט (גלגל שיניים בצד).');
  }
  Logger.log('===============================');
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function upsertSetting_(sheet, key, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function seedDictionary_(ss) {
  var sheet = ss.getSheetByName('מילון');
  if (sheet) return; // never overwrite Ziv's edits
  sheet = ss.insertSheet('מילון');
  sheet.setRightToLeft(true);
  sheet.appendRow(['סוג', 'שם', 'כינויים נוספים (מופרדים בפסיק)']);
  sheet.setFrozenRows(1);
  var rows = [
    ['אדם', 'זיו', ''],
    ['אדם', 'חן', 'חן גבע'],
    ['אדם', 'ספיר', ''],
    ['אדם', 'נווה', ''],
    ['אדם', 'אבי', ''],
    ['מקום', 'בית מיכל', 'ראשון לציון'],
    ['מקום', 'מעון רוחמה', 'רוחמה, כפר סבא'],
    ['מקום', 'ארקדש', 'באר קדש, יהוד'],
    ['מקום', 'נווה האירוס', 'דוי האירוס, נווה אירוס'],
    ['מקום', 'אקים-כלנית', 'אקים, כלנית'],
    ['מקום', 'מעון הוד', 'הוד'],
    ['מקום', 'הגנים', ''],
    ['מקום', 'בית קסלר', 'קסלר'],
    ['מקום', 'טירת הכרמל', 'טירת כרמל'],
  ];
  rows.forEach(function (r) { sheet.appendRow(r); });
}

// ------------------------------------------------------------- web app ----

function doGet() {
  return jsonOut_({ ok: true, service: 'העוזרת', time: new Date().toISOString() });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'bad-json' });
  }

  var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || body.token !== token) {
    return jsonOut_({ ok: false, error: 'bad-token' });
  }

  try {
    switch (body.action) {
      case 'capture': return jsonOut_(capture_(body));
      case 'inbox_list': return jsonOut_(inboxList_());
      case 'inbox_answer': return jsonOut_(inboxAnswer_(body.id, body.answer, !!body.noPlace));
      case 'tasks_list': return jsonOut_(tasksList_());
      case 'tasks_toggle': return jsonOut_(tasksToggle_(body.id));
      case 'tasks_set_color': return jsonOut_(tasksSetColor_(body.id, body.color));
      case 'morning_start': return jsonOut_(morningStart_());
      case 'morning_snooze': return jsonOut_(morningSnooze_());
      case 'inbox_confirm': return jsonOut_(inboxConfirm_(body.id));
      case 'inbox_delete': return jsonOut_(inboxDelete_(body.id));
      case 'today': return jsonOut_(today_());
      default: return jsonOut_({ ok: false, error: 'unknown-action' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: 'server-error', detail: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------- capture ----

function capture_(body) {
  var ss = getSpreadsheet_();
  var id = Utilities.getUuid();
  var audioUrl = '';

  if (body.audioBase64) {
    var folder = getAudioFolder_();
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HH-mm-ss');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(body.audioBase64),
      body.audioMime || 'audio/webm',
      'הקלטה_' + stamp + '.webm'
    );
    audioUrl = folder.createFile(blob).getUrl();
  }

  var text = (body.text || '').trim();
  var needsTranscription = !!body.needsTranscription;

  // Audit trail first — a capture is never lost, whatever happens next.
  ss.getSheetByName('קלט').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    text, audioUrl, needsTranscription ? 'כן' : '', id,
  ]);

  var inbox = ss.getSheetByName('ממתין');

  if (needsTranscription || !text) {
    inbox.appendRow([
      id, now_(), 'פתוח', 'לתמלול',
      'הקלטת אודיו ממתינה לתמלול',
      'מה נאמר בהקלטה? (אפשר להקליד את התוכן כתשובה)',
      '', JSON.stringify({ audioUrl: audioUrl }), text,
    ]);
    return { ok: true, kind: 'לתמלול', summary: 'ההקלטה נשמרה ותסומן לתמלול' };
  }

  var parsed = parseHebrew_(text);

  // Tasks don't need Ziv's אשר - per his rule, only meetings truly need a
  // place/time slot filled. A task goes straight to the משימות list.
  if (parsed.kind === 'משימה') {
    var task = createTask_(text, parsed);
    return { ok: true, kind: 'משימה', summary: 'נוסף למשימות (' + task.color + ')' };
  }

  inbox.appendRow([
    id, now_(), 'פתוח', parsed.kind,
    parsed.summary, parsed.question, '',
    JSON.stringify(parsed), text,
  ]);
  return { ok: true, kind: parsed.kind, summary: parsed.summary, question: parsed.question };
}

function getAudioFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('AUDIO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recreate */ }
  }
  var folder = DriveApp.createFolder(AUDIO_FOLDER_NAME);
  props.setProperty('AUDIO_FOLDER_ID', folder.getId());
  return folder;
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

// -------------------------------------------------------------- parser ----
// Deterministic Hebrew slot-filler. Closed vocabulary from the מילון sheet.
// The one hard rule: NEVER guess money, dates, or people. Missing = ask.

function parseHebrew_(text) {
  var clean = text.replace(/[.,!?;"']/g, ' ').replace(/\s+/g, ' ').trim();

  var kind = detectIntent_(clean);
  var date = parseDate_(clean);
  var time = parseTime_(clean);
  var dict = loadDictionary_();
  var place = matchEntry_(clean, dict.places);
  var person = matchEntry_(clean, dict.people);

  var result = {
    kind: kind, text: text,
    dateISO: date ? Utilities.formatDate(date, TZ, 'yyyy-MM-dd') : null,
    dateHuman: date ? hebrewDate_(date) : null,
    time: time ? pad_(time.h) + ':' + pad_(time.m) : null,
    place: place, person: person,
    title: buildTitle_(kind, clean, person, place),
  };

  result.noPlaceConfirmed = false;
  return finalizeParse_(result);
}

function finalizeParse_(result) {
  var needDay = false, needTime = false, needPlace = false;
  if (result.kind === 'פגישה') {
    needDay = !result.dateISO;
    needTime = !result.time;
    needPlace = !result.place && !result.noPlaceConfirmed;
  } else if (result.kind === 'תזכורת') {
    needDay = !result.dateISO;
    needTime = !result.time;
  }
  result.question = buildQuestion_(needDay, needTime, needPlace);
  result.complete = !needDay && !needTime && !needPlace &&
    (result.kind === 'פגישה' || result.kind === 'תזכורת');
  result.summary = buildSummary_(result);
  return result;
}

// One short, natural sentence instead of several concatenated questions -
// per Ziv: don't pile up questions, ask the minimum in one line.
function buildQuestion_(needDay, needTime, needPlace) {
  if (!needDay && !needTime && !needPlace) return '';
  if (needDay && needTime && needPlace) return 'מתי ואיפה?';
  if (needDay && needTime) return 'מתי?';
  if (needDay && needPlace) return 'באיזה יום ואיפה?';
  if (needTime && needPlace) return 'באיזו שעה ואיפה?';
  if (needDay) return 'באיזה יום?';
  if (needTime) return 'באיזו שעה?';
  return 'איפה?';
}

function detectIntent_(text) {
  if (/(^|\s)(קבע|תקבע|לקבוע|קבעי|נקבע)/.test(text)) return 'פגישה';
  if (/(^|\s)(תזכיר|תזכורת|להזכיר|תזכירי)/.test(text)) return 'תזכורת';
  if (/(^|\s)(בטל|לבטל|תבטל|מבוטל)/.test(text)) return 'ביטול';
  if (/(^|\s)(העבר|הזז|תעביר|תזיז|לדחות|דחה)/.test(text)) return 'העברה';
  return 'משימה';
}

var DAY_NAMES = { 'ראשון': 0, 'שני': 1, 'שלישי': 2, 'רביעי': 3, 'חמישי': 4, 'שישי': 5, 'שבת': 6 };
var DAY_HUMAN = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

function parseDate_(text) {
  var today = startOfToday_();
  if (/מחרתיים/.test(text)) return addDays_(today, 2);
  if (/מחר/.test(text)) return addDays_(today, 1);
  if (/היום/.test(text)) return today;

  var m = text.match(/(?:ביום |יום |ב)?(ראשון|שני|שלישי|רביעי|חמישי|שישי|שבת)/);
  if (m) {
    var target = DAY_NAMES[m[1]];
    var dow = today.getDay(); // JS: Sunday=0 — matches DAY_NAMES
    var diff = (target - dow + 7) % 7;
    if (/שבוע הבא/.test(text)) diff += 7;
    return addDays_(today, diff);
  }
  return null;
}

var WORD_HOURS = {
  'אחת': 1, 'שתיים': 2, 'שלוש': 3, 'ארבע': 4, 'חמש': 5, 'שש': 6,
  'שבע': 7, 'שמונה': 8, 'תשע': 9, 'עשר': 10, 'אחת עשרה': 11, 'שתים עשרה': 12,
};

function parseTime_(text) {
  var h = null, m = 0, match;

  if ((match = text.match(/רבע ל[־-]?(\d{1,2})/)) ||
      (match = matchWordHour_(text, /רבע ל[־-]?([א-ת]+(?: עשרה)?)/))) {
    h = (typeof match === 'number' ? match : parseInt(match[1], 10)) - 1;
    m = 45;
  } else if ((match = text.match(/(\d{1,2})[:.](\d{2})/))) {
    h = parseInt(match[1], 10); m = parseInt(match[2], 10);
  } else if ((match = text.match(/(\d{1,2}) וחצי/))) {
    h = parseInt(match[1], 10); m = 30;
  } else if ((match = matchWordHour_(text, /([א-ת]+(?: עשרה)?) וחצי/))) {
    h = match; m = 30;
  } else if ((match = text.match(/(\d{1,2}) ורבע/))) {
    h = parseInt(match[1], 10); m = 15;
  } else if ((match = text.match(/בשעה (\d{1,2})/)) || (match = text.match(/(^|\s)ב[־-]?(\d{1,2})($|\s)/))) {
    h = parseInt(match[2] !== undefined ? match[2] : match[1], 10);
  } else if ((match = matchWordHour_(text, /בשעה ([א-ת]+(?: עשרה)?)/))) {
    h = match;
  }

  if (h === null || isNaN(h) || h > 23 || m > 59) return null;

  // Colloquial Hebrew: "ב-3" means 15:00, "בערב" bumps to evening.
  // Anything still odd is visible in the summary before Ziv confirms.
  var evening = /בערב|אחר הצהריים|אחה"?צ|בלילה/.test(text);
  var morning = /בבוקר/.test(text);
  if (!morning && h >= 1 && h <= 6) h += 12;
  else if (evening && h < 12) h += 12;

  return { h: h, m: m };
}

function matchWordHour_(text, regex) {
  var m = text.match(regex);
  if (m && WORD_HOURS.hasOwnProperty(m[1])) return WORD_HOURS[m[1]];
  return null;
}

function loadDictionary_() {
  var sheet = getSpreadsheet_().getSheetByName('מילון');
  var people = [], places = [];
  if (!sheet) return { people: people, places: places };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var type = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    if (!name) continue;
    var aliases = String(data[i][2] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    var entry = { name: name, aliases: [name].concat(aliases) };
    if (type === 'מקום') places.push(entry);
    else if (type === 'אדם') people.push(entry);
  }
  return { people: people, places: places };
}

// Fuzzy match: exact substring first, then Levenshtein ≤ 2 against every
// window of the same word-length (voice typos: "באר קדש" → "ארקדש").
function matchEntry_(text, entries) {
  for (var i = 0; i < entries.length; i++) {
    for (var j = 0; j < entries[i].aliases.length; j++) {
      if (text.indexOf(entries[i].aliases[j]) !== -1) return entries[i].name;
    }
  }
  var words = text.split(' ');
  for (var i = 0; i < entries.length; i++) {
    for (var j = 0; j < entries[i].aliases.length; j++) {
      var alias = entries[i].aliases[j];
      var len = alias.split(' ').length;
      for (var k = 0; k + len <= words.length; k++) {
        var window = words.slice(k, k + len).join(' ');
        if (window.length >= 3 && levenshtein_(window, alias) <= 2) return entries[i].name;
      }
    }
  }
  return null;
}

function levenshtein_(a, b) {
  if (Math.abs(a.length - b.length) > 2) return 99;
  var prev = [], cur = [];
  for (var j = 0; j <= b.length; j++) prev[j] = j;
  for (var i = 1; i <= a.length; i++) {
    cur = [i];
    for (var j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

function buildTitle_(kind, clean, person, place) {
  if (kind === 'פגישה' && person) return 'פגישה עם ' + person;
  if (kind === 'פגישה' && place) return 'מפגש ב' + place;
  if (kind === 'תזכורת') {
    var stripped = clean.replace(/(תזכיר לי|תזכיר|תזכורת|להזכיר)/g, '').trim();
    return 'תזכורת: ' + (stripped || clean).slice(0, 60);
  }
  return clean.slice(0, 60);
}

function buildSummary_(r) {
  var parts = [];
  if (r.kind === 'פגישה') parts.push('לקבוע: ' + r.title);
  else if (r.kind === 'תזכורת') parts.push(r.title);
  else if (r.kind === 'ביטול') parts.push('לבטל: ' + r.text);
  else if (r.kind === 'העברה') parts.push('להעביר: ' + r.text);
  else parts.push('משימה: ' + r.text);

  if (r.dateHuman) parts.push(r.dateHuman);
  if (r.time) parts.push('בשעה ' + r.time);
  if (r.place) parts.push(r.place.charAt(0) === 'ב' ? r.place : 'ב' + r.place);
  return parts.join(' · ');
}

function hebrewDate_(date) {
  return 'יום ' + DAY_HUMAN[date.getDay()] + ' ' + pad_(date.getDate()) + '/' + pad_(date.getMonth() + 1);
}

function startOfToday_() {
  var s = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var p = s.split('-');
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
}

function addDays_(d, n) {
  var r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }

// -------------------------------------------------------------- tasks -----
// Per Ziv: a task never blocks on אשר. It gets a color automatically
// (combination approach - auto rule first, he can override by hand):
//   אדום = due today, ירוק = sport/fun, כחול = routine, צהוב = default
// (open, unplanned - "anything written as a task that needs doing").
//
// Color is NEVER frozen at creation time - it's recomputed live on every
// list, because "due today" is a moving target (a task made yesterday for
// "tomorrow" must turn red on its own once tomorrow arrives). Column ד
// ("צבע ידני") only holds an explicit manual override; blank = auto.

var TASKS_HEADERS = ['מזהה', 'נוצר', 'טקסט', 'צבע ידני', 'סטטוס', 'תאריך יעד', 'שעה יעד'];
var SPORT_WORDS_ = /ספורט|כושר|חדר כושר|ריצה|לרוץ|שחייה|לשחות|יוגה|פילאטיס|אימון|לשחק|כיף|בילוי|טיול|לטייל/;
var ROUTINE_WORDS_ = /קבוע|שגרה|כל שבוע|כל יום|כל חודש|תמיד/;
var URGENT_WINDOW_MIN = 30;

function getTasksSheet_() {
  var sheet = getOrCreateSheet_(getSpreadsheet_(), 'משימות', TASKS_HEADERS);
  if (sheet.getLastColumn() < TASKS_HEADERS.length) {
    // One-time migration from the old frozen-color layout: any color that
    // was auto-assigned at creation gets cleared so it goes back to being
    // computed live (a real manual pick would have to be re-clicked once).
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 4, lastRow - 1, 1).clearContent();
    sheet.getRange(1, 1, 1, TASKS_HEADERS.length).setValues([TASKS_HEADERS]);
  }
  return sheet;
}

function classifyTaskColor_(text, dueDateISO) {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (dueDateISO === today) return 'אדום';
  if (SPORT_WORDS_.test(text)) return 'ירוק';
  if (ROUTINE_WORDS_.test(text)) return 'כחול';
  return 'צהוב';
}

function createTask_(text, parsed) {
  var sheet = getTasksSheet_();
  var id = Utilities.getUuid();
  // Tasks are low-risk and freely editable (unlike meetings, which must
  // never guess a date) - a bare time with no day mentioned ("לשלם ב-3")
  // defaults to today, since that's what it means in ordinary speech.
  var dueDate = parsed.dateISO;
  if (!dueDate && parsed.time) {
    dueDate = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  }
  var title = parsed.title || text;

  // Per Ziv (2026-07-13, explicit override of the "never write without אשר"
  // rule, scoped to dated tasks only): a task with a date must land on the
  // calendar the moment it's captured, or he forgets it entirely. Meetings
  // are unaffected - they still gate on place/time via ממתין/אשר.
  if (dueDate) createTaskCalendarEvent_(title, dueDate, parsed.time);

  sheet.appendRow([id, now_(), title, '', 'פתוח', dueDate || '', parsed.time || '']);
  return { id: id, color: classifyTaskColor_(text, dueDate) };
}

function createTaskCalendarEvent_(title, dueDateISO, dueTime) {
  var p = dueDateISO.split('-');
  var y = parseInt(p[0], 10), m = parseInt(p[1], 10) - 1, d = parseInt(p[2], 10);
  var cal = CalendarApp.getDefaultCalendar();
  if (dueTime) {
    var t = dueTime.split(':');
    var start = new Date(y, m, d, parseInt(t[0], 10), parseInt(t[1], 10));
    var end = new Date(start.getTime() + 30 * 60 * 1000);
    var event = cal.createEvent('✅ ' + title, start, end);
    event.addPopupReminder(0);
  } else {
    cal.createAllDayEvent('✅ ' + title, new Date(y, m, d));
  }
}

function findTaskRow_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return null;
}

function tasksList_() {
  var sheet = getTasksSheet_();
  var data = sheet.getDataRange().getValues();
  // Google Sheets silently converts an ISO string like "2026-07-13" written
  // to a cell into a real Date value, so reading it back gives a Date object
  // - NOT the "yyyy-MM-dd" string. That broke every dueDate === today check
  // and left dated tasks yellow. Recover the original string by formatting
  // the Date back in the spreadsheet's own timezone (a symmetric round-trip).
  var ssTz = getSpreadsheet_().getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var nowMin = parseInt(Utilities.formatDate(new Date(), TZ, 'H'), 10) * 60 +
               parseInt(Utilities.formatDate(new Date(), TZ, 'm'), 10);
  var items = [];
  for (var i = 1; i < data.length; i++) {
    var text = data[i][2];
    var manualColor = data[i][3];
    var dueDate = cellToText_(data[i][5], ssTz, 'yyyy-MM-dd');
    var dueTime = cellToText_(data[i][6], ssTz, 'HH:mm');
    var urgent = false;
    if (dueDate === today && dueTime) {
      var t = dueTime.split(':');
      var dueMin = parseInt(t[0], 10) * 60 + parseInt(t[1], 10);
      urgent = (dueMin - nowMin) <= URGENT_WINDOW_MIN;
    }
    items.push({
      id: data[i][0], created: String(data[i][1]), text: text,
      color: manualColor || classifyTaskColor_(text, dueDate),
      status: data[i][4], dueDate: dueDate, dueTime: dueTime || null, urgent: urgent,
    });
  }
  items.reverse();
  return { ok: true, items: items };
}

// Undo Sheets' auto-conversion: a cell we wrote as an ISO string may read
// back as a Date. Format it in the same timezone it was stored in to recover
// the exact original string; pass plain strings straight through.
function cellToText_(val, tz, fmt) {
  if (val === '' || val === null || val === undefined) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, tz, fmt);
  }
  return String(val).trim();
}

function tasksToggle_(id) {
  var sheet = getTasksSheet_();
  var row = findTaskRow_(sheet, id);
  if (!row) return { ok: false, error: 'not-found' };
  var next = sheet.getRange(row, 5).getValue() === 'בוצע' ? 'פתוח' : 'בוצע';
  sheet.getRange(row, 5).setValue(next);
  return { ok: true, status: next };
}

function tasksSetColor_(id, color) {
  var valid = ['אדום', 'כחול', 'ירוק', 'צהוב'];
  if (valid.indexOf(color) === -1) return { ok: false, error: 'bad-color' };
  var sheet = getTasksSheet_();
  var row = findTaskRow_(sheet, id);
  if (!row) return { ok: false, error: 'not-found' };
  sheet.getRange(row, 4).setValue(color);
  return { ok: true };
}

// -------------------------------------------------- morning checklist -----
// "בוקר טוב" is a button Ziv presses (not automatic). Pressing it starts a
// 2-hour clock; an hourly trigger (checkMorningEscalation) nags by email
// every hour after that while tasks remain open, until he marks them done
// or taps "השתק להיום" - stored per-calendar-day so it resets on its own.

function morningStart_() {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var settings = getOrCreateSheet_(getSpreadsheet_(), 'הגדרות', ['מפתח', 'ערך']);
  upsertSetting_(settings, 'בוקר_' + today, new Date().toISOString());
  // בוקר טוב is the morning ritual for TODAY's tasks specifically - per Ziv,
  // not the whole open backlog (that's what the משימות tab + רענן is for).
  var all = tasksList_();
  all.items = all.items.filter(function (t) { return t.color === 'אדום'; });
  return all;
}

function morningSnooze_() {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var settings = getOrCreateSheet_(getSpreadsheet_(), 'הגדרות', ['מפתח', 'ערך']);
  upsertSetting_(settings, 'השתקה_' + today, 'כן');
  return { ok: true };
}

function checkMorningEscalation() {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var startedAt = getSetting_('בוקר_' + today);
  if (!startedAt) return; // בוקר טוב not pressed yet today
  if (getSetting_('השתקה_' + today) === 'כן') return; // silenced for today

  var elapsedMs = new Date().getTime() - new Date(startedAt).getTime();
  if (elapsedMs < 2 * 60 * 60 * 1000) return;

  var open = tasksList_().items.filter(function (t) {
    return t.status !== 'בוצע' && t.color === 'אדום';
  });
  if (!open.length) return;

  var lines = open.map(function (t) { return '• [' + t.color + '] ' + t.text; });
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    '⏰ עדיין יש משימות פתוחות',
    'משימות שעדיין לא סומנו כבוצעו:\n\n' + lines.join('\n') +
    '\n\nלהפסיק את התזכורות להיום: פתח את העוזרת, לשונית משימות, "השתק להיום".');
}

// --------------------------------------------------------------- inbox ----

var INBOX_COL = { id: 0, created: 1, status: 2, kind: 3, summary: 4, question: 5, answer: 6, data: 7, original: 8 };

function inboxList_() {
  var sheet = getSpreadsheet_().getSheetByName('ממתין');
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][INBOX_COL.status] !== 'פתוח') continue;
    var extra = {};
    try { extra = JSON.parse(data[i][INBOX_COL.data] || '{}'); } catch (e) {}
    items.push({
      id: data[i][INBOX_COL.id],
      created: String(data[i][INBOX_COL.created]),
      kind: data[i][INBOX_COL.kind],
      summary: data[i][INBOX_COL.summary],
      question: data[i][INBOX_COL.question],
      complete: !!extra.complete,
      audioUrl: extra.audioUrl || null,
    });
  }
  items.reverse(); // newest first
  return { ok: true, items: items };
}

function findInboxRow_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][INBOX_COL.id] === id) return { row: i + 1, values: data[i] };
  }
  return null;
}

function inboxAnswer_(id, answer, noPlace) {
  var sheet = getSpreadsheet_().getSheetByName('ממתין');
  var found = findInboxRow_(sheet, id);
  if (!found) return { ok: false, error: 'not-found' };

  var original = found.values[INBOX_COL.original] || '';
  var newText = (original + ' ' + (answer || '')).trim();
  var parsed = parseHebrew_(newText);

  // When the ONLY thing still missing is the place and the dictionary didn't
  // recognize the answer, the answer IS the place — Ziv answered "איפה?",
  // so any free text is accepted verbatim (his rule: be flexible on places).
  if (parsed.kind === 'פגישה' && !parsed.place && parsed.dateISO && parsed.time && answer) {
    parsed.place = String(answer).trim();
    parsed.title = buildTitle_(parsed.kind, newText, parsed.person, parsed.place);
  }

  // "אין מקום" - Ziv explicitly confirmed there's no place for this one;
  // stop asking and let it through without a place.
  if (noPlace) parsed.noPlaceConfirmed = true;
  finalizeParse_(parsed);

  sheet.getRange(found.row, INBOX_COL.kind + 1).setValue(parsed.kind);
  sheet.getRange(found.row, INBOX_COL.summary + 1).setValue(parsed.summary);
  sheet.getRange(found.row, INBOX_COL.question + 1).setValue(parsed.question);
  sheet.getRange(found.row, INBOX_COL.answer + 1).setValue(answer);
  sheet.getRange(found.row, INBOX_COL.data + 1).setValue(JSON.stringify(parsed));
  sheet.getRange(found.row, INBOX_COL.original + 1).setValue(newText);

  return { ok: true, summary: parsed.summary, question: parsed.question, complete: parsed.complete };
}

function inboxConfirm_(id) {
  var sheet = getSpreadsheet_().getSheetByName('ממתין');
  var found = findInboxRow_(sheet, id);
  if (!found) return { ok: false, error: 'not-found' };
  if (found.values[INBOX_COL.status] !== 'פתוח') return { ok: false, error: 'already-closed' };

  var parsed = {};
  try { parsed = JSON.parse(found.values[INBOX_COL.data] || '{}'); } catch (e) {}

  var eventLink = null;
  if (parsed.kind === 'פגישה' || parsed.kind === 'תזכורת') {
    if (!parsed.complete) {
      return { ok: false, error: 'incomplete', question: found.values[INBOX_COL.question] };
    }
    var p = parsed.dateISO.split('-');
    var t = parsed.time.split(':');
    var start = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10),
                         parseInt(t[0], 10), parseInt(t[1], 10));
    var end = new Date(start.getTime() + 60 * 60 * 1000);
    var event = CalendarApp.getDefaultCalendar().createEvent(parsed.title, start, end, {
      location: parsed.place || '',
      description: 'נוצר על ידי העוזרת מתוך: "' + (parsed.text || '') + '"',
    });
    // Reminders pop on the phone: at the moment itself for תזכורת,
    // 30 minutes ahead for a meeting.
    event.addPopupReminder(parsed.kind === 'תזכורת' ? 0 : 30);
    eventLink = 'נוצר אירוע: ' + parsed.title + ' — ' + parsed.dateHuman + ' ' + parsed.time;
  }

  sheet.getRange(found.row, INBOX_COL.status + 1).setValue('אושר');
  return { ok: true, message: eventLink || 'סומן כמאושר' };
}

function inboxDelete_(id) {
  var sheet = getSpreadsheet_().getSheetByName('ממתין');
  var found = findInboxRow_(sheet, id);
  if (!found) return { ok: false, error: 'not-found' };
  sheet.getRange(found.row, INBOX_COL.status + 1).setValue('נמחק');
  return { ok: true };
}

// ------------------------------------------------- hours-filling nudge ----
// Runs every evening via a time trigger (see setupTriggers). If today's
// calendar had 🎵 session events, emails Ziv a reminder to fill his hours.
// Days are configurable in the הגדרות sheet (Hebrew day letters, e.g. א,ב,ג,ד,ה).

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'checkHours' || fn === 'checkMorningEscalation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkHours')
    .timeBased().everyDays(1).atHour(19).nearMinute(30).inTimezone(TZ)
    .create();
  ScriptApp.newTrigger('checkMorningEscalation')
    .timeBased().everyHours(1).inTimezone(TZ)
    .create();

  var settings = getOrCreateSheet_(getSpreadsheet_(), 'הגדרות', ['מפתח', 'ערך']);
  if (!getSetting_('ימי תזכורת שעות')) upsertSetting_(settings, 'ימי תזכורת שעות', 'א,ב,ג,ד,ה');
  upsertSetting_(settings, 'קישור מילוי שעות', getSetting_('קישור מילוי שעות') || '');

  Logger.log('התזכורת הופעלה: כל ערב בסביבות 19:30 (ימים לפי לשונית הגדרות).');
}

var HEB_DAY_LETTERS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש']; // Sunday..Saturday

// Ziv's hours-filling Google Form: entry IDs extracted from the live form.
// A prefilled link pre-answers date + guide + place, so the branching form
// jumps straight to that place's patient list after two "הבא" taps.
var HOURS_FORM = {
  base: 'https://docs.google.com/forms/d/e/1FAIpQLSeMaQ35HjQWits59-hqgQZYhdXjzcTV80T-nrVpk2ozi4ZtOg/viewform',
  dateEntry: 'entry.1556154575',
  guideEntry: 'entry.2037107479',
  guideValue: 'זיו ויסברג',
  placeEntry: 'entry.974072686',
  places: ['ארקדש', 'בית מיכל', 'מעון הוד', 'אקים-כלנית', 'תיכון חדש', 'רוחמה', 'נווה האירוס', 'נהריה'],
};

function hoursFormLink_(dateISO, place) {
  var url = HOURS_FORM.base + '?usp=pp_url' +
    '&' + HOURS_FORM.dateEntry + '=' + dateISO +
    '&' + HOURS_FORM.guideEntry + '=' + encodeURIComponent(HOURS_FORM.guideValue);
  if (place) url += '&' + HOURS_FORM.placeEntry + '=' + encodeURIComponent(place);
  return url;
}

function matchFormPlace_(eventTitle) {
  for (var i = 0; i < HOURS_FORM.places.length; i++) {
    if (eventTitle.indexOf(HOURS_FORM.places[i]) !== -1) return HOURS_FORM.places[i];
  }
  return null;
}

function checkHours() {
  var days = String(getSetting_('ימי תזכורת שעות') || 'א,ב,ג,ד,ה')
    .split(',').map(function (s) { return s.trim(); });
  var todayLetter = HEB_DAY_LETTERS[startOfToday_().getDay()];
  if (days.indexOf(todayLetter) === -1) return;

  var start = startOfToday_();
  var sessions = CalendarApp.getDefaultCalendar().getEvents(start, addDays_(start, 1))
    .filter(function (ev) { return ev.getTitle().indexOf('🎵') === 0; });
  if (!sessions.length) return;

  var dateISO = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
  var lines = sessions.map(function (ev) {
    var line = '• ' + Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm') + ' — ' + ev.getTitle();
    var place = matchFormPlace_(ev.getTitle());
    line += '\n  📝 מילוי שעות (הכל כבר מסומן, רק ללחוץ הבא-הבא):\n  ' + hoursFormLink_(dateISO, place);
    return line;
  });
  var body = 'היו לך היום ' + sessions.length + ' מפגשים:\n\n' + lines.join('\n\n') +
    '\n\nאל תשכח למלא שעות!';

  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    '🎵 תזכורת: מילוי שעות להיום', body);
}

function getSetting_(key) {
  var sheet = getSpreadsheet_().getSheetByName('הגדרות');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// --------------------------------------------------------------- today ----

function today_() {
  var start = startOfToday_();
  var end = addDays_(start, 1);
  var events = CalendarApp.getDefaultCalendar().getEvents(start, end).map(function (ev) {
    return {
      title: ev.getTitle(),
      start: Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'),
      end: Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm'),
      allDay: ev.isAllDayEvent(),
      location: ev.getLocation(),
    };
  });
  // Per Ziv: open ממתין items are "mandatory things to resolve today" even
  // before their own date is known - surface them here so they can't be
  // quietly forgotten in a tab he doesn't open.
  var pending = inboxList_().items;
  return { ok: true, events: events, pending: pending };
}
