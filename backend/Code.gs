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
      case 'inbox_answer': return jsonOut_(inboxAnswer_(body.id, body.answer));
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

  return finalizeParse_(result);
}

function finalizeParse_(result) {
  var missing = [];
  if (result.kind === 'פגישה') {
    if (!result.dateISO) missing.push('באיזה יום?');
    if (!result.time) missing.push('באיזו שעה?');
    if (!result.place) missing.push('איפה?');
  } else if (result.kind === 'תזכורת') {
    if (!result.dateISO) missing.push('באיזה יום להזכיר?');
    if (!result.time) missing.push('באיזו שעה להזכיר?');
  }
  result.question = missing.join(' ');
  result.complete = missing.length === 0 && (result.kind === 'פגישה' || result.kind === 'תזכורת');
  result.summary = buildSummary_(result);
  return result;
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

function inboxAnswer_(id, answer) {
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
    finalizeParse_(parsed);
  }

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
  return { ok: true, events: events };
}
