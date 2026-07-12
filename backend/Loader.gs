/**
 * העוזרת — loader ("הצינור")
 *
 * This is the ONLY code that lives in Ziv's Apps Script project. On every
 * request it pulls the real backend code from the GitHub repo and runs it,
 * so pushing to GitHub updates the backend with no manual redeploys.
 *
 * Speed: the fetched code is cached for up to 6 hours, so a normal request
 * pays ~10ms (cache read) + ~30ms (eval) instead of a ~half-second GitHub
 * fetch every single time. To make a fresh push go live instantly, hit
 *   <exec-url>?flush=1
 * once - that clears the cache so the next request re-fetches the latest.
 */

var CODE_URL = 'https://raw.githubusercontent.com/soundcomhr-hash/VA/main/backend/Code.gs';
var CODE_CACHE_KEY = 'remote_code';
var CODE_CACHE_TTL = 21600; // 6h (CacheService max)

function doGet(e) {
  if (e && e.parameter && e.parameter.flush) {
    CacheService.getScriptCache().remove(CODE_CACHE_KEY);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, flushed: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return runRemote_('doGet', [e]);
}
function doPost(e) { return runRemote_('doPost', [e]); }
function setup() { return runRemote_('setup', []); }
function setupTriggers() { return runRemote_('setupTriggers', []); }
function checkHours() { return runRemote_('checkHours', []); }
function checkMorningEscalation() { return runRemote_('checkMorningEscalation', []); }

function runRemote_(fnName, args) {
  var code = fetchCode_();
  eval(code);
  // The eval'd declarations shadow this file's globals inside this scope,
  // so eval(fnName) resolves to the remote implementation.
  return eval(fnName).apply(null, args);
}

function fetchCode_() {
  var cache = CacheService.getScriptCache();
  var code = cache.get(CODE_CACHE_KEY);
  if (code) return code;
  // Cache miss: pull fresh (query param defeats GitHub's own CDN cache).
  code = UrlFetchApp.fetch(CODE_URL + '?t=' + new Date().getTime(),
    { muteHttpExceptions: true }).getContentText();
  cache.put(CODE_CACHE_KEY, code, CODE_CACHE_TTL);
  return code;
}

// Never called. Google grants permissions by statically scanning this file's
// text, and it can't see through eval — these references force it to request
// the scopes the remote code actually needs.
function scopeHints_() {
  SpreadsheetApp.create('scope-hint');
  CalendarApp.getDefaultCalendar();
  DriveApp.createFolder('scope-hint');
  PropertiesService.getScriptProperties();
  Session.getScriptTimeZone();
  ScriptApp.newTrigger('scope-hint');
  MailApp.sendEmail('scope-hint', 'scope-hint', 'scope-hint');
}
