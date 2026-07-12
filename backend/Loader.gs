/**
 * העוזרת — loader ("הצינור")
 *
 * This is the ONLY code that lives in Ziv's Apps Script project. On every
 * request it pulls the real backend code from the GitHub repo and runs it,
 * so pushing to GitHub updates the backend with no manual redeploys.
 * A cache-busting query param defeats GitHub's own CDN caching, so a push
 * takes effect on the very next request instead of after several minutes.
 */

var CODE_URL = 'https://raw.githubusercontent.com/soundcomhr-hash/VA/main/backend/Code.gs';

function doGet(e) { return runRemote_('doGet', [e]); }
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
  var url = CODE_URL + '?t=' + new Date().getTime();
  return UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
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
