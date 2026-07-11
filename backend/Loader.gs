/**
 * העוזרת — loader ("הצינור")
 *
 * This is the ONLY code that lives in Ziv's Apps Script project. On every
 * request it pulls the real backend code from the GitHub repo and runs it,
 * so pushing to GitHub updates the backend with no manual redeploys.
 * (Cached for 2 minutes; GitHub's raw CDN adds up to ~5 more.)
 */

var CODE_URL = 'https://raw.githubusercontent.com/soundcomhr-hash/VA/main/backend/Code.gs';

function doGet(e) { return runRemote_('doGet', [e]); }
function doPost(e) { return runRemote_('doPost', [e]); }
function setup() { return runRemote_('setup', []); }

function runRemote_(fnName, args) {
  var code = fetchCode_();
  eval(code);
  // The eval'd declarations shadow this file's globals inside this scope,
  // so eval(fnName) resolves to the remote implementation.
  return eval(fnName).apply(null, args);
}

function fetchCode_() {
  var cache = CacheService.getScriptCache();
  var code = cache.get('remote_code');
  if (code) return code;
  code = UrlFetchApp.fetch(CODE_URL).getContentText();
  cache.put('remote_code', code, 120);
  return code;
}
