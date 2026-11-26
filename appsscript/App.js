/**
 * App.js — menu, one-click setup, schedule, web app and the API the UI calls.
 */

/** The model spreadsheet (works from the Sheet, the web app and triggers). */
function ss_() {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) return active;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet not linked. Open the Sheet and run Comparables Automation › Set up.');
  return SpreadsheetApp.openById(id);
}

/* ------------------------------- Menu -------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_NAME)
    .addItem('Open dashboard', 'menuOpenSidebar')
    .addSeparator()
    .addItem('1. Import CIQ exports (from inbox)', 'menuImport')
    .addItem('2. Apply financial changes', 'menuApply')
    .addItem('3. Refresh share prices', 'menuRefreshPrices')
    .addItem('4. Run checks', 'menuChecks')
    .addItem('5. Build validation report', 'menuReport')
    .addSeparator()
    .addItem(RUN_ALL_LABEL, 'menuRunAll')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Tools')
      .addItem('Open CIQ inbox folder', 'menuOpenInbox')
      .addItem('Open web app', 'menuOpenWebApp')
      .addItem('Set web app password', 'menuSetWebPassword')
      .addItem('Repair formulas', 'menuRepair')
      .addItem('Set up / reinstall schedule', 'menuSetup'))
    .addToUi();
}

function menuOpenSidebar() {
  var t = HtmlService.createTemplateFromFile('ui/Index');
  t.mode = 'sidebar';
  t.webAppUrl = webAppUrl_();
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle(APP_NAME));
}
function menuRefreshPrices() { toast_(refreshPrices({ trigger: 'manual' }).summary); }
function menuImport() { var r = stageImport({ trigger: 'manual' }); toast_(r.summary || r.message); var sh = ss_().getSheetByName(SHEETS.REVIEW); if (sh) sh.activate(); }
function menuApply() { toast_(applyImport({ trigger: 'manual' }).summary || 'Nothing to apply.'); }
function menuChecks() { toast_(runModelChecks({ trigger: 'manual' }).summary); var sh = ss_().getSheetByName(SHEETS.EXCEPTIONS); if (sh) sh.activate(); }
function menuReport() { var r = buildValidationReport(); toast_('Validation report rebuilt' + (r.pdfUrl ? '; dated PDF saved to Drive › Reports' : '') + '.'); ss_().getSheetByName(r.sheet).activate(); }
function menuRunAll() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Run all steps', 'This will:\n\n1. Import any CIQ exports in the inbox\n2. Apply every change that passed validation (Ready and Review). Blocked values are never written.\n3. Refresh share prices\n4. Run checks\n5. Build the validation report\n\nItems needing judgement stay open on the Exceptions tab for sign-off. Continue?', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var r = runAll({ trigger: 'manual' });
  ui.alert('Run all: done', r.lines.join('\n'), ui.ButtonSet.OK);
  ss_().getSheetByName(REPORT_SHEET).activate();
}

/**
 * Steps 1-5 in order. Review items are applied (they are flagged, not wrong);
 * Blocked items are never written. Stops cleanly if a step fails.
 */
function runAll(opts) {
  opts = opts || {};
  var lines = [];
  function step(label, fn) {
    try { var r = fn(); lines.push('✓ ' + label + (r ? ': ' + r : '')); return true; }
    catch (e) { lines.push('✗ ' + label + ': ' + (e && e.message || e)); return false; }
  }
  var staged = null;
  step('1. Import', function () { staged = stageImport({ trigger: opts.trigger || 'run-all' }); return staged.summary || staged.message; });
  step('2. Apply', function () {
    if (!staged || !staged.files) return 'nothing to apply';
    if (staged.applied) return staged.applied.summary;
    return applyImport({ trigger: opts.trigger || 'run-all', skipChecks: true }).summary;
  });
  step('3. Prices', function () { return refreshPrices({ trigger: opts.trigger || 'run-all', skipChecks: true }).summary; });
  step('4. Checks', function () { return runModelChecks({ trigger: opts.trigger || 'run-all' }).summary; });
  step('5. Report', function () { buildValidationReport(); return 'Validation Report tab rebuilt and PDF saved to Drive'; });
  return { lines: lines };
}

function menuRepair() { toast_(repairFormulas().summary); }
function menuSetup() { var r = setup(); SpreadsheetApp.getUi().alert(APP_NAME + ' is ready', r.message, SpreadsheetApp.getUi().ButtonSet.OK); }
function menuOpenInbox() {
  var url = getFolders().inbox.getUrl();
  var html = HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Drop CIQ exports here:<br><a target="_blank" href="' + url + '">Open CIQ Inbox folder</a></p>').setWidth(320).setHeight(90);
  SpreadsheetApp.getUi().showModalDialog(html, 'CIQ inbox');
}

function menuOpenWebApp() {
  var url = portalUrl_();
  var body = url
    ? '<a target="_blank" href="' + url + '">Open the Comparables Automation web app</a><p style="color:#5F6368;font-size:12px">' + (PropertiesService.getScriptProperties().getProperty('WEB_PW_HASH') ? 'Sign in with the web app password.' : 'Set a password first: Tools › Set web app password.') + '</p>'
    : 'The web app is not deployed for this copy. See README › Independent test, step 5.';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<div style="font-family:sans-serif">' + body + '</div>').setWidth(360).setHeight(130), 'Web app');
}

function toast_(msg) { try { ss_().toast(msg, APP_NAME, 8); } catch (e) { /* no UI */ } }

/* ------------------------------- Setup ------------------------------- */

/** One-click, idempotent setup. */
function setup() {
  var ss = ss_();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SPREADSHEET_ID') && props.getProperty('SPREADSHEET_ID') !== ss.getId()) {
    // This is a copy of the Sheet: start with its own Drive folders, not the original's.
    ['FOLDER_ROOT', 'FOLDER_INBOX', 'FOLDER_PROCESSED', 'PENDING_BATCH', 'SCHEDULE_TEXT'].forEach(function (k) { props.deleteProperty(k); });
  }
  props.setProperty('SPREADSHEET_ID', ss.getId());
  var tzWanted = String(getSettings().timeZone || 'Europe/London');
  if (ss.getSpreadsheetTimeZone() !== tzWanted) ss.setSpreadsheetTimeZone(tzWanted);
  var model = loadModel();
  var n = installFormulas(model);
  ensureConfigSheet(model.layout.companies);
  var folders = getFolders();
  sheetWithHeaders(SHEETS.EXCEPTIONS, EXC_HEADERS, { widths: [10, 80, 90, 170, 130, 120, 620, 260, 110, 110, 160, 90, 170, 260] });
  sheetWithHeaders(SHEETS.REVIEW, REVIEW_HEADERS, { widths: [60, 170, 130, 130, 110, 110, 80, 140, 230, 260, 150, 120, 260, 10, 10, 10, 10] });
  sheetWithHeaders(SHEETS.HISTORY, HISTORY_HEADERS, { widths: [100, 140, 160, 100, 80, 120, 130, 170, 150] });
  sheetWithHeaders(SHEETS.AUDIT, AUDIT_HEADERS, { widths: [140, 160, 180, 90, 170, 130, 120, 110, 110, 260, 150, 150, 260, 120, 300] });
  sheetWithHeaders(SHEETS.RUNS, RUN_HEADERS, { widths: [170, 150, 150, 200, 90, 90, 90, 700] });
  styleComps_(model);
  buildReadMe_();
  if (!readReviewQueue().length) showQueueEmpty_(null);
  polishDataTabs_();
  try { buildValidationReport({ noPdf: true }); } catch (e) { /* report is rebuilt by step 5 */ }
  buildStartHere_();
  orderTabs_();
  var trig = installTriggers();
  var run = startRun('setup', 'manual');
  finishRun(run, 'OK', 'Installed ' + n + ' formulas; ' + trig);
  var pwNote = props.getProperty('WEB_PW_HASH') ? '' : '\n\nWeb dashboard: set its password with Comparables Automation › Tools › Set web app password.';
  return {
    message: 'Formulas installed (' + n + '). ' + trig + '\n\nCIQ inbox: ' + folders.inbox.getUrl() + pwNote,
    inboxUrl: folders.inbox.getUrl()
  };
}

function orderTabs_() {
  var ss = ss_();
  var order = [START_SHEET, SHEETS.COMPS, SHEETS.EXCEPTIONS, SHEETS.REVIEW, 'Validation Report', SHEETS.README, SHEETS.CONFIG, SHEETS.HISTORY, SHEETS.AUDIT, SHEETS.RUNS];
  var pos = 1;
  order.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(pos++); } });
  ss.setActiveSheet(ss.getSheetByName(START_SHEET) || ss.getSheetByName(SHEETS.COMPS));
}

/** (Re)install the time-driven triggers from the Config tab. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'scheduledWeeklyPrices' || f === 'scheduledDailySnapshot') ScriptApp.deleteTrigger(t);
  });
  var s = getSettings();
  var day = String(s.scheduleDay || 'MONDAY').toUpperCase();
  if (!ScriptApp.WeekDay[day]) day = 'MONDAY';
  var hour = Math.max(0, Math.min(23, Number(s.scheduleHour) || 7));
  ScriptApp.newTrigger('scheduledWeeklyPrices').timeBased().onWeekDay(ScriptApp.WeekDay[day]).atHour(hour).create();
  var msg = 'Weekly price refresh scheduled for ' + day.charAt(0) + day.slice(1).toLowerCase() + 's at ' + hour + ':00 (' + ss_().getSpreadsheetTimeZone() + ').';
  if (String(s.dailySnapshot).toUpperCase() === 'YES') {
    ScriptApp.newTrigger('scheduledDailySnapshot').timeBased().everyDays(1).atHour(22).create();
    msg += ' Daily price snapshot at 22:00 on weekdays.';
  }
  PropertiesService.getScriptProperties().setProperty('SCHEDULE_TEXT', msg);
  return msg;
}

/* ------------------------------ Web app ------------------------------ */
/*
 * The web app is published "Anyone" (no Google sign-in) and runs as the deploying
 * account. Google's sign-in hop is what breaks with several signed-in accounts
 * ("Page not found"), so the portal uses its own password instead:
 *   - password set from the Sheet (Comparables Automation › Tools › Set web app password),
 *     stored only as a salted SHA-256 hash in Script Properties;
 *   - login returns a signed token (HMAC-SHA256, 12 h expiry) that every API call carries;
 *   - from the Sheet sidebar no token is needed: the caller is a signed-in editor.
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, app: APP_NAME, passwordSet: !!PropertiesService.getScriptProperties().getProperty('WEB_PW_HASH') })).setMimeType(ContentService.MimeType.JSON);
  }
  var t = HtmlService.createTemplateFromFile('ui/Index');
  t.mode = 'web';
  t.webAppUrl = webAppUrl_();
  t.hasPassword = !!PropertiesService.getScriptProperties().getProperty('WEB_PW_HASH');
  return t.evaluate().setTitle('Areta Comparables Automation').setFaviconUrl('https://alycz.github.io/Areta-Comparables-Automation/assets/favicon-32.png').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * JSON API for the standalone portal (docs/portal, hosted outside Google).
 * The portal calls this with fetch() and no cookies, so Google's multi-account
 * routing ("Sorry, unable to open the file", "Page not found") never applies.
 * Body: {"fn": "api_dashboard", "args": [token, ...]}. Same password/token gate.
 */
var API_ = {
  api_login: function (a) { return api_login(a[0], a[1]); },
  api_dashboard: function (a) { return api_dashboard(a[0]); },
  api_refreshPrices: function (a) { return api_refreshPrices(a[0]); },
  api_upload: function (a) { return api_upload(a[0], a[1]); },
  api_stageImport: function (a) { return api_stageImport(a[0]); },
  api_apply: function (a) { return api_apply(a[0]); },
  api_discard: function (a) { return api_discard(a[0]); },
  api_setQueue: function (a) { return api_setQueue(a[0], a[1], a[2]); },
  api_checks: function (a) { return api_checks(a[0]); },
  api_repair: function (a) { return api_repair(a[0]); },
  api_review: function (a) { return api_review(a[0], a[1], a[2]); },
  api_report: function (a) { return api_report(a[0]); },
  api_runAll: function (a) { return api_runAll(a[0]); }
};

function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_[body.fn];
    if (!fn) throw new Error('Unknown call.');
    var args = body.args || [];
    if (body.fn !== 'api_login' && !args[0]) throw new Error('SESSION_EXPIRED');
    out = { ok: true, result: fn(args) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function isAllowed_(email) {
  var list = String(getSettings().allowedUsers || '').split(/[,;\s]+/).filter(String).map(function (x) { return x.toLowerCase(); });
  if (!list.length) return true;
  return list.indexOf(String(email || '').toLowerCase()) >= 0;
}

/** Public dashboard (GitHub Pages; talks to doPost). Override with Script Property PORTAL_URL for another host. */
var DEFAULT_PORTAL_URL = 'https://alycz.github.io/Areta-Comparables-Automation/';
function portalUrl_() { try { return PropertiesService.getScriptProperties().getProperty('PORTAL_URL') || DEFAULT_PORTAL_URL; } catch (e) { return DEFAULT_PORTAL_URL; } }

function webAppUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function pwHash_(pw, salt) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)); }
function tokenSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('WEB_TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('WEB_TOKEN_SECRET', s); }
  return s;
}
function sign_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, tokenSecret_())); }

/** Set (or change) the web app password. Changing it signs everyone out. */
function setWebPassword_(pw) {
  pw = String(pw || '');
  if (pw.length < 8) throw new Error('Use at least 8 characters.');
  var props = PropertiesService.getScriptProperties();
  var salt = Utilities.getUuid();
  props.setProperty('WEB_PW_SALT', salt);
  props.setProperty('WEB_PW_HASH', pwHash_(pw, salt));
  props.deleteProperty('WEB_TOKEN_SECRET');
}

function menuSetWebPassword() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Web app password', 'Choose a password for the web dashboard (at least 8 characters). Share it only with people who should run the model. Changing it signs everyone out.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try { setWebPassword_(r.getResponseText()); } catch (e) { ui.alert(e.message); return; }
  ui.alert('Password set', 'Open the web app here:\n\n' + portalUrl_() + '\n\nIt works in any browser, signed in to Google or not.', ui.ButtonSet.OK);
}

/** Web login. Returns a token, or throws. */
function api_login(name, pw) {
  var props = PropertiesService.getScriptProperties();
  var hash = props.getProperty('WEB_PW_HASH'), salt = props.getProperty('WEB_PW_SALT');
  if (!hash) throw new Error('No password set yet. In the Sheet: Comparables Automation › Tools › Set web app password.');
  Utilities.sleep(400); // slows guessing
  if (pwHash_(String(pw || ''), salt) !== hash) throw new Error('Wrong password.');
  name = String(name || '').trim().slice(0, 80) || 'web user';
  if (!isAllowed_(name) && /@/.test(name)) throw new Error(name + ' is not on the access list (Config › allowedUsers).');
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({ u: name, exp: Date.now() + 12 * 3600 * 1000 }));
  return { token: payload + '.' + sign_(payload), user: name };
}

var WEB_USER_ = null;

/** Every API call: a valid web token, or a signed-in editor using the Sheet sidebar. */
function guard_(token) {
  if (token) {
    var parts = String(token).split('.');
    if (parts.length === 2 && sign_(parts[0]) === parts[1]) {
      var p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
      if (p.exp > Date.now()) { WEB_USER_ = p.u + ' (web)'; return; }
    }
    throw new Error('SESSION_EXPIRED');
  }
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) { email = ''; }
  if (!email) throw new Error('SESSION_EXPIRED');
  if (!isAllowed_(email)) throw new Error('Not authorised.');
}

/* -------------------- API used by ui/Index.html -------------------- */

function api_dashboard(token) {
  guard_(token);
  var ss = ss_();
  var ex = readExceptions();
  var open = ex.filter(function (e) { return e.status !== 'Reviewed'; });
  var count = function (sev) { return open.filter(function (e) { return e.severity === sev; }).length; };
  var props = PropertiesService.getScriptProperties();
  var folders = null;
  try { folders = getFolders(); } catch (e) { folders = null; }
  var inboxCount = 0;
  if (folders) { var it = folders.inbox.getFiles(); while (it.hasNext()) { it.next(); inboxCount++; } }
  return JSON.stringify({
    user: currentUser(),
    sheetName: ss.getName(),
    sheetUrl: ss.getUrl(),
    inboxUrl: folders ? folders.inbox.getUrl() : '',
    inboxCount: inboxCount,
    schedule: props.getProperty('SCHEDULE_TEXT') || 'Not scheduled yet: run Set up.',
    snapshot: modelSnapshot(),
    exceptions: ex,
    counts: { error: count('ERROR'), review: count('REVIEW'), info: count('INFO'), reviewed: ex.length - open.length },
    queue: readReviewQueue(),
    queueStatus: props.getProperty('QUEUE_STATUS') || '',
    runs: readRuns(12)
  });
}

function api_refreshPrices(token) { guard_(token); return refreshPrices({ trigger: 'manual' }); }
function api_upload(token, files) { guard_(token); var names = saveUploads(files); return { saved: names }; }
function api_stageImport(token) { guard_(token); return stageImport({ trigger: 'manual' }); }
function api_apply(token) { guard_(token); return applyImport({ trigger: 'manual' }); }
function api_discard(token) { guard_(token); return discardImport(); }
function api_setQueue(token, idxs, value) { guard_(token); return setQueueApply(idxs, value); }
function api_checks(token) { guard_(token); return runModelChecks({ trigger: 'manual' }); }
function api_repair(token) { guard_(token); return repairFormulas(); }
function api_review(token, keys, comment) { guard_(token); return { updated: markExceptionsReviewed(keys, comment) }; }
function api_report(token) { guard_(token); var r = buildValidationReport(); return { url: ss_().getUrl() + '#gid=' + r.gid, pdfUrl: r.pdfUrl }; }
function api_runAll(token) { guard_(token); var r = runAll({ trigger: 'manual' }); return { summary: r.lines.join(' · ') }; }
function api_setup(token) { guard_(token); return setup(); }
