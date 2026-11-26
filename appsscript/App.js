/**
 * App.js — menu, one-click setup, schedule, web app and the API the UI calls.
 */

/** The model spreadsheet (works from the Sheet, the web app and triggers). */
function ss_() {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) return active;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet not linked. Open the Sheet and run Areta Comps › Set up.');
  return SpreadsheetApp.openById(id);
}

/* ------------------------------- Menu -------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_NAME)
    .addItem('Open dashboard', 'menuOpenSidebar')
    .addSeparator()
    .addItem('Refresh share prices now', 'menuRefreshPrices')
    .addItem('Import CIQ exports from inbox', 'menuImport')
    .addItem('Apply reviewed financial changes', 'menuApply')
    .addSeparator()
    .addItem('Run checks', 'menuChecks')
    .addItem('Build validation report', 'menuReport')
    .addItem('Repair formulas', 'menuRepair')
    .addSeparator()
    .addItem('Set up / reinstall schedule', 'menuSetup')
    .addItem('Open CIQ inbox folder', 'menuOpenInbox')
    .addItem('Open web app', 'menuOpenWebApp')
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
function menuReport() { var r = buildValidationReport(); toast_('Validation report updated.'); ss_().getSheetByName(r.sheet).activate(); }
function menuRepair() { toast_(repairFormulas().summary); }
function menuSetup() { var r = setup(); SpreadsheetApp.getUi().alert(APP_NAME + ' is ready', r.message, SpreadsheetApp.getUi().ButtonSet.OK); }
function menuOpenInbox() {
  var url = getFolders().inbox.getUrl();
  var html = HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Drop CIQ exports here:<br><a target="_blank" href="' + url + '">Open CIQ Inbox folder</a></p>').setWidth(320).setHeight(90);
  SpreadsheetApp.getUi().showModalDialog(html, 'CIQ inbox');
}

function menuOpenWebApp() {
  var url = webAppUrl_();
  var body = url
    ? '<a target="_blank" href="' + url + '">Open the Areta Comps web app</a><p style="color:#5F6368;font-size:12px">Signed in to several Google accounts? Open it in a private window signed in only with the account that owns this Sheet.</p>'
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
  var model = loadModel();
  var n = installFormulas(model);
  ensureConfigSheet(model.layout.companies);
  var folders = getFolders();
  sheetWithHeaders(SHEETS.EXCEPTIONS, EXC_HEADERS, { widths: [10, 80, 90, 170, 130, 120, 620, 260, 110, 110, 160, 90, 170, 260] });
  sheetWithHeaders(SHEETS.REVIEW, REVIEW_HEADERS, { widths: [60, 170, 130, 130, 110, 110, 80, 140, 230, 260, 150, 120, 260, 10, 10, 10, 10] });
  sheetWithHeaders(SHEETS.HISTORY, HISTORY_HEADERS, { widths: [100, 140, 160, 100, 80, 120, 130, 170, 150] });
  sheetWithHeaders(SHEETS.AUDIT, AUDIT_HEADERS, { widths: [140, 160, 180, 90, 170, 130, 120, 110, 110, 260, 150, 150, 260, 120, 300] });
  sheetWithHeaders(SHEETS.RUNS, RUN_HEADERS, { widths: [170, 150, 150, 200, 90, 90, 90, 700] });
  addReadMeSection_();
  orderTabs_();
  var trig = installTriggers();
  var run = startRun('setup', 'manual');
  finishRun(run, 'OK', 'Installed ' + n + ' formulas; ' + trig);
  return {
    message: 'Formulas installed (' + n + '). ' + trig + '\n\nCIQ inbox: ' + folders.inbox.getUrl(),
    inboxUrl: folders.inbox.getUrl()
  };
}

function orderTabs_() {
  var ss = ss_();
  var order = [SHEETS.COMPS, SHEETS.EXCEPTIONS, SHEETS.REVIEW, 'Validation Report', SHEETS.README, SHEETS.CONFIG, SHEETS.HISTORY, SHEETS.AUDIT, SHEETS.RUNS];
  var pos = 1;
  order.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(pos++); } });
  ss.setActiveSheet(ss.getSheetByName(SHEETS.COMPS));
}

/** Append the operating instructions to the Read Me tab (once). */
function addReadMeSection_() {
  var sh = ss_().getSheetByName(SHEETS.README);
  if (!sh) return;
  var vals = sh.getRange(1, 2, sh.getMaxRows(), 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).indexOf('Automation (Areta Comps engine)') === 0) return;
  var start = sh.getLastRow() + 2;
  var lines = [
    ['Automation (Areta Comps engine)'],
    ['Prices: refreshed automatically every week (Config tab sets day and hour) from Google Finance, cross-checked against a second public source, and saved to Price History. Prior price = the saved close one week earlier.'],
    ['Financials: drop new CIQ "Financials" exports into the CIQ inbox (menu › Open CIQ inbox folder, or upload in the dashboard), then Import → review → Apply. Companies without a new export keep their current values.'],
    ['Controls: every import rebuilds market cap and EV and reconciles them to CIQ. Exceptions tab lists ERROR (value not written), REVIEW (banker judgement) and INFO items; mark items Reviewed to sign them off.'],
    ['Traceability: Audit Trail records every value written with its source file, tab and cell. Run Log records every run. Notes column items starting "⚑" are written by the engine; your own notes are never overwritten.'],
    ['Formulas (black cells) are installed by the engine. If one is overwritten, Run checks flags it and Repair formulas restores it.']
  ];
  sh.getRange(start, 2, lines.length, 1).setValues(lines);
  sh.getRange(start, 2).setFontWeight('bold');
  sh.getRange(start + 1, 2, lines.length - 1, 1).setWrap(true);
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

function doGet(e) {
  var user = currentUser();
  if (!isAllowed_(user)) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:40px"><h2>Access restricted</h2><p>' + (user || 'Your account') + ' is not on the access list for this comps model. Ask the model owner to add you on the Config tab.</p></div>').setTitle(APP_NAME);
  }
  var t = HtmlService.createTemplateFromFile('ui/Index');
  t.mode = 'web';
  t.webAppUrl = webAppUrl_();
  return t.evaluate().setTitle('Areta Comps').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function isAllowed_(email) {
  var list = String(getSettings().allowedUsers || '').split(/[,;\s]+/).filter(String).map(function (x) { return x.toLowerCase(); });
  if (!list.length) return true;
  return list.indexOf(String(email || '').toLowerCase()) >= 0;
}

function webAppUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

function guard_() { if (!isAllowed_(currentUser())) throw new Error('Not authorised.'); }

/* -------------------- API used by ui/Index.html -------------------- */

function api_dashboard() {
  guard_();
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
    runs: readRuns(12)
  });
}

function api_refreshPrices() { guard_(); return refreshPrices({ trigger: 'manual' }); }
function api_upload(files) { guard_(); var names = saveUploads(files); return { saved: names }; }
function api_stageImport() { guard_(); return stageImport({ trigger: 'manual' }); }
function api_apply() { guard_(); return applyImport({ trigger: 'manual' }); }
function api_discard() { guard_(); return discardImport(); }
function api_setQueue(idxs, value) { guard_(); return setQueueApply(idxs, value); }
function api_checks() { guard_(); return runModelChecks({ trigger: 'manual' }); }
function api_repair() { guard_(); return repairFormulas(); }
function api_review(keys, comment) { guard_(); return { updated: markExceptionsReviewed(keys, comment) }; }
function api_report() { guard_(); var r = buildValidationReport(); return { url: ss_().getUrl() + '#gid=' + r.gid }; }
function api_setup() { guard_(); return setup(); }
