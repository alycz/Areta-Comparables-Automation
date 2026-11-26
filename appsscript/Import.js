/**
 * Import.js — automated financial updates from Capital IQ exports.
 *
 * 1. Banker drops CIQ "Financials" exports (.xls/.xlsx) into the Drive inbox
 *    folder, or uploads them in the web app.
 * 2. stageImport(): converts each file to a temporary Google Sheet, parses it
 *    (Core.parseCiqWorkbook), matches it to a model company by ticker, validates
 *    it, and writes the proposed changes to the Review Queue tab.
 * 3. applyImport(): writes approved values into the blue input cells, logs every
 *    change to the Audit Trail and moves the files to the Processed folder.
 *
 * Companies without a new export keep their existing values.
 */

var REVIEW_HEADERS = ['Apply?', 'Company', 'Ticker', 'Field', 'Current value', 'New value', 'Change', 'Status', 'Flags', 'Source file', 'Source tab', 'Source cell', 'Source label', 'Batch ID', 'File ID', 'Row', 'Field key'];

var FOLDER_KEYS = { root: 'FOLDER_ROOT', inbox: 'FOLDER_INBOX', processed: 'FOLDER_PROCESSED' };

/** Drive folders: "Areta Comps / CIQ Inbox" and "Areta Comps / Processed". */
function getFolders() {
  var props = PropertiesService.getScriptProperties();
  function getOrMake(key, name, parent) {
    var id = props.getProperty(key);
    if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* recreate */ } }
    // Re-use a folder of the same name if one exists (e.g. pre-created by the installer).
    var found = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
    var folder = null;
    while (found.hasNext() && !folder) { var cand = found.next(); if (!cand.isTrashed()) folder = cand; }
    if (!folder) folder = parent ? parent.createFolder(name) : DriveApp.createFolder(name);
    props.setProperty(key, folder.getId());
    return folder;
  }
  var root = getOrMake(FOLDER_KEYS.root, 'Areta Comps — ' + ss_().getName());
  return { root: root, inbox: getOrMake(FOLDER_KEYS.inbox, 'CIQ Inbox', root), processed: getOrMake(FOLDER_KEYS.processed, 'Processed', root) };
}

/** Save uploaded files (base64) from the web app into the inbox. */
function saveUploads(files) {
  var inbox = getFolders().inbox;
  return files.map(function (f) {
    var blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType || 'application/vnd.ms-excel', f.name);
    return inbox.createFile(blob).getName();
  });
}

var SHEET_MIME = 'application/vnd.google-apps.spreadsheet';

/** Read every tab of a CIQ export (xls / xlsx / Google Sheet) as {tab: values}. */
function readWorkbookFile(file) {
  var mime = file.getMimeType();
  var tempId = null, ssId = file.getId();
  if (mime !== SHEET_MIME) {
    // Excel content type by extension: files uploaded by browsers or APIs are often octet-stream,
    // which Drive refuses to convert.
    var blob = file.getBlob();
    if (/\.xls$/i.test(file.getName())) blob.setContentType('application/vnd.ms-excel');
    if (/\.xlsx$/i.test(file.getName())) blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    var converted = Drive.Files.create({ name: '[temp] ' + file.getName(), mimeType: SHEET_MIME, parents: [getFolders().root.getId()] }, blob);
    tempId = converted.id; ssId = converted.id;
  }
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var out = {};
    ss.getSheets().forEach(function (sh) { out[sh.getName()] = sh.getDataRange().getValues(); });
    return out;
  } finally {
    if (tempId) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) { /* ignore */ } }
  }
}

/**
 * Stage everything in the inbox. Returns a summary for the UI.
 * opts.trigger: 'manual' | 'upload'
 */
function stageImport(opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another update is running. Try again in a minute.');
  var run = startRun('import', opts.trigger);
  try {
    var settings = getSettings();
    var model = loadModel();
    var folders = getFolders();
    var tz = ss_().getSpreadsheetTimeZone();
    var todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var files = [];
    var it = folders.inbox.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (/\.(xls|xlsx)$/i.test(f.getName()) || f.getMimeType() === SHEET_MIME) files.push(f);
    }
    if (!files.length) {
      finishRun(run, 'OK', 'Inbox empty: nothing to import.');
      return { runId: run.id, files: 0, message: 'The CIQ inbox is empty. Upload exports first.' };
    }

    // Parse all files first so peer-period checks see the whole batch.
    var recs = files.map(function (f) {
      var rec;
      try { rec = parseCiqWorkbook(readWorkbookFile(f), f.getName()); }
      catch (e) { rec = { fileName: f.getName(), identity: {}, fields: {}, sources: {}, ciq: {}, meta: {}, issues: [{ severity: 'ERROR', code: 'READ_FAILED', field: '', message: 'Could not read file: ' + e.message }] }; }
      rec.fileId = f.getId();
      return rec;
    });

    // Peer period = most common LTM end date across the batch + model.
    var periods = recs.map(function (r) { return r.meta.period && r.meta.period.endDate; });
    var peerPeriod = mode(periods);

    // Match to model companies; resolve duplicates (latest LTM wins).
    var byRow = {}, issues = [], unmatched = [];
    recs.forEach(function (rec) {
      var m = matchCompany(model.layout.companies, rec.identity.ticker);
      if (!m.company) {
        unmatched.push(rec);
        issues.push({ severity: 'ERROR', category: 'Import', company: rec.identity.name || '', ticker: rec.identity.ticker || '', field: '', code: 'UNMATCHED_FILE',
          message: (m.how === 'ambiguous' ? 'Ticker matches more than one model row' : 'No company in the model has ticker ' + rec.identity.ticker) + '. File not imported. Add the company to the Comps tab or fix the ticker.', source: rec.fileName });
        return;
      }
      rec.company = m.company;
      if (m.how === 'symbol') rec.issues.push({ severity: 'INFO', code: 'TICKER_EXCHANGE_DIFFERS', field: '', message: 'Matched on symbol: export says ' + rec.identity.ticker + ', model says ' + m.company.ticker + '.' });
      var prev = byRow[m.company.row];
      if (prev) {
        var newer = (rec.meta.period && rec.meta.period.endDate || '') >= (prev.meta.period && prev.meta.period.endDate || '') ? rec : prev;
        var older = newer === rec ? prev : rec;
        issues.push({ severity: 'REVIEW', category: 'Import', company: m.company.name, ticker: m.company.ticker, field: '', code: 'DUPLICATE_EXPORT',
          message: 'Two exports for this company. Using ' + newer.fileName + ' (LTM ' + (newer.meta.period ? newer.meta.period.endDate : '?') + '); ignored ' + older.fileName + '.', source: older.fileName });
        older.ignored = true;
        byRow[m.company.row] = newer;
      } else byRow[m.company.row] = rec;
    });

    // Validate and build the review queue.
    var queue = [], batchId = run.id;
    Object.keys(byRow).forEach(function (row) {
      var rec = byRow[row];
      var cur = rowValues(model, +row);
      validateFinancials(rec, { todayIso: todayIso, peerPeriod: peerPeriod, existing: cur, settings: settings });
      var fieldErr = {};
      rec.issues.forEach(function (i) {
        if (i.severity === 'ERROR' && i.field) fieldErr[i.field] = true;
        issues.push({ severity: i.severity, category: 'Financials', company: rec.company.name, ticker: rec.company.ticker, field: i.field, code: i.code, message: i.message,
          source: rec.fileName + (rec.sources[i.field] ? ' › ' + rec.sources[i.field].sheet + '!' + rec.sources[i.field].cell : '') });
      });
      var flagsFor = function (field) {
        return rec.issues.filter(function (i) { return i.field === field && i.severity !== 'INFO'; }).map(function (i) { return i.code; }).join(', ');
      };
      FINANCIAL_FIELDS.forEach(function (field) {
        if (model.map.cols[field] === undefined) return;
        var has = rec.fields.hasOwnProperty(field);
        var newV = has ? rec.fields[field] : '';
        var oldV = cur[field];
        var blocked = !has || fieldErr[field];
        var same = has && sameValue(oldV, newV);
        var status = blocked ? 'Blocked: keep current' : same ? 'No change' : (flagsFor(field) ? 'Review' : 'Ready');
        var s = rec.sources[field] || {};
        queue.push([!blocked && !same, rec.company.name, rec.company.ticker, FIELD_LABELS[field], oldV, newV, changeText(oldV, newV), status, flagsFor(field),
          rec.fileName, s.sheet || '', s.cell || '', s.label || '', batchId, rec.fileId, +row + 1, field]);
      });
    });

    // Companies with no new export: keep values, note it.
    var covered = Object.keys(byRow).map(Number);
    var notCovered = model.layout.companies.filter(function (c) { return covered.indexOf(c.row) < 0; });
    notCovered.forEach(function (c) {
      issues.push({ severity: 'INFO', category: 'Import', company: c.name, ticker: c.ticker, field: '', code: 'NO_NEW_EXPORT', message: 'No new export in this batch: existing financials retained (' + (rowValues(model, c.row).financialPeriod || 'no period recorded') + ').', source: '' });
    });

    // Write the queue.
    var sh = sheetWithHeaders(SHEETS.REVIEW, REVIEW_HEADERS, { widths: [60, 170, 130, 130, 110, 110, 80, 140, 230, 260, 150, 120, 260, 10, 10, 10, 10] });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, REVIEW_HEADERS.length).clearContent().clearFormat().clearDataValidations();
    if (queue.length) {
      sh.getRange(2, 1, queue.length, REVIEW_HEADERS.length).setValues(queue);
      sh.getRange(2, 1, queue.length, 1).insertCheckboxes();
      var bg = queue.map(function (q) { var s = q[7]; return [s.indexOf('Blocked') === 0 ? '#FCE8E6' : s === 'Review' ? '#FFF4CE' : s === 'Ready' ? '#E6F4EA' : '#FFFFFF']; });
      sh.getRange(2, 8, queue.length, 1).setBackgrounds(bg);
      sh.getRange(2, 5, queue.length, 2).setNumberFormats(queue.map(function (q) {
        var f = q[16] === 'revenueGrowth' ? '0.0%;(0.0%)' : '#,##0.0;(#,##0.0)';
        return [f, f];
      }));
    }
    sh.hideColumns(14, 4);
    syncExceptions(run, ['Import', 'Financials'], issues, Object.keys(byRow).map(function (r) { return byRow[r].company.ticker; }));

    var counts = { ready: 0, review: 0, blocked: 0, same: 0 };
    queue.forEach(function (q) { var s = q[7]; if (s === 'Ready') counts.ready++; else if (s === 'Review') counts.review++; else if (s.indexOf('Blocked') === 0) counts.blocked++; else counts.same++; });
    var summary = files.length + ' file(s) read, ' + covered.length + ' companies matched, ' + unmatched.length + ' unmatched. Proposed changes: ' + counts.ready + ' ready, ' + counts.review + ' to review, ' + counts.blocked + ' blocked, ' + counts.same + ' unchanged. ' + countBy(issues) + '.';
    finishRun(run, unmatched.length ? 'Partial' : 'OK', summary);
    PropertiesService.getScriptProperties().setProperty('PENDING_BATCH', batchId);

    var result = { runId: run.id, batchId: batchId, files: files.length, matched: covered.length, unmatched: unmatched.length, counts: counts, summary: summary };
    lock.releaseLock();
    if (String(settings.autoApplyClean).toUpperCase() === 'YES' && !issues.some(function (i) { return i.severity === 'ERROR'; })) {
      result.applied = applyImport({ trigger: 'auto' });
    }
    return result;
  } catch (e) {
    finishRun(run, 'Failed', String(e && e.message || e));
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* already released */ }
  }
}

function sameValue(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9 * Math.max(1, Math.abs(a));
  return String(a) === String(b);
}

function changeText(a, b) {
  if (typeof a === 'number' && typeof b === 'number' && a !== 0) return pct(b / a - 1);
  if ((a === '' || a === null) && b !== '') return 'new';
  return '';
}

/** Apply ticked rows of the Review Queue to the Comps tab. */
function applyImport(opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another update is running. Try again in a minute.');
  var run = startRun('apply', opts.trigger);
  try {
    var sh = ss_().getSheetByName(SHEETS.REVIEW);
    if (!sh || sh.getLastRow() < 2) { finishRun(run, 'OK', 'Nothing to apply.'); return { applied: 0, message: 'Nothing pending. Import CIQ exports first.' }; }
    var model = loadModel();
    var q = sh.getRange(2, 1, sh.getLastRow() - 1, REVIEW_HEADERS.length).getValues();
    var audit = [], applied = 0, companies = {}, fileIds = {}, periods = [];
    q.forEach(function (r) {
      var field = r[16], row = r[15] - 1;
      fileIds[r[14]] = true;
      if (field === 'financialPeriod' && r[5]) periods.push(String(r[5]));
      if (r[0] !== true) return;
      if (String(r[7]).indexOf('Blocked') === 0) return;
      var col = model.map.cols[field];
      if (col === undefined) return;
      var cur = model.values[row][col];
      if (sameValue(cur, r[5])) return;
      setCell(model.sheet, row, col, r[5]);
      applied++;
      companies[r[1]] = true;
      audit.push({ action: 'Financials', company: r[1], ticker: r[2], field: r[3], oldV: cur, newV: r[5], source: { file: r[9], sheet: r[10], cell: r[11], label: r[12], raw: r[5] }, note: r[8] ? 'Flags: ' + r[8] : '' });
    });

    // Notes column: engine-written flags, only where the banker has not written their own note.
    writeFlagNotes(model);

    // Header: financials source.
    if (model.header.financialsSource && applied) {
      var tz = ss_().getSpreadsheetTimeZone();
      var label = 'CIQ export, ' + (mode(periods) || '') + ' (imported ' + Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy') + ')';
      setCell(model.sheet, model.header.financialsSource.row, model.header.financialsSource.col, label);
    }
    writeAudit(run, audit);

    // Archive the processed files.
    var folders = getFolders();
    var tz2 = ss_().getSpreadsheetTimeZone();
    var dest = folders.processed.createFolder(Utilities.formatDate(new Date(), tz2, 'yyyy-MM-dd HHmm'));
    Object.keys(fileIds).forEach(function (id) { try { DriveApp.getFileById(id).moveTo(dest); } catch (e) { /* already moved */ } });
    // Unmatched and duplicate files stay in the inbox (flagged in Exceptions) for the banker to fix.

    sh.getRange(2, 1, q.length, REVIEW_HEADERS.length).clearContent().clearFormat().clearDataValidations();
    PropertiesService.getScriptProperties().deleteProperty('PENDING_BATCH');
    SpreadsheetApp.flush();
    var summary = 'Applied ' + applied + ' changes across ' + Object.keys(companies).length + ' companies. Source files moved to Processed/' + dest.getName() + '.';
    finishRun(run, 'OK', summary);
    lock.releaseLock();
    runModelChecks({ trigger: opts.trigger, quiet: true });
    return { applied: applied, companies: Object.keys(companies).length, summary: summary };
  } catch (e) {
    finishRun(run, 'Failed', String(e && e.message || e));
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* already released */ }
  }
}

/** Discard the pending Review Queue (files stay in the inbox). */
function discardImport() {
  var sh = ss_().getSheetByName(SHEETS.REVIEW);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, REVIEW_HEADERS.length).clearContent().clearFormat().clearDataValidations();
  PropertiesService.getScriptProperties().deleteProperty('PENDING_BATCH');
  return { ok: true };
}

/** Pending queue for the web app. */
function readReviewQueue() {
  var sh = ss_().getSheetByName(SHEETS.REVIEW);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, REVIEW_HEADERS.length).getValues().map(function (r, i) {
    return { idx: i, apply: r[0] === true, company: r[1], ticker: r[2], field: r[3], current: r[4], proposed: r[5], change: r[6], status: r[7], flags: r[8], file: r[9], tab: r[10], cell: r[11], label: r[12] };
  });
}

/** Tick / untick queue rows from the web app. */
function setQueueApply(idxs, value) {
  var sh = ss_().getSheetByName(SHEETS.REVIEW);
  if (!sh) return 0;
  idxs.forEach(function (i) {
    var status = sh.getRange(i + 2, 8).getValue();
    if (String(status).indexOf('Blocked') === 0) return;
    sh.getRange(i + 2, 1).setValue(!!value);
  });
  return idxs.length;
}

/** Engine-written flags in the Notes column, prefixed "⚑" so banker notes are never overwritten. */
function writeFlagNotes(model) {
  if (model.map.cols.notes === undefined) return;
  var ex = readExceptions().filter(function (e) { return e.status !== 'Reviewed' && e.severity !== 'INFO' && e.category === 'Financials'; });
  var NOTE_TEXT = {
    PRESS_RELEASE: 'LTM from press release', PERIOD_MISMATCH: 'LTM period differs from peers', FX_TRANSLATION: 'Reports in local currency; FX at spot',
    NO_EV_BRIDGE: 'Net debt from cap structure (no CIQ EV bridge)', BROKER_FUNDING_IN_DEBT: 'Broker funding in debt', IMPLIED_EBITDA: 'Implied EBITDA (EBIT + D&A)',
    IMPLIED_SHARES: 'Implied shares differ', NEGATIVE_NCI: 'Negative NCI', GROWTH_OUTLIER: 'Growth outlier', LARGE_CHANGE: 'Large change vs prior',
    RECON_EV: 'EV does not reconcile to CIQ', RECON_MARKET_CAP: 'Mkt cap does not reconcile to CIQ'
  };
  model.layout.companies.forEach(function (co) {
    var cur = String(model.values[co.row][model.map.cols.notes] || '');
    if (cur && cur.indexOf('⚑') !== 0) return; // banker's own note: leave it
    var codes = ex.filter(function (e) { return e.ticker === co.ticker; }).map(function (e) { return NOTE_TEXT[e.key.split('|')[3]] || ''; }).filter(String);
    var uniq = codes.filter(function (c, i) { return codes.indexOf(c) === i; });
    var txt = uniq.length ? '⚑ ' + uniq.join('; ') : '';
    if (txt !== cur) setCell(model.sheet, co.row, model.map.cols.notes, txt);
  });
}
