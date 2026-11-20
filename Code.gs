/**
 * ============================================================================
 *  CLIENT CRM — private Apps Script web app over the Master List
 * ============================================================================
 *
 *  A phone-first follow-up tool for the CRM_Clients tab of "Master List".
 *  The sheet stays the single source of truth: the app reads it on load and
 *  writes straight back into it.
 *
 *  WHAT IT WRITES — AND WHAT IT NEVER TOUCHES
 *  ------------------------------------------
 *  - Any cell holding a formula is skipped on every write. Next Follow Up is a
 *    formula driven by Last Followed Up + Duration, so marking a follow-up
 *    done only sets Last Followed Up (and Duration if you change the cadence);
 *    the sheet works out the next date itself.
 *  - Rows are matched by Client ID before every write, not by row number
 *    alone, so sorting or filtering the sheet between loads cannot make the
 *    app write into the wrong client.
 *  - Follow-up notes go to a CRM_Log tab (created on first use). Nothing in
 *    CRM_Clients is ever deleted.
 *
 *  Columns are located by header text, so reordering or inserting columns in
 *  the sheet does not break anything.
 *
 *  HOW THE PAGE TALKS TO IT
 *  ------------------------
 *  The page lives on GitHub Pages and POSTs JSON here: { action, key, arg }.
 *  Every call must carry the key stored in Script Properties as CRM_KEY —
 *  the deployment is "Anyone" so GitHub Pages can reach it, which makes the
 *  key the only lock. It never appears in the repo or in a URL query string.
 *
 *  DEPLOY: standalone project, Execute as "Me", Who has access "Anyone",
 *  from the Google account that owns the Master List. See Setup.md.
 * ============================================================================
 */

var SHEET_ID   = '1y8rWTNkjUaKKJH7OP4cs-NyNwynEMGNUPP2Mg9kWgYM';

/* PPC leads: "Our Master List" → CRM tab. Rows whose PIC is one of PPC_PICS
   are copied into CRM_Clients by importPpcLeads(). Topic FAKE is skipped. */
var PPC_SHEET_ID = '1XgTqUIrBnyohNjpzW1kOsUGAL88pmC4nPYgkBP1oXYI';
var PPC_TAB      = 'CRM';
var PPC_PICS     = ['kwang', 'seah'];
var PPC_SKIP_TOPICS = ['fake'];
var CLIENT_GID = 1393914885;          // CRM_Clients — found by id, so a rename is harmless
var CLIENT_TAB = 'CRM_Clients';       // fallback if the gid ever changes
var LOG_TAB    = 'CRM_Log';
var TZ         = 'Asia/Singapore';
var OWNER      = 'Kenn';

/* Header text for each field. Matching ignores case and surrounding spaces. */
var COLS = {
  id:       'Client ID',
  name:     'Name',
  last:     'Last Followed Up',
  next:     'Next Follow Up',
  phone:    'Phone',
  address:  'Address',
  type:     'Type',
  rooms:    'Rooms/Size',
  status:   'Status',
  topic:    'Topic',
  stage:    'Stage',
  duration: 'Duration',
  remarks:  'Other Remarks',
  source:   'Source',      // added by setupColumns(): Master List / PPC
  timeline: 'Timeline'     // added by setupColumns(): how soon they plan to act
};

var SOURCES   = ['Master List', 'PPC'];
var TIMELINES = ['Within 3 months', '3 to 6 months', '6 to 12 months', 'Over 1 year'];

/* Dropdown choices used only when the sheet's own validation can't be read.
   Whatever the sheet actually contains is always merged in on top. */
var DEFAULT_OPTIONS = {
  status:   ['Active', 'Ghost'],
  topic:    ['Pure Buying', 'Pure Selling', 'Upgrade', 'Rightsize', 'Investment', 'Referral Program', 'Kaypoh'],
  stage:    ['Follow Up'],
  duration: ['Daily', '2 Days', 'Weekly', 'Bi-Weekly', 'Monthly', '2 Months'],
  type:     ['HDB', 'Condo', 'EC', 'Landed', 'Rental'],
  rooms:    ['2-Room', '3-Room', '4-Room', '5-Room', 'Executive', '3Gen'],
  source:   SOURCES,
  timeline: TIMELINES
};

var LOG_HEAD = ['Log ID', 'Timestamp', 'Client ID', 'Name', 'Action', 'Note', 'Previous Last Followed Up'];

// ---------------------------------------------------------------------------
//  JSON API (called from the GitHub Pages app)
// ---------------------------------------------------------------------------

var ACTIONS = {
  getData:        function ()  { return getData(); },
  logFollowUp:    function (a) { return logFollowUp(a); },
  undoFollowUp:   function (a) { return undoFollowUp(a); },
  saveClient:     function (a) { return saveClient(a); },
  importPpcLeads: function ()  { return importPpcLeads(); },
  setupColumns:   function ()  { return setupColumns(); }
};

function doGet() {
  return json_({ ok: true, app: 'Client CRM API' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request.' }); }

  var expected = PropertiesService.getScriptProperties().getProperty('CRM_KEY');
  if (!expected) return json_({ ok: false, error: 'CRM_KEY is not set in Script Properties.' });
  if (!body || typeof body.key !== 'string' || !safeEqual_(body.key, expected)) {
    Utilities.sleep(800);                             // slows any guessing
    return json_({ ok: false, error: 'unauthorized' });
  }

  var fn = ACTIONS[body.action];
  if (!fn) return json_({ ok: false, error: 'Unknown action.' });
  try {
    return json_({ ok: true, data: fn(body.arg) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

// ---------------------------------------------------------------------------
//  Read
// ---------------------------------------------------------------------------

/** Everything the page needs in one round trip. */
function getData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = clientSheet_(ss);
  var map = headerMap_(sh);
  var values = sh.getDataRange().getValues();

  var clients = [];
  for (var r = 1; r < values.length; r++) {
    var c = rowToClient_(values[r], map, r + 1);
    if (c.name) clients.push(c);
  }

  return {
    owner:   OWNER,
    today:   todayStr_(),
    clients: clients,
    options: readOptions_(sh, map, clients),
    log:     readLog_(ss, 2000),
    loaded:  Utilities.formatDate(new Date(), TZ, 'd MMM, HH:mm')
  };
}

// ---------------------------------------------------------------------------
//  Write
// ---------------------------------------------------------------------------

/**
 * Records a follow-up: Last Followed Up = today, optional cadence / status /
 * stage change, note into CRM_Log (and optionally appended to Other Remarks).
 * Returns the client re-read from the sheet, so the page shows the date the
 * sheet's own formula produced.
 *
 * p = { row, id, action, note, duration, status, stage, addToRemarks }
 */
function logFollowUp(p) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = clientSheet_(ss);
    var map = headerMap_(sh);
    var row = locateRow_(sh, map, p.row, p.id);

    var prevLast = cellStr_(sh, row, map.last);
    var updates = {};
    updates.last = sheetToday_(ss);
    if (p.duration) updates.duration = p.duration;
    if (p.status)   updates.status   = p.status;
    if (p.stage)    updates.stage    = p.stage;

    var note = String(p.note || '').trim();
    if (note && p.addToRemarks && map.remarks != null) {
      var old = String(sh.getRange(row, map.remarks + 1).getValue() || '').trim();
      var stamp = Utilities.formatDate(new Date(), TZ, 'dd/MM');
      updates.remarks = (old ? old + '\n' : '') + stamp + ': ' + note;
    }

    writeFields_(sh, map, row, updates);
    SpreadsheetApp.flush();

    var name = cellStr_(sh, row, map.name);
    var logId = appendLog_(ss, {
      id: p.id, name: name, action: p.action || 'Followed up', note: note, prevLast: prevLast
    });

    return { client: readRow_(sh, map, row), logEntry: lastLogEntry_(ss, logId) };
  });
}

/** Undo a follow-up logged a moment ago: restores Last Followed Up, drops the log row. */
function undoFollowUp(p) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = clientSheet_(ss);
    var map = headerMap_(sh);
    var row = locateRow_(sh, map, p.row, p.id);

    var restore = {};
    restore.last = p.prevLast ? parseDay_(ss, p.prevLast) : '';
    if (p.prevDuration) restore.duration = p.prevDuration;
    if (p.prevStatus)   restore.status   = p.prevStatus;
    if (p.prevStage)    restore.stage    = p.prevStage;
    if (p.prevRemarks != null) restore.remarks = p.prevRemarks;
    writeFields_(sh, map, row, restore);

    deleteLog_(ss, p.logId);
    SpreadsheetApp.flush();
    return { client: readRow_(sh, map, row) };
  });
}

/**
 * Creates or edits a client. p = { row?, id?, fields: {name, phone, ...} }.
 * New clients go into the first row whose Name is blank, so the table's
 * pre-filled Client IDs, formatting, dropdowns and formulas are reused.
 */
function saveClient(p) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = clientSheet_(ss);
    var map = headerMap_(sh);
    var f = p.fields || {};
    var isNew = !p.id && !p.row;

    // Partial updates (sorting a lead, moving source) send only the changed fields.
    if ((isNew || 'name' in f) && !String(f.name || '').trim()) throw new Error('Name is required.');

    var row;
    if (isNew) {
      row = firstBlankRow_(sh, map);
      var existingId = cellStr_(sh, row, map.id);
      if (!existingId && map.id != null) f.id = nextClientId_(sh, map);
      if (!f.last) f.last = todayStr_();
    } else {
      row = locateRow_(sh, map, p.row, p.id);
    }

    var updates = {};
    Object.keys(COLS).forEach(function (k) {
      if (k === 'next') return;                     // always the sheet's formula
      if (k === 'id' && !isNew) return;             // IDs are never rewritten
      if (!(k in f)) return;
      var v = f[k];
      if (k === 'last') v = v ? parseDay_(ss, v) : '';
      updates[k] = v;
    });
    writeFields_(sh, map, row, updates);
    if (isNew) ensureNextFormula_(sh, map, row);
    SpreadsheetApp.flush();

    var client = readRow_(sh, map, row);
    var action = isNew ? 'Added' : (p.logAction ? String(p.logAction).slice(0, 60) : 'Edited');
    appendLog_(ss, { id: client.id, name: client.name, action: action, note: '', prevLast: '' });
    return { client: client, logEntry: null };
  });
}

/**
 * Copies PPC leads that you or Seah are handling (PIC column) from
 * Our Master List → CRM into CRM_Clients.
 *
 * Safe to press repeatedly: a lead is skipped when its mobile (last 8 digits)
 * — or, with no mobile, its name — is already in CRM_Clients. Nothing in
 * Our Master List is changed; this only reads it.
 *
 * Mapping: Name, Phone, Last Followed Up, Duration, Topic, Stage and Current
 * Address carry straight over; Category "Ghost" becomes Status Ghost, anything
 * else Active; PIC, Interest (project) and Category go into Other Remarks so
 * they stay searchable.
 */
function importPpcLeads() {
  return withLock_(function () {
    var src = SpreadsheetApp.openById(PPC_SHEET_ID).getSheetByName(PPC_TAB);
    if (!src) throw new Error('No "' + PPC_TAB + '" tab in Our Master List.');
    var vals = src.getDataRange().getValues();
    var h = vals[0].map(function (x) { return String(x).trim().toLowerCase(); });
    function at(n) { return h.indexOf(n); }
    var iPic = at('pic'), iName = at('name'), iLast = at('last followed up'), iPhone = at('phone'),
        iCat = at('category'), iTopic = at('topic'), iStage = at('stage'), iDur = at('duration'),
        iInt = at('interest'), iRem = at('other remarks'), iAddr = at('current address');
    if (iPic < 0 || iName < 0) throw new Error('Our Master List → CRM needs PIC and Name columns.');

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = clientSheet_(ss);
    var map = headerMap_(sh);
    var opts = readOptions_(sh, map, []);

    // Who is already in CRM_Clients.
    var have = sh.getDataRange().getValues().slice(1);
    var phones = {}, names = {};
    have.forEach(function (r) {
      var nm = String(r[map.name] || '').trim().toLowerCase();
      if (!nm) return;
      names[nm] = true;
      var p = last8_(map.phone != null ? r[map.phone] : '');
      if (p) phones[p] = true;
    });

    function pick(list, v) {          // match the sheet's own spelling, e.g. KAYPOH → Kaypoh
      v = String(v || '').trim();
      for (var i = 0; i < (list || []).length; i++) if (list[i].toLowerCase() === v.toLowerCase()) return list[i];
      return v;
    }

    var added = [], skipped = 0, already = 0;
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      var name = String(row[iName] || '').trim();
      var pic = String(row[iPic] || '').trim();
      if (!name || PPC_PICS.indexOf(pic.toLowerCase()) < 0) continue;
      var topic = iTopic >= 0 ? String(row[iTopic] || '').trim() : '';
      if (PPC_SKIP_TOPICS.indexOf(topic.toLowerCase()) >= 0) { skipped++; continue; }

      var phone = iPhone >= 0 ? String(row[iPhone] || '').trim() : '';
      var p8 = last8_(phone);
      if ((p8 && phones[p8]) || (!p8 && names[name.toLowerCase()])) { already++; continue; }

      var cat = iCat >= 0 ? String(row[iCat] || '').trim() : '';
      var interest = iInt >= 0 ? String(row[iInt] || '').trim() : '';
      var note = ['PPC lead', 'PIC ' + pic, interest ? 'Interest: ' + interest : '', cat ? 'Category: ' + cat : '']
                   .filter(String).join(' · ');
      var rem = iRem >= 0 ? String(row[iRem] || '').trim() : '';

      var target = firstBlankRow_(sh, map);
      var fields = {
        name: name,
        phone: phone,
        address: iAddr >= 0 ? String(row[iAddr] || '').trim() : '',
        status: /ghost/i.test(cat) ? 'Ghost' : 'Active',
        topic: pick(opts.topic, topic),
        stage: pick(opts.stage, iStage >= 0 ? row[iStage] : '') || 'Follow Up',
        duration: pick(opts.duration, iDur >= 0 ? row[iDur] : '') || 'Monthly',
        remarks: note + (rem ? '\n' + rem : ''),
        source: 'PPC'
      };
      var lastV = iLast >= 0 ? row[iLast] : '';
      fields.last = lastV instanceof Date ? lastV : sheetToday_(ss);
      if (map.id != null && !cellStr_(sh, target, map.id)) fields.id = nextClientId_(sh, map);

      writeFields_(sh, map, target, fields);
      ensureNextFormula_(sh, map, target);
      if (p8) phones[p8] = true;
      names[name.toLowerCase()] = true;
      added.push({ row: target, name: name, pic: pic });
    }
    SpreadsheetApp.flush();

    added.forEach(function (a) {
      var c = readRow_(sh, map, a.row);
      appendLog_(ss, { id: c.id, name: c.name, action: 'Imported from PPC', note: 'PIC ' + a.pic, prevLast: '' });
    });

    return { added: added.length, already: already, skipped: skipped };
  });
}

/**
 * New rows below the table's last formula row may have no Next Follow Up
 * formula. Copy it from the first data row (relative R1C1, so it points at
 * its own row) — never overwrites a cell that already has a formula or value.
 */
function ensureNextFormula_(sh, map, row) {
  if (map.next == null || row <= 2) return;
  var cell = sh.getRange(row, map.next + 1);
  if (cell.getFormula() || cell.getValue() !== '') return;
  var model = sh.getRange(2, map.next + 1).getFormulaR1C1();
  if (model) cell.setFormulaR1C1(model);
}

function last8_(v) {
  var d = String(v == null ? '' : v).replace(/\D/g, '');
  return d.length >= 8 ? d.slice(-8) : '';
}

// ---------------------------------------------------------------------------
//  Sheet helpers
// ---------------------------------------------------------------------------

function clientSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === CLIENT_GID) return sheets[i];
  }
  var byName = ss.getSheetByName(CLIENT_TAB);
  if (byName) return byName;
  throw new Error('Cannot find the CRM_Clients tab.');
}

/** { field: zero-based column index } for every header that was found. */
function headerMap_(sh) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function (h) { return String(h).trim().toLowerCase(); });
  var map = {};
  Object.keys(COLS).forEach(function (k) {
    var i = head.indexOf(COLS[k].toLowerCase());
    if (i >= 0) map[k] = i;
  });
  if (map.name == null) throw new Error('No "Name" column in CRM_Clients.');
  return map;
}

function rowToClient_(row, map, rowNum) {
  function s(k) { return map[k] != null && row[map[k]] != null ? String(row[map[k]]).trim() : ''; }
  function d(k) { return map[k] != null ? dayStr_(row[map[k]]) : ''; }
  return {
    row: rowNum,
    id: s('id'), name: s('name'),
    last: d('last'), next: d('next'),
    phone: s('phone'), address: s('address'),
    type: s('type'), rooms: s('rooms'),
    status: s('status'), topic: s('topic'), stage: s('stage'),
    duration: s('duration'), remarks: s('remarks'),
    source: s('source') || guessSource_(s('remarks')),
    timeline: s('timeline')
  };
}

/** Rows imported from PPC start their remarks with "PPC lead"; everything else is the Master List. */
function guessSource_(remarks) {
  return /^PPC lead/i.test(String(remarks || '')) ? 'PPC' : 'Master List';
}

/**
 * One-time (safe to repeat) sheet setup for the source/timeline layout:
 *  1. Adds "Source" and "Timeline" headers in the first empty columns after
 *     the last header, with dropdowns that still allow other values.
 *  2. Fills Source for every named row that has none (PPC / Master List).
 *  3. Teaches the Next Follow Up formula a "2 Months" cadence. Only cells that
 *     already hold the Next Follow Up formula are rewritten; values are left.
 */
function setupColumns() {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = clientSheet_(ss);
    var out = { added: [], sourceFilled: 0, formulasUpdated: 0 };

    var width = sh.getLastColumn();
    var head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
    var lastHead = 0;
    head.forEach(function (h, i) { if (h) lastHead = i + 1; });
    [['Source', SOURCES], ['Timeline', TIMELINES]].forEach(function (pair) {
      var exists = head.some(function (h) { return h.toLowerCase() === pair[0].toLowerCase(); });
      if (exists) return;
      lastHead += 1;
      sh.getRange(1, lastHead).setValue(pair[0]).setFontWeight('bold');
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(pair[1], true).setAllowInvalid(true).build();
      sh.getRange(2, lastHead, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(rule);
      head[lastHead - 1] = pair[0];
      out.added.push(pair[0] + ' in column ' + lastHead);
    });
    SpreadsheetApp.flush();

    var map = headerMap_(sh);
    var n = Math.max(sh.getLastRow() - 1, 0);
    if (n && map.source != null) {
      var names = sh.getRange(2, map.name + 1, n, 1).getValues();
      var rems = sh.getRange(2, map.remarks + 1, n, 1).getValues();
      var srcR = sh.getRange(2, map.source + 1, n, 1);
      var src = srcR.getValues();
      for (var i = 0; i < n; i++) {
        if (String(names[i][0]).trim() && !String(src[i][0]).trim()) {
          src[i][0] = guessSource_(rems[i][0]);
          out.sourceFilled++;
        }
      }
      srcR.setValues(src);
    }

    if (n && map.next != null && map.last != null && map.duration != null) {
      var d = map.duration - map.next, l = map.last - map.next;
      var L = 'R[0]C[' + d + ']', C = 'R[0]C[' + l + ']';
      var f = '=IF(' + L + '="Daily",' + C + '+1,IF(' + L + '="Weekly",' + C + '+7,IF(' + L + '="2 Days",' + C + '+2,' +
              'IF(' + L + '="Bi-Weekly",' + C + '+14,IF(' + L + '="Monthly",EDATE(' + C + ',1),' +
              'IF(' + L + '="2 Months",EDATE(' + C + ',2),""))))))';
      var rng = sh.getRange(2, map.next + 1, n, 1);
      var cur = rng.getFormulas();
      for (var r = 0; r < n; r++) {
        var fa = cur[r][0];
        if (fa && /"Monthly"/.test(fa) && !/"2 Months"/.test(fa)) {
          sh.getRange(r + 2, map.next + 1).setFormulaR1C1(f);
          out.formulasUpdated++;
        }
      }
    }
    SpreadsheetApp.flush();
    return out;
  });
}

function readRow_(sh, map, row) {
  var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  return rowToClient_(vals, map, row);
}

/**
 * Finds the client's current row. Trusts the row number only if that row
 * still carries the same Client ID; otherwise searches by ID.
 */
function locateRow_(sh, map, row, id) {
  if (map.id == null || !id) {
    if (row) return row;
    throw new Error('Client not found.');
  }
  if (row && row <= sh.getLastRow() && cellStr_(sh, row, map.id) === String(id)) return row;
  var ids = sh.getRange(2, map.id + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id)) return i + 2;
  }
  throw new Error('Client ' + id + ' is no longer in the sheet. Pull to refresh.');
}

/** Writes only the given fields, skipping any cell that holds a formula. */
function writeFields_(sh, map, row, updates) {
  var width = sh.getLastColumn();
  var range = sh.getRange(row, 1, 1, width);
  var formulas = range.getFormulas()[0];
  Object.keys(updates).forEach(function (k) {
    var c = map[k];
    if (c == null) return;
    if (formulas[c]) return;                       // never overwrite a formula
    sh.getRange(row, c + 1).setValue(updates[k]);
  });
}

function firstBlankRow_(sh, map) {
  var last = Math.max(sh.getLastRow(), 1);
  if (last > 1) {
    var names = sh.getRange(2, map.name + 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === '') return i + 2;
    }
  }
  return last + 1;
}

function nextClientId_(sh, map) {
  var last = Math.max(sh.getLastRow(), 1);
  var max = 0;
  if (last > 1) {
    sh.getRange(2, map.id + 1, last - 1, 1).getValues().forEach(function (r) {
      var m = String(r[0]).match(/^CI(\d+)$/i);
      if (m) max = Math.max(max, +m[1]);
    });
  }
  return 'CI' + ('0000' + (max + 1)).slice(-4);
}

/**
 * Dropdown choices, read from the sheet's own data validation on the first
 * data row (so the app offers exactly what the sheet allows), merged with the
 * values already in use and the defaults above.
 */
function readOptions_(sh, map, clients) {
  var out = {};
  Object.keys(DEFAULT_OPTIONS).forEach(function (k) {
    var seen = {}, list = [];
    function add(v) { v = String(v || '').trim(); if (v && !seen[v.toLowerCase()]) { seen[v.toLowerCase()] = 1; list.push(v); } }

    var fromSheet = [];
    if (map[k] != null) {
      try {
        var dv = sh.getRange(2, map[k] + 1).getDataValidation();
        if (dv) {
          var type = dv.getCriteriaType();
          var vals = dv.getCriteriaValues();
          if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) fromSheet = vals[0] || [];
          if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            fromSheet = vals[0].getValues().map(function (r) { return r[0]; });
          }
        }
      } catch (e) { /* table dropdowns can refuse; fall back below */ }
    }
    fromSheet.forEach(add);
    if (!fromSheet.length) DEFAULT_OPTIONS[k].forEach(add);
    clients.forEach(function (c) { add(c[k]); });
    out[k] = list;
    out[k + 'FromSheet'] = fromSheet.length > 0;
  });
  return out;
}

// ---------------------------------------------------------------------------
//  Log tab
// ---------------------------------------------------------------------------

function logSheet_(ss) {
  var sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(6, 420);
  }
  return sh;
}

function appendLog_(ss, e) {
  var sh = logSheet_(ss);
  var id = Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, new Date(), e.id || '', e.name || '', e.action || '', e.note || '', e.prevLast || '']);
  return id;
}

function readLog_(ss, limit) {
  var sh = ss.getSheetByName(LOG_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(limit, sh.getLastRow() - 1);
  var start = sh.getLastRow() - n + 1;
  return sh.getRange(start, 1, n, LOG_HEAD.length).getValues().map(logRow_).reverse();
}

function lastLogEntry_(ss, logId) {
  var sh = ss.getSheetByName(LOG_TAB);
  var row = sh.getRange(sh.getLastRow(), 1, 1, LOG_HEAD.length).getValues()[0];
  return String(row[0]) === logId ? logRow_(row) : null;
}

function logRow_(r) {
  return {
    logId: String(r[0]),
    at: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, "yyyy-MM-dd'T'HH:mm") : String(r[1]),
    id: String(r[2]), name: String(r[3]), action: String(r[4]), note: String(r[5])
  };
}

function deleteLog_(ss, logId) {
  if (!logId) return;
  var sh = ss.getSheetByName(LOG_TAB);
  if (!sh || sh.getLastRow() < 2) return;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(logId)) { sh.deleteRow(i + 2); return; }
  }
}

// ---------------------------------------------------------------------------
//  Small helpers
// ---------------------------------------------------------------------------

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/** Today as a Date at midnight in the spreadsheet's own time zone. */
function sheetToday_(ss) { return parseDay_(ss, todayStr_()); }

function parseDay_(ss, ymd) {
  var tz = ss.getSpreadsheetTimeZone() || TZ;
  return Utilities.parseDate(String(ymd).slice(0, 10), tz, 'yyyy-MM-dd');
}

function cellStr_(sh, row, col) {
  if (col == null) return '';
  var v = sh.getRange(row, col + 1).getValue();
  return v instanceof Date ? dayStr_(v) : String(v == null ? '' : v).trim();
}

/** Sheet value → 'yyyy-MM-dd' (or '' / the raw text if it isn't a date). */
function dayStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);          // 17/09/2026
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return String(v).trim();
}
