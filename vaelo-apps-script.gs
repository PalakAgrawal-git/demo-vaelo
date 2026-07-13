/**
 * VAELO — Google Sheets sync backend (Google Apps Script)
 * ------------------------------------------------------------
 * Turns a Google Sheet into a tiny sync API for the Vaelo login app.
 * Every login (team, Dhruv, client staff, client head) reads/writes
 * through this, so the Sheet is the single shared source of truth.
 *
 * SETUP (one time, ~10 min) — see VAELO-SYNC-SETUP.md for screenshots-level steps:
 *   1. Create a new Google Sheet.
 *   2. Extensions ▸ Apps Script. Delete any code, paste ALL of this file, Save.
 *   3. Deploy ▸ New deployment ▸ type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Deploy, authorise, and COPY the Web app URL (ends with /exec).
 *   4. Paste that URL into SYNC_URL near the top of vaelo-login.html.
 *
 * The app stores its state as JSON under keys in a "store" tab, and also
 * mirrors the content calendar and finance into readable "Calendar" and
 * "Finance" tabs so you can view/print them like a normal sheet.
 */

var STORE_TAB = 'store';
var STAGES = ['Idea','Calendar sent','Calendar approved','Creative in progress','Internal approval','Client approval','Scheduled/Posted'];

function storeSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STORE_TAB);
  if (!sh) { sh = ss.insertSheet(STORE_TAB); sh.appendRow(['key', 'value', 'updatedAt']); }
  return sh;
}

function readAll_() {
  var sh = storeSheet_();
  var rows = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < rows.length; i++) { if (rows[i][0]) out[rows[i][0]] = rows[i][1]; }
  return out;
}

function writeKey_(key, value) {
  var sh = storeSheet_();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      sh.getRange(i + 1, 3).setValue(new Date());
      mirror_(key, value);
      return;
    }
  }
  sh.appendRow([key, value, new Date()]);
  mirror_(key, value);
}

/* Mirror app state into human-readable tabs. */
function mirror_(key, value) {
  try {
    var data = JSON.parse(value);
    if (key === 'vaelo-content-state' && data && data.items) {
      var sh = tab_('Calendar');
      sh.clear();
      sh.appendRow(['Client', 'Date', 'Format', 'Topic', 'Product', 'Stage']);
      data.items.forEach(function (it) {
        sh.appendRow([it.client, it.date, it.format, it.topic, it.product, STAGES[it.stage] || it.stage]);
      });
    }
    if (key === 'vaelo-finance-state' && data) {
      var fh = tab_('Finance');
      fh.clear();
      fh.appendRow(['Client', 'Date', 'Item', 'Category', 'Amount', 'Paid', 'Added by']);
      Object.keys(data).forEach(function (client) {
        (data[client] || []).forEach(function (x) {
          fh.appendRow([client, x.date, x.item, x.cat, x.amt, x.paid ? 'Yes' : 'No', x.addedBy]);
        });
      });
    }
  } catch (e) { /* value wasn't JSON we mirror — ignore */ }
}

function tab_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/* GET ?key=... returns {key, value}; GET with no key returns all keys. */
function doGet(e) {
  var key = e && e.parameter && e.parameter.key;
  var payload;
  if (key) { var all = readAll_(); payload = { key: key, value: all.hasOwnProperty(key) ? all[key] : null }; }
  else { payload = readAll_(); }
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

/* POST body {key, value}  (or {updates:{k:v,...}}) upserts. */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.updates) { Object.keys(body.updates).forEach(function (k) { writeKey_(k, body.updates[k]); }); }
    else if (body.key) { writeKey_(body.key, body.value); }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
