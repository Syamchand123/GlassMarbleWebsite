/**
 * GlassMarble waitlist — Google Sheets backend (Google Apps Script)
 *
 * The website (index.html) POSTs JSON to the Web App URL with
 * Content-Type: text/plain (avoids CORS preflight). This script
 * parses it, validates it, and appends a row to the "Waitlist" tab.
 *
 * ─────────────────────────────────────────────────────────────────
 * ONE-TIME SETUP  (~3 minutes)
 * 1. Create a Google Sheet:  https://sheet.new
 * 2. In that Sheet: Extensions → Apps Script
 * 3. Delete the default function and paste this whole file in,
 *    then press Ctrl+S to save.
 * 4. Deploy → New deployment → gear icon → type: Web app
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Click Deploy and authorise (Advanced → Go to project (unsafe)).
 * 5. Copy the Web app URL, e.g.
 *      https://script.google.com/macros/s/AKfycb....../exec
 * 6. In index.html, paste it into the WAITLIST_ENDPOINT constant
 *    (it is near the top of the <script> block, currently '').
 * 7. Test on the site: submit the form → a row appears in the Sheet.
 * ─────────────────────────────────────────────────────────────────
 *
 * Sheet columns:  Timestamp | Name | Email | Source
 */

/** POST / — receives a waitlist submission. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return respond_(false, 'Server busy — please try again in a moment.');
  }

  try {
    var payload = {};
    try {
      payload = JSON.parse(e.postData.contents || '{}');
    } catch (err) {
      return respond_(false, 'Malformed request.');
    }

    var name  = String(payload.name  || '').trim();
    var email = String(payload.email || '').trim();

    if (!name) {
      return respond_(false, 'Name is required.');
    }
    if (!isValidEmail_(email)) {
      return respond_(false, 'A valid email address is required.');
    }

    var sheet = getWaitlistSheet_();

    /* Duplicate check: if this email is already on the list, say so
       instead of adding a second row. Case-insensitive. */
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][2]).toLowerCase() === email.toLowerCase()) {
        return respond_(true, 'already-registered', { count: getCount_() });
      }
    }

    sheet.appendRow([
      new Date().toISOString(),
      name,
      email,
      String(payload.source || 'website')
    ]);

    return respond_(true, 'ok', { count: getCount_() });
  } catch (err) {
    console.error('doPost failed: ' + err);
    return respond_(false, 'Something went wrong — please try again.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * GET / — health check (visiting the URL in a browser shows "ok").
 * GET /?action=count — returns the number of people on the waitlist:
 *   { "ok": true, "count": 42 }
 * The website fetches this on load to show the live join counter.
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'count') {
      return respond_(true, 'ok', { count: getCount_() });
    }
    return respond_(true, 'ok');
  } catch (err) {
    return respond_(false, 'Could not read the waitlist count.');
  }
}

/** Returns the number of waitlist signups (data rows, header excluded). */
function getCount_() {
  var values = getWaitlistSheet_().getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][2]).trim() !== '') count++;  // email column
  }
  return count;
}

/** Returns the "Waitlist" tab, creating it with headers if missing. */
function getWaitlistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Waitlist');
  if (!sheet) {
    sheet = ss.insertSheet('Waitlist');
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Source']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(4, 120);
  }
  return sheet;
}

/** Minimal email sanity check (mirrors the browser's check). */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/** Builds a JSON response. `extra` is optional and merged into the payload. */
function respond_(ok, message, extra) {
  var payload = { ok: ok, message: message };
  if (extra) {
    for (var key in extra) payload[key] = extra[key];
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
