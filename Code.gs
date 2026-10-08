/**
 * Drive -> Zip -> Email mailer (Google Apps Script)
 *
 * Reads rows from the "Mailer" sheet. For every active row it zips the files in
 * the given Google Drive folder and emails the zip to the recipient(s).
 * Docs/Sheets/Slides are exported as PDF; sub-folders are included when asked.
 */

var SHEET_NAME = 'Mailer';
var MAX_ATTACHMENT_BYTES = 24 * 1024 * 1024; // Gmail limit is 25 MB; keep headroom

// Column positions (1-based) in the "Mailer" sheet.
var COL = {
  NAME: 1,        // Label used in the zip name and email
  FOLDER: 2,      // Drive folder URL or ID
  TO: 3,          // Recipient email(s), comma separated
  CC: 4,          // Optional CC
  SUBJECT: 5,     // Optional subject (default: "Documents: <Name>")
  REPEAT: 6,      // once | daily | weekly | monthly | off
  SUBFOLDERS: 7,  // yes | no  (include sub-folders)
  LAST_SENT: 8,   // written by the script
  STATUS: 9       // written by the script
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Zip Mailer')
    .addItem('Send now (all due rows)', 'runNow')
    .addItem('Send selected row only', 'sendSelectedRow')
    .addSeparator()
    .addItem('Install daily trigger (8 AM)', 'installDailyTrigger')
    .addItem('Remove triggers', 'removeTriggers')
    .addItem('Create sample sheet', 'createSampleSheet')
    .addToUi();
}

/** Entry point for the time trigger and the menu. */
function runNow() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var sent = 0, failed = 0;
  for (var row = 2; row <= last; row++) {
    if (!isDue_(sheet, row)) continue;
    if (processRow_(sheet, row)) sent++; else failed++;
  }
  log_('Run finished. Sent: ' + sent + ', failed: ' + failed);
}

function sendSelectedRow() {
  var sheet = getSheet_();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a data row first (not the header).');
    return;
  }
  processRow_(sheet, row);
}

/** Is this row due for sending right now? */
function isDue_(sheet, row) {
  var v = sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
  if (!v[COL.FOLDER - 1] || !v[COL.TO - 1]) return false;
  var repeat = String(v[COL.REPEAT - 1] || 'once').toLowerCase().trim();
  var lastSent = v[COL.LAST_SENT - 1];
  if (repeat === 'off') return false;
  if (repeat === 'once') return !lastSent;
  if (!lastSent) return true;
  var days = (new Date() - new Date(lastSent)) / 86400000;
  var needed = { daily: 0.9, weekly: 6.9, monthly: 27 }[repeat];
  return needed !== undefined && days >= needed;
}

function processRow_(sheet, row) {
  var v = sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
  var name = String(v[COL.NAME - 1] || 'Documents').trim();
  try {
    var folder = DriveApp.getFolderById(extractId_(v[COL.FOLDER - 1]));
    var includeSub = String(v[COL.SUBFOLDERS - 1]).toLowerCase().trim() === 'yes';
    var blobs = collectBlobs_(folder, includeSub, '');
    if (!blobs.length) throw new Error('Folder is empty');

    var zipName = sanitize_(name) + '_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.zip';
    var zip = Utilities.zip(blobs, zipName);
    var subject = v[COL.SUBJECT - 1] || ('Documents: ' + name);
    var options = { name: 'Zip Mailer' };
    if (v[COL.CC - 1]) options.cc = String(v[COL.CC - 1]);

    if (zip.getBytes().length <= MAX_ATTACHMENT_BYTES) {
      options.attachments = [zip];
      options.htmlBody = body_(name, blobs.length, null);
    } else {
      // Too big for email: store the zip in Drive and send a link instead.
      var file = folder.createFile(zip);
      options.htmlBody = body_(name, blobs.length, file.getUrl());
    }
    MailApp.sendEmail(String(v[COL.TO - 1]), subject,
      'Please view this email in an HTML-capable client.', options);

    sheet.getRange(row, COL.LAST_SENT).setValue(new Date());
    sheet.getRange(row, COL.STATUS).setValue('Sent (' + blobs.length + ' files)');
    return true;
  } catch (e) {
    sheet.getRange(row, COL.STATUS).setValue('Error: ' + e.message);
    log_('Row ' + row + ' failed: ' + e.message);
    return false;
  }
}

/** Collect every file in the folder as a blob (optionally recursive). */
function collectBlobs_(folder, includeSub, prefix) {
  var blobs = [];
  var seen = {};
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var blob = toBlob_(f);
    if (!blob) continue;
    var n = prefix + blob.getName();
    if (seen[n]) n = prefix + f.getId().slice(0, 6) + '_' + blob.getName(); // keep names unique
    seen[n] = true;
    blobs.push(blob.setName(n));
  }
  if (includeSub) {
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sub = subs.next();
      blobs = blobs.concat(collectBlobs_(sub, true, prefix + sub.getName() + '/'));
    }
  }
  return blobs;
}

/** Native Google files cannot be zipped directly: export them as PDF. */
function toBlob_(file) {
  var type = file.getMimeType();
  if (type.indexOf('application/vnd.google-apps.') === 0) {
    if (type === 'application/vnd.google-apps.folder' ||
        type === 'application/vnd.google-apps.shortcut') return null;
    return file.getAs(MimeType.PDF).setName(file.getName() + '.pdf');
  }
  return file.getBlob();
}

function body_(name, count, link) {
  var html = '<p>Hello,</p><p>Please find the documents for <b>' + name +
    '</b> (' + count + ' files).</p>';
  if (link) html += '<p>The zip was too large to attach: <a href="' + link +
    '">download it here</a>.</p>';
  return html + '<p>This email was sent automatically.</p>';
}

function extractId_(text) {
  var s = String(text).trim();
  var m = s.match(/(?:\/folders\/|id=)([A-Za-z0-9_-]{10,})/);
  return m ? m[1] : s;
}

function sanitize_(s) {
  return String(s).replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '_') || 'documents';
}

function getSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME +
    '" not found. Use Zip Mailer > Create sample sheet.');
  return sheet;
}

function log_(msg) { console.log(msg); }

// ---- Triggers -------------------------------------------------------------

function installDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runNow').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getActive().toast('Daily trigger installed (runs around 8 AM).');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runNow') ScriptApp.deleteTrigger(t);
  });
}

// ---- Setup helper ---------------------------------------------------------

function createSampleSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, COL.STATUS).setValues([[
    'Name', 'Drive folder (URL or ID)', 'To', 'CC', 'Subject (optional)',
    'Repeat (once/daily/weekly/monthly/off)', 'Include sub-folders (yes/no)',
    'Last sent', 'Status'
  ]]).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(2, 1, 1, 7).setValues([[
    'Sample client', 'https://drive.google.com/drive/folders/YOUR_FOLDER_ID',
    'recipient@example.com', '', '', 'weekly', 'yes'
  ]]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, COL.STATUS);
}
