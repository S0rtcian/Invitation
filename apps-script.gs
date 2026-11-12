/**
 * RSVP collector for Asher's 1st Birthday invitation.
 * Saves every RSVP to a private Google Sheet AND emails it to the host.
 *
 * FIRST-TIME SETUP (about 5 minutes):
 * 1. Create a new Google Sheet (sheets.new). Name it "Asher RSVPs".
 * 2. In the Sheet: Extensions > Apps Script. Delete the sample code and paste ALL of this file.
 * 3. Run the function  authorizeEmail  once (pick it in the dropdown at the top, click Run).
 *    Allow the permissions (Advanced > Go to project). A test email lands in your inbox.
 * 4. Click Deploy > New deployment > gear icon > Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy.
 * 5. Copy the Web app URL (it ends with /exec).
 * 6. Open index.html, find  const SCRIPT_URL = "";  and paste the URL between the quotes.
 *
 * ALREADY DEPLOYED BEFORE? You only need to update it:
 * 1. Paste this whole file over the old code and save.
 * 2. Run  authorizeEmail  once and allow the new permission (sending email).
 * 3. Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.
 *    The URL stays the same, so index.html does not need to change.
 *
 * Only YOU can open the Sheet, so only you can see the guest list.
 * Guests can only send an RSVP. They can never read what's inside.
 * Your email address lives here in the script, never in the public page.
 * To get an Excel file: in the Sheet, File > Download > Microsoft Excel (.xlsx).
 */

var SHEET_NAME = 'RSVPs';
var NOTIFY_EMAIL = 'micandremenorgawat@gmail.com';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['Submitted', 'Name', 'Attending', 'Guests', 'Message']);
      sh.setFrozenRows(1);
      sh.getRange('A1:E1').setFontWeight('bold');
    }
    var p = e.parameter || {};
    var guests = Number(p.guests) || 0;
    sh.appendRow([
      new Date(),
      clean_(p.name, 80),
      clean_(p.attending, 3),
      guests,
      clean_(p.message, 300)
    ]);
    notify_(p, guests);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Emails the host. If email ever fails, the RSVP is still safe in the Sheet.
function notify_(p, guests) {
  try {
    var name = oneLine_(p.name, 80) || 'A guest';
    var yes = String(p.attending) === 'Yes';
    var when = Utilities.formatDate(new Date(), 'Asia/Manila', 'MMM d, yyyy h:mm a');
    var lines = [
      'Name: ' + name,
      'Attending: ' + (yes ? 'Yes' : 'No')
    ];
    if (yes) lines.push('Guests: ' + guests);
    var msg = String(p.message || '').substring(0, 300).trim();
    if (msg) lines.push('Message: ' + msg);
    lines.push('', 'Received: ' + when + ' (Philippine time)');
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'Asher RSVP: ' + name + (yes ? ' is coming' : " can't make it"),
      body: lines.join('\n'),
      name: "Asher's Invitation"
    });
  } catch (err) {
    console.error('Email failed: ' + err);
  }
}

// Run this once from the editor to give the script permission to send email.
function authorizeEmail() {
  MailApp.sendEmail(NOTIFY_EMAIL, 'Asher RSVP test', 'If you can read this, RSVP emails are working.');
}

// Stops spreadsheet formulas from being injected (=, +, -, @ at the start) and trims length.
function clean_(v, max) {
  v = String(v || '').substring(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

// One line only, so a guest can't sneak extra lines into the email subject.
function oneLine_(v, max) {
  return String(v || '').replace(/[\r\n]+/g, ' ').substring(0, max).trim();
}

// Opening the Web app URL in a browser just shows this message (no guest data is exposed).
function doGet() {
  return ContentService.createTextOutput('RSVP collector is running.');
}
