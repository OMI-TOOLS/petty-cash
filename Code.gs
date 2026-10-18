/**
 * PETTY CASH RECONCILIATION - GOOGLE APPS SCRIPT
 * Fixed: openById + email notifications
 */

var SHEET_ID = '1ZzHErzSa57TXE1oG9w-TPOuppziy2pCk1zJSFBojXis';
var NOTIFY_EMAILS = ['margaret@omiflo.com'];

/**
 * RUN THIS FUNCTION ONCE to authorize MailApp.
 * Select "testEmail" from the function dropdown and click Run.
 * This triggers the authorization prompt for sending email.
 */
function testEmail() {
  MailApp.sendEmail({
    to: NOTIFY_EMAILS.join(','),
    subject: 'Petty Cash: Email notifications activated',
    body: 'This confirms that petty cash email notifications are now working. You will receive an email each time a transaction is submitted to the Google Sheet.\n\nView the sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
  });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Petty Cash') || ss.insertSheet('Petty Cash');

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 11).setValues([[
        'Date', 'Time', 'Transaction Code', 'Amount (Ksh)', 'Recipient',
        'Account', 'Tx Cost (Ksh)', 'M-PESA Balance (Ksh)', 'Description',
        'Submitted By', 'Submission Time'
      ]]);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
      sheet.getRange(1, 1, 1, 11).setBackground('#43d643');
      sheet.getRange(1, 1, 1, 11).setFontColor('#0a0f0a');
      sheet.setFrozenRows(1);
    }

    var submittedBy = data.submittedBy || 'Leonard';
    var submitTime = new Date(data.submitTime || Date.now()).toLocaleString('en-GB');

    var rows = data.transactions.map(function(t) {
      return [
        t.date, t.time, t.code, t.amount, t.recipient,
        t.account || '', t.cost, t.balance,
        t.description || '', submittedBy, submitTime
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('#,##0.00');
      sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0.00');
      sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('#,##0.00');
    }

    sheet.autoResizeColumns(1, 11);

    sendNotification(rows, submittedBy, submitTime);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, rowsAdded: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotification(rows, submittedBy, submitTime) {
  try {
    var subject = 'Petty Cash: ' + rows.length + ' new entr' + (rows.length === 1 ? 'y' : 'ies') + ' by ' + submittedBy;
    var body = rows.length + ' new entr' + (rows.length === 1 ? 'y was' : 'ies were') + ' posted to the Petty Cash sheet:\n\n';
    var total = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var amount = parseFloat(r[3]) || 0;
      total += amount;
      body += 'Date: ' + r[0] + '\n';
      body += 'Time: ' + r[1] + '\n';
      body += 'Code: ' + r[2] + '\n';
      body += 'Amount: Ksh ' + amount + '\n';
      body += 'Recipient: ' + r[4] + '\n';
      if (r[5]) body += 'Account: ' + r[5] + '\n';
      if (r[8]) body += 'Description: ' + r[8] + '\n';
      body += 'Submitted by: ' + r[9] + '\n';
      body += '---\n\n';
    }
    body += 'Total: Ksh ' + total + '\n\n';
    body += 'View the full sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit\n';
    MailApp.sendEmail({
      to: NOTIFY_EMAILS.join(','),
      subject: subject,
      body: body
    });
  } catch (err) {
    Logger.log('Email notification failed: ' + err.message);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'Petty Cash Reconciliation' }))
    .setMimeType(ContentService.MimeType.JSON);
}
