/*
RESOURCES: http://stackoverflow.com/questions/33499410/send-reminder-emails-based-on-date
 */

function sendEmails() {
  var today = new Date().toLocaleDateString();  // Today's date, without time
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 999;   // Number of rows to process
  // Fetch the range of cells A2:B999
  var dataRange = sheet.getRange(startRow, 1, numRows, 999)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row          = data[i];
    var emailAddress = row[0];  // First column
    var subject      = row[1];       // Second column
    var message      = row[2];       // Third column
    var emailSent    = row[3];     // Date Email was sent
    var status       = row[4]; // Status
    var x            = row[5].toString();  // Date specified in cell F
    var reminderDate = new Date(x).toLocaleDateString();
    
    if (reminderDate < today)
      continue;
    
    if (status == "Open"){
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 4).setValue('SENT '+today);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}