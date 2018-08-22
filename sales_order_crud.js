var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shLookupSources = ss.getSheetByName("lookup_sources");
var shSalesOrder = ss.getSheetByName("Sales Order");

// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';
var EMAIL_STATUS = 'EMAIL_STATUS';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("lookup_sources");
  var columnNumber_w = 23

  var startRow = 2; // First row of data to process
  var numRows = 2; // Number of rows to process

  // Fetch the range of cells W2:W3
  //var dataRange = sheet.getRange(startRow, 1, numRows, 3);
  var dataRange = sheet.getRange(startRow, columnNumber_w, numRows, 3) // column W = 23

  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1]; // Second column
    var emailSent = row[2]; // Third column
    if (emailSent != EMAIL_STATUS) { // Prevents sending duplicates
      var subject = 'New Sales Order';
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 25).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
};

function createSubmittedSalesOrder() {
  SpreadsheetApp.flush();

  var currentDateTime = shLookupSources.getRange('AA2').getValue();
  var licenseNumber = shSalesOrder.getRange('retailerLicenseNumber').getValue();
  var googleDriveFolderId = shLookupSources.getRange('googleDriveFolderId').getValue();

  //folder = DriveApp.getFolderById("15wX8drVVF_dqLLmMmYUj03n5_rm_vJ9j")
  folder = DriveApp.getFolderById(googleDriveFolderId)

  var ss = SpreadsheetApp.create("SubmittedSalesOrder_" + licenseNumber + '_' + currentDateTime);

  var temp = DriveApp.getFileById(ss.getId());

  folder.addFile(temp)
  DriveApp.getRootFolder().removeFile(temp)
};

function copyfile() {
  var fileId = shLookupSources.getRange('googleSpreadsheetSalesOrder').getValue();
//  var fileId = '1Uk1ChLfXuXylj1Z8feJsWq6KQj6giU9-M38CsHjEye0'
  var file = DriveApp.getFileById(fileId);

  SpreadsheetApp.flush();

  var formattedDate = Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
  //ui.alert(formattedDate);

  var currentDateTime = shLookupSources.getRange('AA2').getValue();

  // SOURCE
  var googleDriveFolderId = shLookupSources.getRange('googleDriveFolderId').getValue();
  var source_folder = DriveApp.getFolderById(googleDriveFolderId)

  // DESTINATION
  var googleDriveDestinationFolderId = shLookupSources.getRange('googleDriveDestinationFolderId').getValue();
  var dest_folder = DriveApp.getFolderById(googleDriveDestinationFolderId);

  var licenseNumber = shSalesOrder.getRange('retailerLicenseNumber').getValue();
  var newSpreadsheetFileName = "Sales Order (License " + licenseNumber + ') Placed On ' + formattedDate;

  // Make a backup copy.
  var file2 = file.makeCopy(newSpreadsheetFileName);
  dest_folder.addFile(file2);
  //source_folder.removeFile(file2);
}

function submitSalesOrder() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Please confirm (SUBMIT)',
     'Are you sure you want to submit the Sales Order?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".

    // Create new spreadsheet in shared google drive folder
    //createSubmittedSalesOrder();

    var ss1 = SpreadsheetApp.getActiveSpreadsheet();
    var shSalesOrder1 = ss.getSheetByName("Sales Order");

    var targetSpreadsheetId = shLookupSources.getRange('googleSpreadsheetSalesOrder').getValue();

    var direction = SpreadsheetApp.Direction;
    var lastRow = shSalesOrder1.getRange("C"+(shSalesOrder1.getLastRow()+1)).getNextDataCell(direction.UP).getRow();
    if (lastRow < 25) {
      lastRow = 25
    }

    var newRange = "Sales Order!C24:E" + lastRow

    var SRange = shSalesOrder1.getRange(newRange).activate();

    //get A1 notation identifying the range
    var A1Range = SRange.getA1Notation();

    var SData = SRange.getValues();

    //var tss = SpreadsheetApp.openById('1Uk1ChLfXuXylj1Z8feJsWq6KQj6giU9-M38CsHjEye0'); // tss = target spreadsheet
    var tss = SpreadsheetApp.openById(targetSpreadsheetId); // tss = target spreadsheet
    var ts = tss.getSheetByName('Sheet1'); // ts = target sheet

    //set the target range to the values of the source data
    ts.getRange(A1Range).setValues(SData);

    // Delete blank rows
    tss.getRange('1:23').activate();
    tss.getActiveSheet().deleteRows(tss.getActiveRange().getRow(), tss.getActiveRange().getNumRows());

    // Delete blank columns
    tss.getRange('A:B').activate();
    tss.getActiveSheet().deleteColumns(tss.getActiveRange().getColumn(), tss.getActiveRange().getNumColumns());

    // Create a copy of the 'Sales Order' template and save with new file name
    copyfile();

    // Send email notification of submitted sales order
    //sendEmails();

    shSalesOrder1.getRange('C1').activate();

    ui.alert('Sales order submitted successfully.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Sales order not submitted.');
  }
};
