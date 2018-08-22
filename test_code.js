var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shLookupSources = ss.getSheetByName("lookup_sources");
var shSalesOrder = ss.getSheetByName("Sales Order");

function examplePrompt() {
  // Display a dialog box with a message, input field, and "Yes" and "No" buttons. The user can
  // also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('May I know your name?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}

function CopyDataToNewFile() {
  var sss = SpreadsheetApp.openById('0AjN7uZG....'); // sss = source spreadsheet
  var ss = sss.getSheetByName('Monthly'); // ss = source sheet
  //Get full range of data
  var SRange = ss.getDataRange();
  //get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();
  //get the data values in range
  var SData = SRange.getValues();

  var tss = SpreadsheetApp.openById('8AjN7u....'); // tss = target spreadsheet
  var ts = tss.getSheetByName('RAWData'); // ts = target sheet
  //set the target range to the values of the source data
  ts.getRange(A1Range).setValues(SData);
}

function myTestName() {
  //  ui.alert("Functionality Disabled")

  // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  //var uiD = DocumentApp.getUi();
  var uiD = SpreadsheetApp.getUi();
  var response = uiD.prompt('Getting to know you', 'May I know your name?', uiD.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == uiD.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == uiD.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function testFileNameDate() {
  var formattedDate = Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
  ui.alert(formattedDate);

  //var currentDateTime = shLookupSources.getRange('AA2').getValue();
  //ui.alert(currentDateTime);
}

function getSheetIdtest(){
  var id  = SpreadsheetApp.getActiveSheet().getSheetId();// get the actual id
  Logger.log(id);// log
  var sheets = SpreadsheetApp.getActive().getSheets();
  for(var n in sheets){ // iterate all sheets and compare ids
    if(sheets[n].getSheetId()==id){break}
  }
  Logger.log('tab index = '+n);// your result, test below just to confirm the value
  var currentSheetName = SpreadsheetApp.getActive().getSheets()[n].getName();
  Logger.log('current Sheet Name = '+currentSheetName);
}

function copyfile() {
  var file = DriveApp.getFileById("1pkwQ9te-EtpqC_NC3BoHzOTUoC7axZDcAfxrqMgslwg");
  var source_folder = DriveApp.getFolderById("0B8_ub-Gf21e-fkxjSUwtczJGb3picl9LUVVPbnV6Vy1aRFRWc21IVjRkRjBPTV9xMWJLRFU")
  var dest_folder = DriveApp.getFolderById("0B8_ub-Gf21e-flJ4VmxvaWxmM2NpZHFyWWxRejE5Y09CRWdIZDhDQzBmU2JnZnhyMTU2ZHM")
  // Make a backup copy.
  var file2 = file.makeCopy('BACKUP ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.' + file.getName());
  dest_folder.addFile(file2);
  source_folder.removeFile(file2);
};

function testAutoFitColumn() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O:S').activate();
  spreadsheet.getActiveSheet().autoResizeColumns(15, 5);
};

function testInsertCartItemRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('25:25').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('C25').activate();
};

function testDeleteColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:B').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('A1').activate();
};
