var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shLookupFormFields = ss.getSheetByName("lookup_form_fields");
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

// copy data from Google Sheet A to Google Sheet B
function cloneGoogleSheet(ssA, ssB) {
  // source doc
  var sss = SpreadsheetApp.openById(ssA);

  // source sheet
  var ss = sss.getSheetByName('Source spreadsheet');

  // Get full range of data
  var SRange = ss.getDataRange();

  // get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();

  // get the data values in range
  var SData = SRange.getValues();

  // target spreadsheet
  var tss = SpreadsheetApp.openById(ssB);

  // target sheet
  var ts = tss.getSheetByName('Target Spreadsheet');

  // Clear the Google Sheet before copy
  ts.clear({contentsOnly: true});

  // set the target range to the values of the source data
  ts.getRange(A1Range).setValues(SData);
};

function testGetNameOfFilesInDrive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('get_sheet_id');
  sh.getRange('A:A').clearContent();

  var rowNumber = 1;

  //var file = DriveApp.getFileById("1pkwQ9te-EtpqC_NC3BoHzOTUoC7axZDcAfxrqMgslwg");
  //var source_folder = DriveApp.getFolderById("1ZG3J8uFTZqkih1x6TxapeFLWJeD_Y-KB")

  // Log the name of every file in the user's Drive that modified after February 28,
  // 2013 whose name contains "untitled".
  //var files = DriveApp.searchFiles('modifiedDate > "2013-02-28" and title contains "untitled"');
  var files = DriveApp.searchFiles('title contains "Catalog-Groups"');
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log(file.getName());
    sh.getRange('A' + rowNumber).setValue(file.getName());
    sh.getRange('A' + rowNumber).offset(0,1).setValue(file.getId());
    sh.getRange('A' + rowNumber).offset(0,2).setValue(file.getUrl());
    rowNumber++;
  }
};

function searchForIdBasedOnUserInput() {
  // Prompt the user for a search term
  var searchTerm = Browser.inputBox("Enter the string to search for:");

  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('get_sheet_id');
  var sheetGetCatalogGroups = ss.getSheetByName('get_catalog_groups');
  //var sheet = ss.getActiveSheet();

  // Set up the spreadsheet to display the results
  var headers = [["File Name", "File Type", "URL", "Parent Folder", "Last Updated", "ID"]];
  sheet.clear();
  sheet.getRange("A1:F1").setValues(headers);

  // Search the files in the user's Google Drive for the search term
  // See documentation for search parameters you can use
  // https://developers.google.com/apps-script/reference/drive/drive-app#searchFiles(String)
  var files = DriveApp.searchFiles("title contains '"+searchTerm.replace("'","\'")+"'");

  var urlPrefix = "https://docs.google.com/spreadsheets/d/"

  // create an array to store our data to be written to the sheet
  var output = [];
  // Loop through the results and get the file name, file type, and URL
  while (files.hasNext()) {
    var file = files.next();

    var name = file.getName();
    var type = file.getMimeType();
    var url = file.getUrl();
    var parent = file.getParents().next().getName();
    var fileDate = file.getLastUpdated();
    var id = getIdFrom(url);

    // push the file details to our output array (essentially pushing a row of data)
    output.push([name, type, url, parent, fileDate, id]);

    break;
  }

  var sss = SpreadsheetApp.openById(id); //replace with source ID
  var ss = sss.getSheetByName('Catalog Groups'); //replace with source Sheet tab name
  var range = ss.getRange('A1:E15').activate(); //assign the range you want to copy
  var data = range.getValues();

  var numRows = range.getRows();
  var numCols = range.getColumns();

  //var tss = SpreadsheetApp.openById('spreadsheet_key'); //replace with destination ID
  //var ts = tss.getSheetByName('get_sheet_id'); //replace with destination Sheet tab name
  sheetGetCatalogGroups.getRange(1, 1, numRows, numCols).setValues(data);
  //ts.getRange(ts.getLastRow()+1, 1,5,5).setValues(data); //you will need to define the size of the copied data see getRange()

  // write data to the sheet
  //sheet.getRange(2, 1, output.length, 6).setValues(output);

  //sheet.getRange('A5').clearContent();
  //sheet.getRange('A5').setFormula('=IMPORTRANGE(\"' + urlPrefix + id + '\",\"A2:C15\")');
  ui.alert("done");
}

function getIdFrom(url) {
  var id = "";
  var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
    id = (parts[6].split("=")[1]).replace("&usp","");
    return id;
  } else {
    id = parts[5].split("/");
    //Using sort to get the id as it is the longest element.
    var sortArr = id.sort(function(a,b){return b.length - a.length});
    id = sortArr[0];
    return id;
  }
}

function original_copyDataBetweenSpreadsheets() {
  var sss = SpreadsheetApp.openById('spreadsheet_key'); //replace with source ID
  var ss = sss.getSheetByName('Source'); //replace with source Sheet tab name
  var range = ss.getRange('A1:E15'); //assign the range you want to copy
  var data = range.getValues();

  var tss = SpreadsheetApp.openById('spreadsheet_key'); //replace with destination ID
  var ts = tss.getSheetByName('SavedData'); //replace with destination Sheet tab name
  ts.getRange(ts.getLastRow()+1, 1,5,5).setValues(data); //you will need to define the size of the copied data see getRange()
}

function runConvertExcelToGoogleSpreadsheet() {
  //var fileName = "Catalog-Groups.xlsx"
  var fileName = "https://drive.google.com/drive/folders/1ZG3J8uFTZqkih1x6TxapeFLWJeD_Y-KB"
  convertExceltoGoogleSpreadsheet2(fileName);
  ui.alert("DONE");
}

function convertExceltoGoogleSpreadsheet(fileName) {
  try {
    fileName = fileName || "microsoft-excel.xlsx";

    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;
    var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: folderId}],
    };
    Drive.Files.insert(resource, blob);
  } catch (f) {
    Logger.log(f.toString());
  }
}

function convertExceltoGoogleSpreadsheet2(fileName) {
  try {
    fileName = fileName || "microsoft-excel.xlsx";

    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;
    var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName().replace(/.xlsx?/, ""),
      key: fileId
    };
    Drive.Files.insert(resource, blob, {
      convert: true
    });
  } catch (f) {
    Logger.log(f.toString());
  }
}
