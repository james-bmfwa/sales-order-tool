function testCopyPasteFormula() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('R7:R19').activate();
  spreadsheet.getRange('R6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('R6').activate();
};
