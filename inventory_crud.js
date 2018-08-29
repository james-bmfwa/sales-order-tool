var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shConnCatalogGroups = ss.getSheetByName("conn_catalog_groups");
var shStageInventory = ss.getSheetByName("stg_inventory");
var shSalesOrder = ss.getSheetByName('Sales Order');
var shLookupFormFields = ss.getSheetByName("lookup_form_fields");

function getNewFileIdForCatalogGroups() {
  var folderId = shLookupFormFields.getRange('srcFolderId_DataSources').getValue();
  var sourceFolder = DriveApp.getFolderById(folderId);
  var file = sourceFolder.getFilesByName('Catalog-Groups').next();
  var newFileId = file.getId();

  // Update the new file id for srcFileId_CatalogGroups
  shLookupFormFields.getRange('srcFileId_CatalogGroups').setValue(newFileId);
};

// Refresh Catalog Groups data
function refreshCatalogGroupsData() {
  var srcFileId_CatalogGroups = shLookupFormFields.getRange('srcFileId_CatalogGroups').getValue();
  var sss = SpreadsheetApp.openById(srcFileId_CatalogGroups); // source ID
  var ss = sss.getSheetByName('Catalog Groups'); // source Sheet tab name
  var range = ss.getRange('A1:H'); // assign the range you want to copy

  var numCols = range.getNumColumns();
  var numRows = range.getNumRows();
  var data = range.getValues();

  // Clear the existing Catalog Groups data
  shConnCatalogGroups.getRange('A:H').clearContent;

  // Set the new Catalog Groups data
  shConnCatalogGroups.getRange(1, 1, numRows, numCols).setValues(data); // destination Sheet tab name

  // Delete the unnecessary columns
  shConnCatalogGroups.deleteColumns(1, 3);

  SpreadsheetApp.flush();
};

function stageInventory() {
  var orderQuantityColumn = 'R'; // I
  var pricePerUnitColumn = 'S'; // H

  // Activate the stg_inventory range to be cleared
  shStageInventory.getRange('2:2').activate();
  var currentCell = shStageInventory.getCurrentCell();
  shStageInventory.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();

  // Clear the active range
  shStageInventory.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  // Get last row in shConnCatalogGroups
  var lastRow = shConnCatalogGroups.getRange('A:A').getLastRow();

  // Copy data (product id, product) from shConnCatalogGroups sheet to shStageInventory sheet.
  var productRange = "A2:B" + lastRow
  shConnCatalogGroups.getRange(productRange).copyTo(shStageInventory.getRange('A2'),SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  // Copy data (units for sale, unit price) from shConnCatalogGroups sheet to shStageInventory sheet.
  var productRange = "D2:E" + lastRow
  shConnCatalogGroups.getRange(productRange).copyTo(shStageInventory.getRange('C2'),SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  // SPLIT the existing product into separate columns of strings
  var ss = SpreadsheetApp.getActive();
  shStageInventory.getRange('E2').activate(); // stg_product is column E
  ss.getCurrentCell().setFormula('=SPLIT(JOIN("",ARRAYFORMULA(MID(B2,LEN(B2)-ROW(INDIRECT("1:"&LEN(B2)))+1,1))),"-")'); // product is column B

  // CAUTION: this may need to be updated to use 'last row' vs fill down (in case there are ever blank rows in price)
  ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Get the last row in shStageInventory
  lastRow = shStageInventory.getRange('A:A').getLastRow();

  // PRODUCT NOTE (column I, sourced from H)
  shStageInventory.getRange('I2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(TRIM(JOIN("",ARRAYFORMULA(MID(H2,LEN(H2)-ROW(INDIRECT("1:"&LEN(H2)))+1,1)))),"")');
  shStageInventory.getRange('I3:I' + lastRow).activate();
  shStageInventory.getRange('I2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false)

  // PRODUCT NAME (column J, sourced from G)
  shStageInventory.getRange('J2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(TRIM(JOIN("",ARRAYFORMULA(MID(G2,LEN(G2)-ROW(INDIRECT("1:"&LEN(G2)))+1,1)))),"")');
  shStageInventory.getRange('J3:J' + lastRow).activate();
  shStageInventory.getRange('J2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // PRODUCT (column K, sourced from H + G)
  shStageInventory.getRange('K2').activate();
  shStageInventory.getCurrentCell().setFormula('=IF(I2="",J2,JOIN(" - ",I2,J2))');
  shStageInventory.getRange('K3:K' + lastRow).activate();
  shStageInventory.getRange('K2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // STRAIN (column M, sourced from F)
  shStageInventory.getRange('M2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(TRIM(JOIN("",ARRAYFORMULA(MID(F2,LEN(F2)-ROW(INDIRECT("1:"&LEN(F2)))+1,1)))),"")');
  shStageInventory.getRange('M3:M' + lastRow).activate();
  shStageInventory.getRange('M2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // STRAIN TYPE (column L, lookup by value in column M)
  shStageInventory.getRange('L2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(OFFSET(lookup_strains!$D$1,MATCH(M2,lstStrainNames,0)-1,0),"")');
  shStageInventory.getRange('L3:L' + lastRow).activate();
  shStageInventory.getRange('L2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // SIZE (column N, sourced from E)
  shStageInventory.getRange('N2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(TRIM(JOIN("",ARRAYFORMULA(MID(E2,LEN(E2)-ROW(INDIRECT("1:"&LEN(E2)))+1,1)))),"")');
  shStageInventory.getRange('N3:N' + lastRow).activate();
  shStageInventory.getRange('N2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // PRICE (column O, sourced from D)
  shStageInventory.getRange('O2').activate();
  shStageInventory.getCurrentCell().setFormula('=D2');
  shStageInventory.getRange('O3:O' + lastRow).activate();
  shStageInventory.getRange('O2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // AVAILABLE UNITS (column P, sourced from C)
  shStageInventory.getRange('P2').activate();
  shStageInventory.getCurrentCell().setFormula('=IF(AND(C2>500,C2<1000),(roundup(sum(300, (C2*0.15)))), if(AND(C2>1000,C2<3000),(roundup(sum(200,(C2*0.1)))),if(C2>3000,(roundup(sum(100,(C2*0.1)))), C2)))');
  shStageInventory.getRange('P3:P' + lastRow).activate();
  shStageInventory.getRange('P2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // MINIMUM PRICING (column Q, sourced from conn_product_pricing column G)
  shStageInventory.getRange('Q2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(OFFSET(conn_product_pricing!$G$1,MATCH(K2&" - "&N2,conn_product_pricing!F:F,0)-1,0),"")');
  shStageInventory.getRange('Q3:Q' + lastRow).activate();
  shStageInventory.getRange('Q2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // IS EXCLUDED PRODUCT (column R, sourced from lookup_excluded_products column B
  shStageInventory.getRange('R2').activate();
  shStageInventory.getCurrentCell().setFormula('=IFERROR(OFFSET(lookup_excluded_products!$F$1,MATCH(IF(I2="",J2,I2),lookup_excluded_products!D:D,0)-1,0),"No")');
  //shStageInventory.getCurrentCell().setFormula('=IFERROR(OFFSET(lookup_excluded_products!$D$1,MATCH(I2,lookup_excluded_products!B:B,0)-1,0),"No")');
  shStageInventory.getRange('R3:R' + lastRow).activate();
  shStageInventory.getRange('R2').copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Auto resize column widths
  spreadsheet.getRange('A:R').activate();
  shStageInventory.autoResizeColumns(1, 18);

  // To speed up performance, copy/paste values for stg_inventory
  SpreadsheetApp.flush();
  shStageInventory.getRange('A2').activate();
  shStageInventory.getRange('A2:Q' + lastRow).copyTo(shStageInventory.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  shStageInventory.getRange('A1').activate();

  // Clear existing columns Order Quantity and Per Unit Price in the Sales Order sheet
  lastRow = shSalesOrder.getRange(orderQuantityColumn + ':' + orderQuantityColumn).getLastRow();
  shSalesOrder.getRange(orderQuantityColumn + '6:' + pricePerUnitColumn + lastRow).clearContent();

  // Hide the Stage Inventory sheet
  shStageInventory.hideSheet();

  SpreadsheetApp.flush();

  shSalesOrder.getRange('C12').activate();
};
