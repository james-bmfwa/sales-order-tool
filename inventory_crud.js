var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shConnCatalogGroups = ss.getSheetByName("conn_catalog_groups");
var shStageInventory = ss.getSheetByName("stg_inventory");
var shSalesOrder = ss.getSheetByName('Sales Order');

function stageInventory() {
  var unitPriceColumn = 'H';
  var quantityColumn = 'I';

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

  // Auto resize column widths
  spreadsheet.getRange('A:P').activate();
  shStageInventory.autoResizeColumns(1, 16);

  shStageInventory.getRange('A1').activate();

  // Clear existing Unit Price, Quantity
  shSalesOrder.getRange(unitPriceColumn + '7:' + quantityColumn + '7').clearContent();

  SpreadsheetApp.flush();

  shSalesOrder.getRange('C1').activate();
};

function addItemToOrder() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Please confirm',
     'Add products to the Cart Items?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".

    // Get last row in shSalesOrder column Q (Order Price)
    var direction = SpreadsheetApp.Direction;
    var lastRow = shSalesOrder.getRange("Q"+(shSalesOrder.getLastRow()+1)).getNextDataCell(direction.UP).getRow();
    // Check if the last row is the header row
    if (lastRow < 6) {
      lastRow = 6;
    };

    var range = shSalesOrder.getRange('J6:J' + lastRow).activate();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();

    for (var i = 1; i <= numRows; i++) {
      for (var j = 1; j <= numCols; j++) {
        // Check if an Order Quantity was entered by the user
        if (range.getCell(i,j).offset(0,7).getValue() > 0) {
          var productId = range.getCell(i,j).getValue();
          var orderQuantity = range.getCell(i,j).offset(0,7).getValue();
          var orderPrice = range.getCell(i,j).offset(0,8).getValue();

          // Insert new row
          shSalesOrder.getRange('C25:F25').activate();
          shSalesOrder.insertRowsBefore(shSalesOrder.getActiveRange().getRow(),1);

          // Need to check if Order Quantity is more than Available Units
          var availableUnitsColumn = 'P';
          // ADD CODE HERE FOR THE CHECK

          // Add the current record to the cart items
          shSalesOrder.getRange('C25').setValue(productId);
          shSalesOrder.getRange('D25').setValue(orderQuantity);
          shSalesOrder.getRange('E25').setValue(orderPrice);

          // Add formula for the Product Name lookup
          shSalesOrder.getRange('F25').setFormula("=OFFSET(stg_inventory!$B$1,MATCH(C25,stg_inventory!A:A,0)-1,0)");
        };
      }
    }

    // Reset the Order Quantity and Order Price columns
    shSalesOrder.getRange('Q6:R').clearContent();

    SpreadsheetApp.flush();

    ui.alert('Products were successfully added to the Cart.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('No products were added to the Cart.');
  }
};
