var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shConnCatalogGroups = ss.getSheetByName("conn_catalog_groups");
var shStageInventory = ss.getSheetByName("stg_inventory");
var shSalesOrder = ss.getSheetByName('Sales Order');

function addItemToOrder() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  // ORDER QUANTITY
  var orderQuantityColumn = "R";
  var orderQuantityOffsetValue = 8;

  // ORDER PRICE
  var orderPriceColumn = "S";
  var orderPriceOffsetValue = 9;

  // AVAILABLE UNITS
  var availableUnitsColumn = "P";
  var availableUnitsOffsetValue = 6;

  var result = ui.alert(
     'Please confirm',
     'Add products to the Cart Items?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".

    // Get last row in shSalesOrder column R (Order Quantity)
    var direction = SpreadsheetApp.Direction;
    var lastRow = shSalesOrder.getRange(orderQuantityColumn+(shSalesOrder.getLastRow()+1)).getNextDataCell(direction.UP).getRow();
    // Check if the last row is the header row
    if (lastRow < 6) {
      lastRow = 6;
    };

    var range = shSalesOrder.getRange('J6:J' + lastRow).activate();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();

    for (var i = 1; i <= numRows; i++) {
      for (var j = 1; j <= numCols; j++) {
        // Check if an Order Quantity (column R) was entered by the user
        if (range.getCell(i,j).offset(0,orderQuantityOffsetValue).getValue() > 0) {
          var productId = range.getCell(i,j).getValue();
          var orderQuantity = range.getCell(i,j).offset(0,orderQuantityOffsetValue).getValue();
          var orderPrice = range.getCell(i,j).offset(0,orderPriceOffsetValue).getValue();

          // Check if BACKORDER message is needed (Order Quantity is MORE THAN Available Quantity)
          if (range.getCell(i,j).offset(0,orderQuantityOffsetValue).getValue() > range.getCell(i,j).offset(0,availableUnitsOffsetValue).getValue()) {
            var currentProductName = range.getCell(i,j).offset(0,1).getValue() + ' - ' + range.getCell(i,j).offset(0,3).getValue() + ' - ' + range.getCell(i,j).offset(0,4).getValue();
            ui.alert('This is a BACKORDER ITEM: \n\n' + currentProductName);
          };

          // Insert new row (new item begins row 25)
          shSalesOrder.getRange('C25:F25').activate();
          shSalesOrder.insertRowsBefore(shSalesOrder.getActiveRange().getRow(),1);

          // Need to check if Order Quantity is more than Available Units
          var availableUnitsColumn = 'P';
          // ADD CODE HERE FOR THE CHECK

          // Add the current record to the cart items (new items begins row 25)
          shSalesOrder.getRange('C25').setValue(productId);
          shSalesOrder.getRange('D25').setValue(orderQuantity);
          shSalesOrder.getRange('E25').setValue(orderPrice);

          // Add formula for the Product Name lookup
          shSalesOrder.getRange('F25').setFormula("=OFFSET(stg_inventory!$B$1,MATCH(C25,stg_inventory!A:A,0)-1,0)");
        };
      }
    }

    // Reset the Order Quantity and Order Price columns
    shSalesOrder.getRange(orderQuantityColumn + '6:' + orderPriceColumn).clearContent();

    // Reset the default Order Price column
    updateOrderQuantityDefaultValues();

    SpreadsheetApp.flush();

    ui.alert('Products were successfully added to the Cart.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('No products were added to the Cart.');
  }
};
