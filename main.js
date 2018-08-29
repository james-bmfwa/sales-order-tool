/**
 * SPREADSHEET CONSTANTS
 */
var ui = SpreadsheetApp.getUi();

// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();

// SHEET
var shConnHeadset = ss.getSheetByName("conn_headset");
var shConnTopShelfData = ss.getSheetByName("conn_topshelfdata");
var shLookupSources = ss.getSheetByName("lookup_sources");
var shSalesOrder = ss.getSheetByName("Sales Order");
var shLookupFormFields = ss.getSheetByName('lookup_form_fields');

// DIRECTION
var direction = SpreadsheetApp.Direction;

// (function) MENU ITEMS
function onOpen() {
  showSidebar();

  ui.createMenu('BMF TOOLS')
  .addItem('Refresh BMF Inventory', 'refreshBmfInventory')
  .addItem('Clear Cart Items','clearCartItems')
  .addItem('Submit Sales Order', 'submitSalesOrder')
  .addToUi();
};

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'BMF Washington');
}

// (SPREADSHEET) ON EDIT
function onEdit(e) {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var filterItems1 = sh.getRange('rngAvailableItems1'); // D12
  var filterItems2 = sh.getRange('rngAvailableItems2'); // D15
  var filterItems3 = sh.getRange('rngAvailableItems3'); // D18

  var filterChanged = false;

  // get the row & column indexes of the active cell
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // check that the active cell is within the named range for FILTER 01
  if (col >= filterItems1.getColumn() && col <= filterItems1.getLastColumn() && row >= filterItems1.getRow() && row <= filterItems1.getLastRow()) {
    filterChanged = true;
  };

  // check that the active cell is within the named range for FILTER 02
  if (col >= filterItems2.getColumn() && col <= filterItems2.getLastColumn() && row >= filterItems2.getRow() && row <= filterItems2.getLastRow()) {
    filterChanged = true;
  };

  // check that the active cell is within the named range for FILTER 03
  if (col >= filterItems3.getColumn() && col <= filterItems3.getLastColumn() && row >= filterItems3.getRow() && row <= filterItems3.getLastRow()) {
    filterChanged = true;
  };

  // A filter was changed, clear existing order quantity and price, and product list
  if (filterChanged === true) {
    updateUserFilteredProductList();
  };
}

// CLEAR EXISTING ORDER QUANTITY & PRICE COLUMNS
function clearOrderQuantityOrderPrice() {
  shSalesOrder.getRange('R6:U').clearContent();
  SpreadsheetApp.flush();
}

// UPDATE ORDER QUANTITY DEFAULT VALUES
function updateOrderQuantityDefaultValues() {
  // Automatically pre-fill the Order Price based on the BMF Price

  // Get the last row in shSalesOrder Column J
  var direction = SpreadsheetApp.Direction;
  var lastRow = shSalesOrder.getRange("J"+(shSalesOrder.getLastRow()+1)).getNextDataCell(direction.UP).getRow();

  shSalesOrder.getRange('O6:O' + lastRow).copyTo(shSalesOrder.getRange('S6:S' + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  shSalesOrder.getRange('R6').activate(); // Order quantity column

  SpreadsheetApp.flush();
}

// USER FILTERED PRODUCT LIST
function updateUserFilteredProductList() {
  SpreadsheetApp.flush();

  var formulaStringStart = "=IFERROR(SORT(UNIQUE(QUERY(stg_inventory!A2:Q,\"select A,K,L,M,N,O,P,Q where \""
  var formulaString = ''

  if (shLookupFormFields.getRange('G4').getValue() != "empty") {
    formulaString = formulaStringStart + " & lookup_form_fields!F2 & \" = '\" & " +
      "lookup_form_fields!G2 & \"' and \" & lookup_form_fields!F3 & \" = '\" & lookup_form_fields!G3 & \"' and \" & lookup_form_fields!F4 & \" = '\" & lookup_form_fields!G4 & \"'\")),7,FALSE),\"\")";
  }else if (shLookupFormFields.getRange('G3').getValue() != "empty"){
    formulaString = formulaStringStart + " & lookup_form_fields!F2 & \" = '\" & lookup_form_fields!G2 & \"' and \" & lookup_form_fields!F3 & \" = '\" & lookup_form_fields!G3 & \"'\")),7,FALSE),\"\")";
  }else{
    formulaString = formulaStringStart + " & lookup_form_fields!F2 & \" = '\" & lookup_form_fields!G2 & \"'\")),7,FALSE),\"\")";
  }

  // Sales Order Product Search Results
  shSalesOrder.getRange('searchResultFirstItem').setFormula(formulaString)

  shSalesOrder.autoResizeColumns(10,12); // start with column J = 10

  SpreadsheetApp.flush();
}

// UPDATE NAMED RANGE - MME NAMES
function updateNamedRange_MmeNames(sh,suffix) {
  // Clear existing selectedRetailer
  var selectedRetailer = "selectedRetailer_" + suffix;
  var columnLetter = "";

  if (suffix === "SO") {
    columnLetter = "H"
  }else{
    columnLetter = "J"
  }

  // Clear the existing "RETAILER"
  sh.getRange(selectedRetailer).clearContent();

  // Set the new range
  //var direction = SpreadsheetApp.Direction;
  var lastRow = shLookupSources.getRange(columnLetter+(shLookupSources.getLastRow()+1)).getNextDataCell(direction.UP).getRow();
  var newRange = "lookup_sources!" + columnLetter + "2:" + columnLetter + lastRow

  // set the new range
  var listName = "lstMmeNames_" + suffix;
  SpreadsheetApp.getActive().setNamedRange(listName, SpreadsheetApp.getActive().getRange(newRange));
}

function refreshBmfInventory() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to refresh the BMF Inventory?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".

    // Get the new file id for the Catalog Groups file
    getNewFileIdForCatalogGroups();

    // Refresh the Catalog Groups data from the s2solutions exported file
    refreshCatalogGroupsData();

    // Refresh the BMF inventory
    stageInventory();

    ui.alert('BMF Inventory was successfully refreshed.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('BMF Inventory was NOT refreshed.');
  }
}

function clearCartItems() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to clear the Cart Items?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".

    // Create new spreadsheet in shared google drive folder
    var ss1 = SpreadsheetApp.getActiveSpreadsheet();
    var shSalesOrder1 = ss.getSheetByName("Sales Order");

    // Get the last row in shStageInventory
    var lastRow = shSalesOrder1.getRange('C:C').getLastRow();
    //  Check to make sure the last row is not the header row (Row 11)
    if (lastRow < 25) {
      lastRow = 25
    }

    // Clear the Order Details
    shSalesOrder1.getRange('C25:H' + lastRow).clearContent();

    SpreadsheetApp.flush();

    ui.alert('The previous Cart Items were successfully cleared.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('The existing Cart Items were not changed.');
  }
}

function showSidebar() {
  // Log the email address of the person running the script.
  //Logger.log(Session.getActiveUser().getEmail());

  var html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('BMF Washington')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);

  ui.alert("Hi " + getOwnName() + "! Watch for updated BMF sales news (shown to the right >>).");
}

/**
 * Get current user's name, by accessing their contacts.
 *
 * @returns {String} First name (GivenName) if available,
 *                   else FullName, or login ID (userName)
 *                   if record not found in contacts.
 */
function getOwnName(){
  var email = Session.getEffectiveUser().getEmail();
  var self = ContactsApp.getContact(email);

  // If user has themselves in their contacts, return their name
  if (self) {
    // Prefer given name, if that's available
    var name = self.getGivenName();
    // But we will settle for the full name
    if (!name) name = self.getFullName();
    return name;
  }
  // If they don't have themselves in Contacts, return the bald userName.
  else {
    var userName = Session.getEffectiveUser().getUsername();
    return userName;
  }
}
