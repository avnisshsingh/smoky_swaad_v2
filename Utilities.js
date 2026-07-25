/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * UTILITIES
 * ==========================================================
 */


/**
 * ==========================================================
 * Get Active Spreadsheet
 * ==========================================================
 */
function getSS() {

  return SpreadsheetApp.getActiveSpreadsheet();

}


/**
 * ==========================================================
 * Get Sheet By Name
 * ==========================================================
 */
function getSheet(sheetName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Sheet not found: " + sheetName);
  }

  return sheet;
}


/**
 * ==========================================================
 * Get Next Empty Row
 * ==========================================================
 */
function getNextRow(sheet) {

  return sheet.getLastRow() + 1;

}


/**
 * ==========================================================
 * Generate Next ID
 * Example:
 * SS00001
 * CUS00001
 * ==========================================================
 */
function generateNextId(sheet, prefix) {

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {

    return prefix + "00001";

  }

  const lastId = String(
    sheet.getRange(lastRow, 1).getValue()
  );

  const lastNumber = Number(
    lastId.replace(prefix, "")
  );

  const nextNumber = lastNumber + 1;

  return prefix + String(nextNumber).padStart(5, "0");

}


/**
 * ==========================================================
 * Success Toast
 * ==========================================================
 */
function showSuccess(message) {

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(message, "Smoky Swaad", 3);

}


/**
 * ==========================================================
 * Error Alert
 * ==========================================================
 */
function showError(message) {

  SpreadsheetApp
    .getUi()
    .alert(message);

}


/**
 * ==========================================================
 * Read Complete POS Form
 * ==========================================================
 */
function getPOSFormData() {

  const pos = getSheet(SHEETS.POS);

  const form =
    pos.getRange("B2:B30").getValues();

  return {

    customerName: form[0][0],

    mobile: form[1][0],

    orderDate: form[2][0],

    deliveryArea: form[3][0],

    houseAddress: form[4][0],

    deliverySlot: form[5][0],

    orderType: form[6][0],

    paymentMode: form[7][0],

    paymentStatus: form[8][0],

    specialInstructions: form[9][0],

    subtotal: form[25][0],

    discount: form[26][0],

    deliveryCharge: form[27][0],

    grandTotal: form[28][0]

  };

}


/**
 * ==========================================================
 * Read All Order Items
 * ==========================================================
 */
function getOrderItems() {

  const pos = getSheet(SHEETS.POS);

  return pos.getRange(

    POS.ITEM_START_ROW,

    POS.ITEM_COL,

    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,

    4

  ).getValues();

}



/**
 * ==========================================================
 * Get Invoice Folder
 * ==========================================================
 *
 * Folder Structure
 *
 * Smoky Swaad ERP
 *      └── Invoices
 *              └── YYYY
 *                      └── MM-Month
 *
 * Example:
 *
 * Smoky Swaad ERP
 *      └── Invoices
 *              └── 2026
 *                      └── 07-July
 *
 * ==========================================================
 */
function getInvoiceFolder() {

  const ROOT_FOLDER = "Smoky Swaad ERP";

  const INVOICE_FOLDER = "Invoices";

  const today = new Date();

  const year = Utilities.formatDate(
    today,
    Session.getScriptTimeZone(),
    "yyyy"
  );

  const month = Utilities.formatDate(
    today,
    Session.getScriptTimeZone(),
    "MM-MMMM"
  );

  // ------------------------------------
  // Root Folder
  // ------------------------------------

  const rootFolder =
    getOrCreateFolder_(ROOT_FOLDER);

  // ------------------------------------
  // Invoice Folder
  // ------------------------------------

  const invoiceFolder =
    getOrCreateFolder_(
      INVOICE_FOLDER,
      rootFolder
    );

  // ------------------------------------
  // Year Folder
  // ------------------------------------

  const yearFolder =
    getOrCreateFolder_(
      year,
      invoiceFolder
    );

  // ------------------------------------
  // Month Folder
  // ------------------------------------

  const monthFolder =
    getOrCreateFolder_(
      month,
      yearFolder
    );

  return monthFolder;

}



/**
 * ==========================================================
 * Get Or Create Folder
 * ==========================================================
 */
function getOrCreateFolder_(folderName, parentFolder) {

  let folders;

  if (parentFolder) {

    folders =
      parentFolder.getFoldersByName(folderName);

  } else {

    folders =
      DriveApp.getFoldersByName(folderName);

  }

  if (folders.hasNext()) {

    return folders.next();

  }

  if (parentFolder) {

    return parentFolder.createFolder(folderName);

  }

  return DriveApp.createFolder(folderName);

}



/**
 * ==========================================================
 * TEST
 * ==========================================================
 */
function testInvoiceFolder(){

  const folder = getInvoiceFolder();

  Logger.log(folder.getName());

  Logger.log(folder.getId());

}