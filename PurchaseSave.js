
/**
 * ==========================================
 * Purchase Save Module
 * ==========================================
 */

function savePurchase() {

  const validation = validatePurchase();

  if (!validation.valid) {
    SpreadsheetApp.getUi().alert(
      validation.errors.join("\n")
    );
    return;
  }

  const purchaseID = generatePurchaseID();

  savePurchaseEntry(purchaseID);

  clearPurchaseEntry();   // <-- This line is important

  SpreadsheetApp.getUi().alert(
    "Purchase Saved Successfully.\n\nPurchase ID : " + purchaseID
  );

}


/**
 * Generate Purchase ID
 */
function generatePurchaseID() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const config = ss.getSheetByName(SHEETS.CONFIG);

  let lastNumber = Number(config.getRange("B2").getValue());

  lastNumber++;

  config.getRange("B2").setValue(lastNumber);

  return "PUR" + lastNumber.toString().padStart(5, "0");

}

/**
 * Save Purchase Entry
 */
function savePurchaseEntry(purchaseID) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const purchaseEntry = ss.getSheetByName("Purchase Entry");

  const purchaseRegister = ss.getSheetByName("Purchase Register");

  const nextRow = purchaseRegister.getLastRow() + 1;

  const rowData = [[

    purchaseID,

    purchaseEntry.getRange(PURCHASE.PURCHASE_DATE).getValue(),

    purchaseEntry.getRange(PURCHASE.ITEM_NAME).getValue(),

    purchaseEntry.getRange(PURCHASE.QUANTITY).getValue(),

    purchaseEntry.getRange(PURCHASE.UNIT).getValue(),

    purchaseEntry.getRange(PURCHASE.PAYMENT_TYPE).getValue(),

    purchaseEntry.getRange(PURCHASE.AMOUNT).getValue(),

    purchaseEntry.getRange(PURCHASE.REMARKS).getValue(),

    purchaseEntry.getRange(PURCHASE.SUPPLIER).getValue()

  ]];

  purchaseRegister
    .getRange(nextRow, 1, 1, rowData[0].length)
    .setValues(rowData);

}

/**
 * Clear Purchase Entry Screen
 */
function clearPurchaseEntry() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Purchase Entry");

  sheet.getRange(PURCHASE.PURCHASE_DATE).clearContent();

  sheet.getRange(PURCHASE.ITEM_NAME).clearContent();

  sheet.getRange(PURCHASE.QUANTITY).clearContent();

  sheet.getRange(PURCHASE.UNIT).clearContent();

  sheet.getRange(PURCHASE.PAYMENT_TYPE).clearContent();

  sheet.getRange(PURCHASE.AMOUNT).clearContent();

  sheet.getRange(PURCHASE.SUPPLIER).clearContent();

  sheet.getRange(PURCHASE.REMARKS).clearContent();

}