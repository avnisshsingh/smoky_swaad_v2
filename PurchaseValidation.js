
/**
 * ==========================================
 * Purchase Validation
 * ==========================================
 */

function validatePurchase() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Purchase Entry");

  let errors = [];

  // Purchase Date
  if (!sheet.getRange(PURCHASE.PURCHASE_DATE).getValue()) {

    errors.push("Purchase Date is required.");

  }

  // Item Name
  if (sheet.getRange(PURCHASE.ITEM_NAME).getDisplayValue().trim() == "") {

    errors.push("Item Name is required.");

  }

  // Quantity
  const qty = Number(sheet.getRange(PURCHASE.QUANTITY).getValue());

  if (qty <= 0) {

    errors.push("Quantity must be greater than zero.");

  }

  // Unit
  if (sheet.getRange(PURCHASE.UNIT).getDisplayValue().trim() == "") {

    errors.push("Unit is required.");

  }

  // Payment Type
  if (sheet.getRange(PURCHASE.PAYMENT_TYPE).getDisplayValue().trim() == "") {

    errors.push("Payment Type is required.");

  }

  // Amount
  const amount = Number(sheet.getRange(PURCHASE.AMOUNT).getValue());

  if (amount <= 0) {

    errors.push("Amount must be greater than zero.");

  }

  return {

    valid: errors.length == 0,

    errors: errors

  };

}