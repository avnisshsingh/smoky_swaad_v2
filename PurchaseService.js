/**
 * ==========================================
 * Load Purchase Settings
 * ==========================================
 */
/**
 * ==========================================
 * Load Purchase Settings
 * ==========================================
 */
function loadPurchaseSettings() {

  const sheet = getSheet(SHEETS.SETTINGS);

  return {

    // Units (Column S)
    units: sheet
      .getRange("S2:S100")
      .getValues()
      .flat()
      .filter(String),

    // Payment Types (Column D)
    paymentTypes: sheet
      .getRange("D2:D100")
      .getValues()
      .flat()
      .filter(String),

    // Suppliers (Column Y)
    suppliers: sheet
      .getRange("Y2:Y100")
      .getValues()
      .flat()
      .filter(String)

  };

}


/**
 * ==========================================
 * Search Purchase Items
 * ==========================================
 */
function searchPurchaseItems(keyword) {

  const sheet = getSheet(SHEETS.SETTINGS);

  const items = sheet
      .getRange("X2:X200")
      .getValues()
      .flat()
      .filter(String);

  if (!keyword) return [];

  keyword = keyword.toLowerCase();

  return items
      .filter(item =>
          item.toLowerCase().includes(keyword)
      )
      .map(item => ({
          itemName: item
      }))
      .slice(0, 10);

}


/**
 * ==========================================
 * Add Purchase Item
 * ==========================================
 */
function addPurchaseItem(itemName) {

  itemName = String(itemName).trim();

  if (!itemName) {

    throw new Error("Item Name is required.");

  }

  const sheet = getSheet(SHEETS.SETTINGS);

  const items = sheet
      .getRange("X2:X1000")
      .getValues()
      .flat();

  // Duplicate Check
  for (let i = 0; i < items.length; i++) {

    if (
      String(items[i]).trim().toLowerCase() ===
      itemName.toLowerCase()
    ) {

      return {

        success: true,

        alreadyExists: true,

        itemName: items[i]

      };

    }

  }

  // Find first empty row
  let nextRow = 2;

  while (sheet.getRange(nextRow, 24).getValue() !== "") {

    nextRow++;

  }

  sheet.getRange(nextRow, 24).setValue(itemName);

  return {

    success: true,

    alreadyExists: false,

    itemName: itemName

  };

}





/**
 * ==========================================
 * Save Purchase From Web
 * ==========================================
 */
function savePurchaseFromWeb(purchaseData) {

  try {

    validatePurchaseData(purchaseData);

const db = {

    purchases: getSheet(SHEETS.PURCHASE_REGISTER)

};

const purchaseId = generateNextId(
    db.purchases,
    "PUR"
);

// Save Purchase
savePurchaseEntry(
    db.purchases,
    purchaseId,
    purchaseData
);

return {

    success: true,

    purchaseId: purchaseId,

    message: "Purchase Saved Successfully"

};

  }

  catch (error) {

    return {

      success: false,

      message: error.message

    };

  }

}


/**
 * ==========================================
 * Validate Purchase Data
 * ==========================================
 */
function validatePurchaseData(purchaseData) {

  const purchase = purchaseData.purchase;

  if (!purchase.purchaseDate)
    throw new Error("Purchase Date is required.");

  if (!purchase.itemName)
    throw new Error("Item Name is required.");

  if (!purchase.quantity || purchase.quantity <= 0)
    throw new Error("Quantity should be greater than zero.");

  if (!purchase.unit)
    throw new Error("Unit is required.");

  if (!purchase.paymentType)
    throw new Error("Payment Type is required.");

  if (!purchase.amount || purchase.amount <= 0)
    throw new Error("Amount should be greater than zero.");

  return true;

}



/**
 * ==========================================
 * Save Purchase Entry
 * ==========================================
 */
function savePurchaseEntry(
    purchaseSheet,
    purchaseId,
    purchaseData
) {

    const purchase = purchaseData.purchase;

    const row = [[

        purchaseId,

        Utilities.formatDate(
            new Date(purchase.purchaseDate),
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
        ),

        purchase.itemName,

        purchase.quantity,

        purchase.unit,

        purchase.paymentType,

        purchase.amount,

        purchase.supplier,

        purchase.remarks,

        new Date()

    ]];

    purchaseSheet
        .getRange(
            purchaseSheet.getLastRow() + 1,
            1,
            1,
            row[0].length
        )
        .setValues(row);

}