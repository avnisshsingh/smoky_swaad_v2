
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

    const values = sheet
        .getRange("X2:X1000")
        .getValues();

    let nextRow = null;

    for (let i = 0; i < values.length; i++) {

        const value = String(values[i][0]).trim();

        if (
            value &&
            value.toLowerCase() === itemName.toLowerCase()
        ) {

            return {
                success: true,
                alreadyExists: true,
                itemName: value
            };

        }

        if (!value && nextRow === null) {
            nextRow = i + 2;
        }

    }

    if (nextRow === null) {
        nextRow = values.length + 2;
    }

    sheet
        .getRange(nextRow, 24)
        .setValue(itemName);

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

        const purchaseSheet =
            getSheet(SHEETS.PURCHASE_REGISTER);

        const purchaseId =
            generateNextId(purchaseSheet, "PUR");

        savePurchaseEntry(
            purchaseSheet,
            purchaseId,
            purchaseData
        );

        return {
            success: true,
            purchaseId,
            message: "Purchase Saved Successfully"
        };

    } catch (error) {

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

purchase.unit = purchase.unit || "Gram";

purchase.paymentType =
    purchase.paymentType || "PhonePe";

purchase.supplier =
    purchase.supplier || "Local Market";

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

    const p = purchaseData.purchase;

    const nextRow = purchaseSheet.getLastRow() + 1;

    purchaseSheet
        .getRange(nextRow, 1, 1, 10)
        .setValues([[
            purchaseId,
            new Date(p.purchaseDate),
            p.itemName,
            Number(p.quantity),
            p.unit,
            p.paymentType,
            Number(p.amount),
            p.supplier,
            p.remarks || "",
            new Date()
        ]]);

}








function getPurchaseItems() {
   return getCachedData("CACHE_PURCHASE_ITEMS", function() {
      const sheet = getSheet(SHEETS.SETTINGS);

      return sheet
         .getRange("X2:X1000")
         .getValues()
         .flat()
         .filter(item => item && String(item).trim() !== "")
         .map(item => ({
            itemName: String(item).trim(),
            searchName: String(item).trim().toLowerCase()
         }));
   });
}