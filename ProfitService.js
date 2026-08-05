
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

}





/**
 * ==========================================================
 * GET AVAILABLE PROFIT MONTHS
 * ==========================================================
 */
function getAvailableProfitMonths() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const monthMap = Object.create(null);

  // --------------------------------------------------------
  // Orders
  // Column B = Order Date
  // --------------------------------------------------------

  const orderSheet = ss.getSheetByName("Orders");

  if (orderSheet && orderSheet.getLastRow() > 1) {

    const lastRow = orderSheet.getLastRow();

    const dates = orderSheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getValues();

    for (let i = 0; i < dates.length; i++) {

      const date = dates[i][0];

      if (!(date instanceof Date)) {
        continue;
      }

      const key =
        date.getFullYear() +
        "-" +
        String(date.getMonth() + 1).padStart(2, "0");

      monthMap[key] = date;

    }

  }

  // --------------------------------------------------------
  // Purchase Register
  // Column B = Purchase Date
  // --------------------------------------------------------

  const purchaseSheet =
    ss.getSheetByName("Purchase Register");

  if (purchaseSheet && purchaseSheet.getLastRow() > 1) {

    const lastRow = purchaseSheet.getLastRow();

    const dates = purchaseSheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getValues();

    for (let i = 0; i < dates.length; i++) {

      const date = dates[i][0];

      if (!(date instanceof Date)) {
        continue;
      }

      const key =
        date.getFullYear() +
        "-" +
        String(date.getMonth() + 1).padStart(2, "0");

      monthMap[key] = date;

    }

  }

  // --------------------------------------------------------
  // Convert to dropdown values
  // --------------------------------------------------------

  const timezone =
    Session.getScriptTimeZone();

  return Object
    .keys(monthMap)
    .sort()
    .reverse()
    .map(function(key) {

      return Utilities.formatDate(
        monthMap[key],
        timezone,
        "MMMM yyyy"
      );

    });

}









/**
 * ==========================================================
 * GET MONTHLY PROFIT SUMMARY
 * ==========================================================
 */
function getMonthlyProfitSummary(selectedMonth) {

  selectedMonth = String(selectedMonth || "").trim();

  if (!selectedMonth) {

    return {
      sales: 0,
      purchases: 0,
      grossProfit: 0,
      profitMargin: null
    };

  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sales = 0;
  let purchases = 0;

  // Convert "August 2026" -> year/month
  const parts = selectedMonth.split(" ");

  const monthNames = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ];

  const targetMonth = monthNames.indexOf(parts[0]);
  const targetYear = Number(parts[1]);

  if (targetMonth === -1 || !targetYear) {

    return {
      sales: 0,
      purchases: 0,
      grossProfit: 0,
      profitMargin: null
    };

  }

  // ========================================================
  // SALES
  //
  // Orders:
  // B = Order Date
  // K = Grand Total
  // L = Status
  // ========================================================

  const orderSheet = ss.getSheetByName("Orders");

  if (orderSheet && orderSheet.getLastRow() > 1) {

    const lastRow = orderSheet.getLastRow();

    const orders = orderSheet
      .getRange(
        2,
        2,
        lastRow - 1,
        11
      )
      .getValues();

    for (let i = 0; i < orders.length; i++) {

      const row = orders[i];
      const orderDate = row[0];

      if (!(orderDate instanceof Date)) {
        continue;
      }

      if (
        orderDate.getFullYear() !== targetYear ||
        orderDate.getMonth() !== targetMonth
      ) {
        continue;
      }

      const status =
        String(row[10] || "")
          .trim()
          .toLowerCase();

      if (status === "cancelled") {
        continue;
      }

      sales += Number(row[9]) || 0;

    }

  }

  // ========================================================
  // PURCHASES
  //
  // Purchase Register:
  // B = Purchase Date
  // G = Amount
  // ========================================================

  const purchaseSheet =
    ss.getSheetByName("Purchase Register");

  if (purchaseSheet && purchaseSheet.getLastRow() > 1) {

    const lastRow = purchaseSheet.getLastRow();

    const purchaseData = purchaseSheet
      .getRange(
        2,
        2,
        lastRow - 1,
        6
      )
      .getValues();

    for (let i = 0; i < purchaseData.length; i++) {

      const row = purchaseData[i];
      const purchaseDate = row[0];

      if (!(purchaseDate instanceof Date)) {
        continue;
      }

      if (
        purchaseDate.getFullYear() !== targetYear ||
        purchaseDate.getMonth() !== targetMonth
      ) {
        continue;
      }

      purchases += Number(row[5]) || 0;

    }

  }

  // ========================================================
  // PROFIT
  // ========================================================

  const grossProfit = sales - purchases;

  const profitMargin =
    sales > 0
      ? Number(
          ((grossProfit / sales) * 100).toFixed(1)
        )
      : null;

  return {

    sales: sales,

    purchases: purchases,

    grossProfit: grossProfit,

    profitMargin: profitMargin

  };

}