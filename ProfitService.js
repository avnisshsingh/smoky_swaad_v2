/**
 * ==========================================================
 * MONTHLY PROFIT SERVICE
 *
 * Sales      = Orders Grand Total
 * Purchases  = Purchase Register Amount
 * Profit     = Sales - Purchases
 *
 * ==========================================================
 */


/**
 * Format date for display.
 *
 * Example:
 * August 2026
 */
function formatMonthYear(date) {

  if (!(date instanceof Date)) {
    return "";
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  );

}


/**
 * Internal sortable month key.
 *
 * Example:
 * 2026-08
 */
function getProfitMonthKey(date) {

  if (!(date instanceof Date)) {
    return "";
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );

}


/**
 * ==========================================================
 * GET AVAILABLE PROFIT MONTHS
 * ==========================================================
 */
function getAvailableProfitMonths() {

  const ss =
    SpreadsheetApp.getActive();

  const monthMap =
    Object.create(null);


  // --------------------------------------------------------
  // Orders
  // --------------------------------------------------------

  const orderSheet =
    ss.getSheetByName("Orders");

  if (
    orderSheet &&
    orderSheet.getLastRow() > 1
  ) {

    const lastRow =
      orderSheet.getLastRow();

    // Only read Order Date column B.
    const dates =
      orderSheet
        .getRange(
          2,
          2,
          lastRow - 1,
          1
        )
        .getValues();

    for (
      let i = 0;
      i < dates.length;
      i++
    ) {

      const date =
        dates[i][0];

      if (!(date instanceof Date)) {
        continue;
      }

      const key =
        getProfitMonthKey(date);

      monthMap[key] =
        formatMonthYear(date);

    }

  }


  // --------------------------------------------------------
  // Purchase Register
  // --------------------------------------------------------

  const purchaseSheet =
    ss.getSheetByName(
      "Purchase Register"
    );

  if (
    purchaseSheet &&
    purchaseSheet.getLastRow() > 1
  ) {

    const lastRow =
      purchaseSheet.getLastRow();

    // Only read Purchase Date column B.
    const dates =
      purchaseSheet
        .getRange(
          2,
          2,
          lastRow - 1,
          1
        )
        .getValues();

    for (
      let i = 0;
      i < dates.length;
      i++
    ) {

      const date =
        dates[i][0];

      if (!(date instanceof Date)) {
        continue;
      }

      const key =
        getProfitMonthKey(date);

      monthMap[key] =
        formatMonthYear(date);

    }

  }


  // yyyy-MM sorts correctly as a string.
  return Object
    .keys(monthMap)
    .sort(function(a, b) {

      return b.localeCompare(a);

    })
    .map(function(key) {

      return monthMap[key];

    });

}


/**
 * ==========================================================
 * GET MONTHLY PROFIT SUMMARY
 * ==========================================================
 */
function getMonthlyProfitSummary(
  selectedMonth
) {

  selectedMonth =
    String(
      selectedMonth || ""
    ).trim();


  if (!selectedMonth) {

    return {
      sales: 0,
      purchases: 0,
      grossProfit: 0,
      profitMargin: null
    };

  }


  const ss =
    SpreadsheetApp.getActive();


  let sales = 0;
  let purchases = 0;


  // --------------------------------------------------------
  // SALES
  //
  // Orders:
  // B = Order Date
  // K = Grand Total
  // L = Status
  // --------------------------------------------------------

  const orderSheet =
    ss.getSheetByName("Orders");


  if (
    orderSheet &&
    orderSheet.getLastRow() > 1
  ) {

    const lastRow =
      orderSheet.getLastRow();

    /*
     * Read B:L once.
     *
     * Relative indexes:
     *
     * B -> 0  Order Date
     * K -> 9  Grand Total
     * L -> 10 Status
     */
    const orders =
      orderSheet
        .getRange(
          2,
          2,
          lastRow - 1,
          11
        )
        .getValues();


    for (
      let i = 0;
      i < orders.length;
      i++
    ) {

      const row =
        orders[i];

      const orderDate =
        row[0];

      if (!(orderDate instanceof Date)) {
        continue;
      }


      if (
        formatMonthYear(orderDate) !==
        selectedMonth
      ) {
        continue;
      }


      const status =
        String(
          row[10] || ""
        )
        .trim()
        .toLowerCase();


      if (status === "cancelled") {
        continue;
      }


      const grandTotal =
        Number(row[9]) || 0;


      sales +=
        grandTotal;

    }

  }


  // --------------------------------------------------------
  // PURCHASES
  //
  // Purchase Register:
  // B = Purchase Date
  // G = Amount
  // --------------------------------------------------------

  const purchaseSheet =
    ss.getSheetByName(
      "Purchase Register"
    );


  if (
    purchaseSheet &&
    purchaseSheet.getLastRow() > 1
  ) {

    const lastRow =
      purchaseSheet.getLastRow();

    /*
     * Read B:G once.
     *
     * Relative indexes:
     *
     * B -> 0 Purchase Date
     * G -> 5 Amount
     */
    const purchaseData =
      purchaseSheet
        .getRange(
          2,
          2,
          lastRow - 1,
          6
        )
        .getValues();


    for (
      let i = 0;
      i < purchaseData.length;
      i++
    ) {

      const row =
        purchaseData[i];

      const purchaseDate =
        row[0];


      if (!(purchaseDate instanceof Date)) {
        continue;
      }


      if (
        formatMonthYear(
          purchaseDate
        ) !== selectedMonth
      ) {
        continue;
      }


      purchases +=
        Number(row[5]) || 0;

    }

  }


  // --------------------------------------------------------
  // PROFIT
  // --------------------------------------------------------

  const grossProfit =
    sales - purchases;


  let profitMargin =
    null;


  if (sales > 0) {

    profitMargin =
      Number(
        (
          grossProfit /
          sales *
          100
        ).toFixed(1)
      );

  }


  return {

    sales: sales,

    purchases: purchases,

    grossProfit: grossProfit,

    profitMargin: profitMargin

  };

}