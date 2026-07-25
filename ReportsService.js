/**
 * ==========================================
 * Get Reports Data
 * ==========================================
 */
function getReportsData(fromDate, toDate) {

  let orders = [];

  // ==========================================
  // Sheets
  // ==========================================
  const sheet = getSheet(SHEETS.ORDERS);
  const orderItemsSheet = getSheet(SHEETS.ORDER_ITEMS);
  const menuSheet = getSheet(SHEETS.MENU);

  // ==========================================
  // Orders Data
  // ==========================================
  const data = sheet.getDataRange().getValues();
  data.shift();

  // ==========================================
  // Order Items Data
  // ==========================================
  const orderItems = orderItemsSheet.getDataRange().getValues();
  orderItems.shift();

  // ==========================================
  // Menu Data
  // ==========================================
  const menuData = menuSheet.getDataRange().getValues();
  menuData.shift();

  // ==========================================
  // Menu Cost Lookup
  // ==========================================
  const menuCostMap = {};

  const topSellingMap = {};

  menuData.forEach(function (row) {

    const menuItem = String(row[1]).trim();

    menuCostMap[menuItem] = Number(row[3]) || 0;

  });

  // ==========================================
  // Report Variables
  // ==========================================
  let totalOrders = 0;
  let totalSales = 0;
  let totalCost = 0;
  let grossProfit = 0;

  let todayOrders = 0;
  let todaySales = 0;
  let todayProfit = 0;

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  // ==========================================
  // Process Orders
  // ==========================================
  data.forEach(function (row) {

    const orderId = String(row[0]).trim();

    const orderDate = Utilities.formatDate(
      new Date(row[1]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    const grandTotal = Number(row[10]) || 0;

    // ==========================================
    // Calculate Food Cost For This Order
    // ==========================================
    let orderCost = 0;

    orderItems.forEach(function (itemRow) {

      if (String(itemRow[0]).trim() !== orderId) {

        return;

      }

      const menuItem = String(itemRow[1]).trim();

      const qty = Number(itemRow[2]) || 0;

      const costPrice = menuCostMap[menuItem] || 0;

      orderCost += qty * costPrice;

    });

    const orderProfit = grandTotal - orderCost;

    // ==========================================
    // Today's KPI
    // ==========================================
    if (orderDate === today) {

      todayOrders++;

      todaySales += grandTotal;

      todayProfit += orderProfit;

    }

    // ==========================================
    // Date Range Report
    // ==========================================
    if (fromDate && toDate &&
      orderDate >= fromDate &&
      orderDate <= toDate) {

      totalOrders++;

      totalSales += grandTotal;

      totalCost += orderCost;

      grossProfit += orderProfit;

      orderItems.forEach(function (itemRow) {

        if (String(itemRow[0]).trim() !== orderId) {

          return;

        }

        const menuItem = String(itemRow[1]).trim();

        const qty = Number(itemRow[2]) || 0;

        const costPrice = menuCostMap[menuItem] || 0;

        if (!topSellingMap[menuItem]) {

          topSellingMap[menuItem] = {

            menuItem: menuItem,

            quantity: 0,

            revenue: 0,

            cost: 0,

            profit: 0

          };

        }

        topSellingMap[menuItem].quantity += qty;

        topSellingMap[menuItem].revenue += qty * Number(itemRow[3]);

        topSellingMap[menuItem].cost += qty * costPrice;

        topSellingMap[menuItem].profit =
          topSellingMap[menuItem].revenue -
          topSellingMap[menuItem].cost;

      });

      orders.push({

        orderId: row[0],

        orderDate: Utilities.formatDate(
          new Date(row[1]),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        ),

        customer: row[2],

        mobile: row[3],

        paymentMode: row[8],

        paymentStatus: row[9],

        grandTotal: grandTotal

      });

    }

  });

  const topSellingItems = Object.values(topSellingMap)

    .sort(function (a, b) {

      return b.quantity - a.quantity;

    });

  return {

    success: true,

    todaySales: todaySales,

    todayOrders: todayOrders,

    todayProfit: todayProfit,

    totalOrders: totalOrders,

    totalSales: totalSales,

    totalCost: totalCost,

    grossProfit: grossProfit,

    profitPercentage:
      totalSales > 0
        ? ((grossProfit / totalSales) * 100).toFixed(2)
        : "0.00",



    topSellingItems: topSellingItems,

    orders: orders


  };

}