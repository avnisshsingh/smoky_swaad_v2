/**
 * ==========================================
 * Get Reports Data
 * PERFORMANCE PROFILING VERSION
 * ==========================================
 *
 * IMPORTANT:
 * This version does NOT change report calculations.
 * It only measures execution time of each stage.
 *
 * Use this version to establish the performance
 * baseline before making optimization changes.
 * ==========================================
 */
function getReportsData(fromDate, toDate) {

  const totalStart = Date.now();

  console.log(
    "=========================================="
  );

  console.log(
    "PERF | Reports | START"
  );

  console.log(
    "PERF | Reports | From Date:",
    fromDate
  );

  console.log(
    "PERF | Reports | To Date:",
    toDate
  );


  let orders = [];


  // ==========================================
  // Sheets
  // ==========================================

  let start = Date.now();

  const sheet =
    getSheet(
      SHEETS.ORDERS
    );

  const orderItemsSheet =
    getSheet(
      SHEETS.ORDER_ITEMS
    );

  const menuSheet =
    getSheet(
      SHEETS.MENU
    );


  console.log(
    "PERF | Reports | Get Sheets:",
    Date.now() - start,
    "ms"
  );


  // ==========================================
  // Orders Data
  // ==========================================

  start = Date.now();

  const data =
    sheet
      .getDataRange()
      .getValues();

  data.shift();


  console.log(
    "PERF | Reports | Read Orders:",
    Date.now() - start,
    "ms",
    "| Rows:",
    data.length
  );


  // ==========================================
  // Order Items Data
  // ==========================================

  start = Date.now();

  const orderItems =
    orderItemsSheet
      .getDataRange()
      .getValues();

  orderItems.shift();


  console.log(
    "PERF | Reports | Read Order Items:",
    Date.now() - start,
    "ms",
    "| Rows:",
    orderItems.length
  );


  // ==========================================
  // Order Items Pre-indexing Map
  // ==========================================

  start = Date.now();

  const orderItemsMap = {};


  orderItems.forEach(
    function (itemRow) {

      const orderId =
        String(
          itemRow[0]
        ).trim();


      if (
        !orderItemsMap[orderId]
      ) {

        orderItemsMap[orderId] = [];

      }


      orderItemsMap[orderId].push(
        itemRow
      );

    }
  );


  console.log(
    "PERF | Reports | Build Order Items Map:",
    Date.now() - start,
    "ms",
    "| Orders Indexed:",
    Object.keys(orderItemsMap).length
  );


  // ==========================================
  // Menu Data
  // ==========================================

  start = Date.now();

  const menuData =
    menuSheet
      .getDataRange()
      .getValues();

  menuData.shift();


  console.log(
    "PERF | Reports | Read Menu:",
    Date.now() - start,
    "ms",
    "| Rows:",
    menuData.length
  );


  // ==========================================
  // Menu Cost Lookup
  // ==========================================

  start = Date.now();

  const menuCostMap = {};
  const topSellingMap = {};


  menuData.forEach(
    function (row) {

      const menuItem =
        String(
          row[1]
        ).trim();


      menuCostMap[menuItem] =
        Number(
          row[3]
        ) || 0;

    }
  );


  console.log(
    "PERF | Reports | Build Menu Cost Map:",
    Date.now() - start,
    "ms",
    "| Menu Items:",
    Object.keys(menuCostMap).length
  );


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


  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );


  // ==========================================
  // Process Orders
  // ==========================================

  start = Date.now();


  data.forEach(
    function (row) {

      const orderId =
        String(
          row[0]
        ).trim();


      const orderDate =
        Utilities.formatDate(
          new Date(row[1]),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );


      const grandTotal =
        Number(
          row[10]
        ) || 0;


      // ----------------------------------------
      // Get items from pre-indexed map
      // ----------------------------------------

      const itemsForOrder =
        orderItemsMap[orderId] || [];


      let orderCost = 0;


      itemsForOrder.forEach(
        function (itemRow) {

          const menuItem =
            String(
              itemRow[1]
            ).trim();


          const qty =
            Number(
              itemRow[2]
            ) || 0;


          const costPrice =
            menuCostMap[menuItem] || 0;


          orderCost +=
            qty * costPrice;

        }
      );


      const orderProfit =
        grandTotal -
        orderCost;


      // ========================================
      // Today's KPI
      // ========================================

      if (
        orderDate === today
      ) {

        todayOrders++;

        todaySales +=
          grandTotal;

        todayProfit +=
          orderProfit;

      }


      // ========================================
      // Date Range Report
      // ========================================

      if (
        fromDate &&
        toDate &&
        orderDate >= fromDate &&
        orderDate <= toDate
      ) {

        totalOrders++;

        totalSales +=
          grandTotal;

        totalCost +=
          orderCost;

        grossProfit +=
          orderProfit;


        itemsForOrder.forEach(
          function (itemRow) {

            const menuItem =
              String(
                itemRow[1]
              ).trim();


            const qty =
              Number(
                itemRow[2]
              ) || 0;


            const costPrice =
              menuCostMap[menuItem] || 0;


            if (
              !topSellingMap[menuItem]
            ) {

              topSellingMap[menuItem] = {

                menuItem:
                  menuItem,

                quantity:
                  0,

                revenue:
                  0,

                cost:
                  0,

                profit:
                  0

              };

            }


            topSellingMap[
              menuItem
            ].quantity +=
              qty;


            topSellingMap[
              menuItem
            ].revenue +=
              qty *
              Number(
                itemRow[3]
              );


            topSellingMap[
              menuItem
            ].cost +=
              qty *
              costPrice;


            topSellingMap[
              menuItem
            ].profit =
              topSellingMap[
                menuItem
              ].revenue -
              topSellingMap[
                menuItem
              ].cost;

          }
        );


        orders.push({

          orderId:
            row[0],

          orderDate:
            Utilities.formatDate(
              new Date(row[1]),
              Session.getScriptTimeZone(),
              "dd/MM/yyyy"
            ),

          customer:
            row[2],

          mobile:
            row[3],

          paymentMode:
            row[8],

          paymentStatus:
            row[9],

          grandTotal:
            grandTotal

        });

      }

    }
  );


  console.log(
    "PERF | Reports | Process Orders:",
    Date.now() - start,
    "ms",
    "| Orders Processed:",
    data.length,
    "| Orders Matched:",
    totalOrders
  );


  // ==========================================
  // Top Selling Items
  // ==========================================

  start = Date.now();


  const topSellingItems =
    Object.values(
      topSellingMap
    )
    .sort(
      function (a, b) {

        return (
          b.quantity -
          a.quantity
        );

      }
    );


  console.log(
    "PERF | Reports | Build Top Selling:",
    Date.now() - start,
    "ms",
    "| Items:",
    topSellingItems.length
  );


  // ==========================================
  // TOTAL PERFORMANCE
  // ==========================================

  const totalTime =
    Date.now() -
    totalStart;


  console.log(
    "PERF | Reports | TOTAL:",
    totalTime,
    "ms"
  );


  console.log(
    "=========================================="
  );


  // ==========================================
  // RETURN
  // ==========================================

  return {

    success:
      true,

    todaySales:
      todaySales,

    todayOrders:
      todayOrders,

    todayProfit:
      todayProfit,

    totalOrders:
      totalOrders,

    totalSales:
      totalSales,

    totalCost:
      totalCost,

    grossProfit:
      grossProfit,

    profitPercentage:
      totalSales > 0
        ? (
            (grossProfit / totalSales) *
            100
          ).toFixed(2)
        : "0.00",

    topSellingItems:
      topSellingItems,

    orders:
      orders

  };

}