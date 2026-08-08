/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * WEB ORDER SERVICE
 * ==========================================================
 */

/**
 * ==========================================================
 * Save Order From Web
 * HIGH-SPEED OPTIMIZED VERSION
 * ==========================================================
 */
function saveOrderFromWeb(orderData) {

  const totalStart = Date.now();

  try {

    // ======================================================
    // 1. VALIDATION
    // ======================================================
    validateOrderData(orderData);


    // ======================================================
    // 2. LOAD SHEETS (Optimized: Removed unused Menu, Settings, Customers)
    // ======================================================
    const db = {
      orders: getSheet(SHEETS.ORDERS),
      orderItems: getSheet(SHEETS.ORDER_ITEMS)
    };


    // ======================================================
    // 3. GENERATE ORDER ID
    // ======================================================
    const orderId = generateNextId(db.orders, "SS");


    try {

      // ====================================================
      // 4. SAVE ORDER HEADER
      // ====================================================
      saveOrderHeader(db.orders, orderId, orderData);


      // ====================================================
      // 5. SAVE ORDER ITEMS (Already batch-optimized)
      // ====================================================
      saveOrderItems(db.orderItems, orderId, orderData);


      // ====================================================
      // 6. SAVE / UPDATE CUSTOMER 
      // (CustomerService handles its own sheet loading)
      // ====================================================
      saveOrUpdateCustomer(orderData, orderId);


      // ====================================================
      // 7. CLEAR HIGH-SPEED CACHE
      // ====================================================
      if (typeof clearAppCache === "function") {
        clearAppCache();
      }

      // ====================================================
      // SUCCESS
      // ====================================================
      const totalTime = Date.now() - totalStart;
      console.log("PERF | TOTAL SAVE ORDER:", totalTime, "ms");

      return {
        success: true,
        orderId: orderId,
        message: "Order Saved Successfully",
        performance: {
          totalMs: totalTime
        }
      };

    } catch (error) {

      // ====================================================
      // ROLLBACK
      // ====================================================
      rollbackOrder(db, orderId);
      throw error;

    }

  } catch (error) {

    console.error("Save Order Error:", error);
    return {
      success: false,
      message: error.message
    };

  }
}





/**
 * ==========================================================
 * VALIDATE ORDER DATA
 * ==========================================================
 */
function validateOrderData(orderData) {

  if (!orderData) {

    throw new Error(
      "Order data is missing."
    );

  }


  if (
    !orderData.customer ||
    !orderData.customer.customerName
  ) {

    throw new Error(
      "Customer Name is required."
    );

  }


  if (
    !orderData.customer.mobile
  ) {

    throw new Error(
      "Mobile Number is required."
    );

  }


  const mobile =
    String(
      orderData.customer.mobile
    ).trim();


  if (
    !/^\d{10}$/.test(mobile)
  ) {

    throw new Error(
      "Invalid Mobile Number."
    );

  }


  if (
    !Array.isArray(orderData.cart) ||
    orderData.cart.length === 0
  ) {

    throw new Error(
      "Cart is empty."
    );

  }


  if (
    !orderData.totals ||
    !Number.isFinite(
      Number(
        orderData.totals.grandTotal
      )
    ) ||
    Number(
      orderData.totals.grandTotal
    ) <= 0
  ) {

    throw new Error(
      "Grand Total should be greater than zero."
    );

  }


  // ========================================================
  // VALIDATE CUSTOM ADD-ONS
  // ========================================================

  if (
    orderData.customAddons !== undefined &&
    !Array.isArray(orderData.customAddons)
  ) {

    throw new Error(
      "Invalid Add-on data."
    );

  }


  if (
    Array.isArray(orderData.customAddons)
  ) {

    orderData.customAddons.forEach(
      function(addon) {

        if (!addon) {

          throw new Error(
            "Invalid Add-on."
          );

        }


        const name =
          String(
            addon.name || ""
          ).trim();


        const price =
          Number(
            addon.price
          );


        if (!name) {

          throw new Error(
            "Add-on Name is required."
          );

        }


        if (
          !Number.isFinite(price) ||
          price <= 0
        ) {

          throw new Error(
            "Invalid Add-on Price for " +
            name +
            "."
          );

        }

      }
    );

  }


  return true;

}






/**
 * ==========================================================
 * SAVE ORDER HEADER
 * ==========================================================
 */
function saveOrderHeader(ordersSheet, orderId, orderData) {

  const now = new Date();


  // ========================================================
  // SAFE CUSTOM ADD-ONS
  // ========================================================

  const rawAddons =
    Array.isArray(orderData.customAddons)
      ? orderData.customAddons
      : [];


  const validAddons = [];


  rawAddons.forEach(function(addon) {

    if (!addon) {
      return;
    }


    const name =
      String(addon.name || "").trim();


    const price =
      Number(addon.price);


    // Ignore malformed / empty add-ons
    if (
      !name ||
      !Number.isFinite(price) ||
      price <= 0
    ) {

      return;

    }


    validAddons.push({

      name: name,

      price: price

    });

  });


  // ========================================================
  // SERVER-SIDE ADD-ON TOTAL
  // ========================================================

  const addonTotal =
    validAddons.reduce(
      function(total, addon) {

        return total + addon.price;

      },
      0
    );


  // ========================================================
  // ADD-ON DESCRIPTION
  // ========================================================

  const addonDescription =
    validAddons
      .map(function(addon) {

        return (
          addon.name +
          " ₹" +
          addon.price.toFixed(2)
        );

      })
      .join(" | ");


  // ========================================================
  // ORDER ROW
  // ========================================================

  const row = [

    // A - Order ID
    orderId,


    // B - Order Date
    Utilities.formatDate(
      new Date(orderData.meta.orderDate),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    ),


    // C - Customer Name
    orderData.customer.customerName,


    // D - Mobile
    orderData.customer.mobile,


    // E - Delivery Area
    orderData.customer.deliveryArea,


    // F - House Address
    orderData.customer.houseAddress,


    // G - Delivery Slot
    orderData.customer.deliverySlot,


    // H - Order Type
    orderData.order.orderType,


    // I - Payment Mode
    orderData.payment.paymentMode,


    // J - Payment Status
    orderData.payment.paymentStatus,


    // K - Grand Total
    Number(orderData.totals.grandTotal) || 0,


    // L - Status
    "New",


    // M - Special Instructions
    orderData.order.specialInstructions || "",


    // N - Created At
    now,


    // O - Delivery Charge
    Number(orderData.totals.deliveryCharge) || 0,


    // P - Discount
    Number(orderData.totals.discount) || 0,


    // Q - Add-ons
    addonDescription,


    // R - Add-on Total
    addonTotal

  ];


  // ========================================================
  // SAVE ORDER
  // ========================================================

  const nextRow =
    ordersSheet.getLastRow() + 1;


  ordersSheet
    .getRange(
      nextRow,
      1,
      1,
      row.length
    )
    .setValues([row]);


  // ========================================================
  // CREATED AT FORMAT
  // ========================================================

  ordersSheet
    .getRange(
      nextRow,
      14
    )
    .setNumberFormat(
      "dd/MM/yyyy HH:mm:ss"
    );

}










function saveOrderItems(orderItemsSheet, orderId, orderData) {

  const rows = [];

  orderData.cart.forEach(item => {
    rows.push([
      orderId,
      item.itemName,
      item.qty,
      item.price,
      item.lineTotal
    ]);
  });

  if (rows.length > 0) {
    orderItemsSheet.getRange(
      orderItemsSheet.getLastRow() + 1,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);
  }

}







function rollbackOrder(db, orderId) {

  deleteRowsByOrderId(db.orders, orderId);
  deleteRowsByOrderId(db.orderItems, orderId);

}






function deleteRowsByOrderId(sheet, orderId) {

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {

    if (values[i][0] === orderId) {
      sheet.deleteRow(i + 2);
    }

  }

}





/**
 * ==========================================================
 * DELETE ORDER FROM WEB
 * ==========================================================
 *
 * Deletes:
 * 1. Order row from Orders
 * 2. All matching rows from Order Items
 *
 * Then recalculates:
 * 3. First Order Date
 * 4. Last Order Date
 * 5. Total Orders
 * 6. Lifetime Spend
 * 7. Last Order ID
 *
 * Customer row itself is NEVER deleted.
 * ==========================================================
 */
function deleteOrderFromWeb(orderId) {

  try {

    // ------------------------------------------------------
    // Validate Order ID
    // ------------------------------------------------------

    orderId = String(orderId || "")
      .trim()
      .toUpperCase();

    if (!orderId) {

      throw new Error("Order ID is required.");

    }


    // ------------------------------------------------------
    // Load Required Sheets
    // ------------------------------------------------------

    const ordersSheet =
      getSheet(SHEETS.ORDERS);

    const orderItemsSheet =
      getSheet(SHEETS.ORDER_ITEMS);

    const customerSheet =
      getSheet(SHEETS.CUSTOMERS);


    // ------------------------------------------------------
    // Find Order
    // ------------------------------------------------------

    const orderInfo =
      findOrderForDeletion(
        ordersSheet,
        orderId
      );


    if (!orderInfo) {

      throw new Error(
        "Order ID " +
        orderId +
        " was not found."
      );

    }


    // ------------------------------------------------------
    // Capture Customer Mobile BEFORE deleting order
    // ------------------------------------------------------

    const customerMobile =
      String(orderInfo.mobile || "").trim();


    if (!customerMobile) {

      throw new Error(
        "Customer mobile number is missing for Order " +
        orderId +
        ". Order was not deleted."
      );

    }


    // ------------------------------------------------------
    // Find Customer BEFORE deleting anything
    // ------------------------------------------------------

    const customerRow =
      findCustomerRowByMobile(
        customerSheet,
        customerMobile
      );


    if (customerRow === -1) {

      throw new Error(
        "Customer record was not found for Order " +
        orderId +
        ". Order was not deleted."
      );

    }


    // ------------------------------------------------------
    // Delete Order Items
    // ------------------------------------------------------

    deleteOrderItemsForOrder(
      orderItemsSheet,
      orderId
    );


    // ------------------------------------------------------
    // Delete Entire Order Row
    // ------------------------------------------------------

    ordersSheet.deleteRow(
      orderInfo.rowNumber
    );


    // ------------------------------------------------------
    // Recalculate Customer From Remaining Orders
    // ------------------------------------------------------

    recalculateCustomerAfterOrderDeletion(
      ordersSheet,
      customerSheet,
      customerRow,
      customerMobile
    );


    // ------------------------------------------------------
    // Success
    // ------------------------------------------------------

    return {

      success: true,

      orderId: orderId,

      message:
        "Order Deleted Successfully"

    };


  } catch (error) {

    console.error(
      "Delete Order Error:",
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : "Unable to delete order."

    };

  }

}



/**
 * ==========================================================
 * FIND ORDER FOR DELETION
 * ==========================================================
 *
 * Orders Sheet:
 *
 * A = Order ID
 * B = Order Date
 * C = Customer Name
 * D = Mobile
 * E = Delivery Area
 * F = House Address
 * G = Delivery Slot
 * H = Order Type
 * I = Payment Mode
 * J = Payment Status
 * K = Grand Total
 * L = Status
 * M = Special Instructions
 * N = Created At
 * O = Delivery Charge
 * P = Discount
 *
 * ==========================================================
 */
function findOrderForDeletion(
  ordersSheet,
  orderId
) {

  const lastRow =
    ordersSheet.getLastRow();


  if (lastRow <= 1) {

    return null;

  }


  // Search only Order ID column
  const match = ordersSheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .createTextFinder(orderId)
    .matchEntireCell(true)
    .findNext();


  if (!match) {

    return null;

  }


  const rowNumber =
    match.getRow();


  // Read the complete order BEFORE deletion
  const row = ordersSheet
    .getRange(
      rowNumber,
      1,
      1,
      16
    )
    .getValues()[0];


  return {

    rowNumber: rowNumber,

    orderId: row[0],

    orderDate: row[1],

    customerName: row[2],

    mobile: row[3],

    grandTotal:
      Number(row[10] || 0)

  };

}



/**
 * ==========================================================
 * DELETE ORDER ITEMS
 * ==========================================================
 *
 * Deletes ALL Order Items rows matching the Order ID.
 *
 * Iterates from bottom to top so row deletion does not
 * disturb the remaining row indexes.
 * ==========================================================
 */
function deleteOrderItemsForOrder(
  orderItemsSheet,
  orderId
) {

  const lastRow =
    orderItemsSheet.getLastRow();


  if (lastRow <= 1) {

    return;

  }


  const orderIds = orderItemsSheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues();


  for (
    let i = orderIds.length - 1;
    i >= 0;
    i--
  ) {

    if (
      String(orderIds[i][0]).trim() ===
      String(orderId).trim()
    ) {

      orderItemsSheet.deleteRow(
        i + 2
      );

    }

  }

}



/**
 * ==========================================================
 * RECALCULATE CUSTOMER AFTER ORDER DELETION
 * ==========================================================
 *
 * Rebuilds customer statistics using remaining Orders.
 *
 * Customers:
 *
 * A = Customer ID
 * B = Customer Name
 * C = Mobile
 * D = Delivery Area
 * E = House Address
 * F = First Order Date
 * G = Last Order Date
 * H = Total Orders
 * I = Lifetime Spend
 * J = Last Order ID
 *
 * ==========================================================
 */
function recalculateCustomerAfterOrderDeletion(
  ordersSheet,
  customerSheet,
  customerRow,
  customerMobile
) {

  const lastRow =
    ordersSheet.getLastRow();


  // --------------------------------------------------------
  // Read current customer record
  // --------------------------------------------------------

  const customerData =
    customerSheet
      .getRange(
        customerRow,
        1,
        1,
        10
      )
      .getValues()[0];


  // --------------------------------------------------------
  // No Orders Exist At All
  // --------------------------------------------------------

  if (lastRow <= 1) {

    clearCustomerOrderHistoryAfterDeletion(
      customerSheet,
      customerRow,
      customerData
    );

    return;

  }


  // --------------------------------------------------------
  // Read Required Order Data
  //
  // A = Order ID
  // B = Order Date
  // C = Customer Name
  // D = Mobile
  // ...
  // K = Grand Total
  // --------------------------------------------------------

  const orders =
    ordersSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        11
      )
      .getValues();


  const remainingOrders = [];


  // --------------------------------------------------------
  // Find Orders Belonging To This Customer
  // --------------------------------------------------------

  for (
    let i = 0;
    i < orders.length;
    i++
  ) {

    const mobile =
      String(orders[i][3] || "")
        .trim();


    if (
      mobile ===
      String(customerMobile).trim()
    ) {

      const parsedDate =
        parseStoredOrderDate(
          orders[i][1]
        );


      remainingOrders.push({

        orderId:
          String(
            orders[i][0] || ""
          ).trim(),

        orderDate:
          parsedDate,

        grandTotal:
          Number(
            orders[i][10] || 0
          )

      });

    }

  }


  // --------------------------------------------------------
  // Customer Has No Remaining Orders
  // --------------------------------------------------------

  if (remainingOrders.length === 0) {

    clearCustomerOrderHistoryAfterDeletion(
      customerSheet,
      customerRow,
      customerData
    );

    return;

  }


  // --------------------------------------------------------
  // Sort Orders By Business Order Date
  // --------------------------------------------------------

  remainingOrders.sort(
    function(a, b) {

      return (
        a.orderDate.getTime() -
        b.orderDate.getTime()
      );

    }
  );


  // --------------------------------------------------------
  // First / Last Order
  // --------------------------------------------------------

  const firstOrder =
    remainingOrders[0];

  const lastOrder =
    remainingOrders[
      remainingOrders.length - 1
    ];


  // --------------------------------------------------------
  // Lifetime Spend
  // --------------------------------------------------------

  let lifetimeSpend = 0;


  remainingOrders.forEach(
    function(order) {

      lifetimeSpend +=
        Number(
          order.grandTotal || 0
        );

    }
  );


  // --------------------------------------------------------
  // Update Customer A:J
  // --------------------------------------------------------

  customerSheet
    .getRange(
      customerRow,
      1,
      1,
      10
    )
    .setValues([[

      customerData[0],             // A Customer ID

      customerData[1],             // B Customer Name

      customerData[2],             // C Mobile

      customerData[3],             // D Delivery Area

      customerData[4],             // E House Address

      firstOrder.orderDate,        // F First Order Date

      lastOrder.orderDate,         // G Last Order Date

      remainingOrders.length,      // H Total Orders

      lifetimeSpend,               // I Lifetime Spend

      lastOrder.orderId            // J Last Order ID

    ]]);

}



/**
 * ==========================================================
 * CLEAR CUSTOMER ORDER HISTORY
 * ==========================================================
 *
 * Used when the deleted order was the customer's
 * only remaining order.
 *
 * Customer identity/address information is preserved.
 * ==========================================================
 */
function clearCustomerOrderHistoryAfterDeletion(
  customerSheet,
  customerRow,
  customerData
) {

  customerSheet
    .getRange(
      customerRow,
      1,
      1,
      10
    )
    .setValues([[

      customerData[0],     // A Customer ID

      customerData[1],     // B Customer Name

      customerData[2],     // C Mobile

      customerData[3],     // D Delivery Area

      customerData[4],     // E House Address

      "",                  // F First Order Date

      "",                  // G Last Order Date

      0,                   // H Total Orders

      0,                   // I Lifetime Spend

      ""                   // J Last Order ID

    ]]);

}



/**
 * ==========================================================
 * PARSE STORED ORDER DATE
 * ==========================================================
 *
 * Handles:
 *
 * 1. Native Google Sheets Date values
 *
 * 2. Existing Orders saved as:
 *    dd/MM/yyyy
 *
 * Returns a Date at noon to avoid timezone date shifting.
 * ==========================================================
 */
function parseStoredOrderDate(value) {

  // --------------------------------------------------------
  // Native Date
  // --------------------------------------------------------

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]" &&
    !isNaN(value.getTime())
  ) {

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      12,
      0,
      0
    );

  }


  // --------------------------------------------------------
  // dd/MM/yyyy String
  // --------------------------------------------------------

  const text =
    String(value || "").trim();


  const parts =
    text.split("/");


  if (parts.length !== 3) {

    throw new Error(
      "Invalid Order Date found while recalculating customer history: " +
      text
    );

  }


  const day =
    Number(parts[0]);

  const month =
    Number(parts[1]);

  const year =
    Number(parts[2]);


  if (
    !day ||
    !month ||
    !year
  ) {

    throw new Error(
      "Invalid Order Date found while recalculating customer history: " +
      text
    );

  }


  return new Date(
    year,
    month - 1,
    day,
    12,
    0,
    0
  );

}
