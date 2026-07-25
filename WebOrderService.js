/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * WEB ORDER SERVICE
 * ==========================================================
 */

function saveOrderFromWeb(orderData) {

  try {

    validateOrderData(orderData);

    const db = {
      orders: getSheet(SHEETS.ORDERS),
      orderItems: getSheet(SHEETS.ORDER_ITEMS),
      customers: getSheet(SHEETS.CUSTOMERS),
      menu: getSheet(SHEETS.MENU),
      settings: getSheet(SHEETS.SETTINGS)
    };

   const orderId = generateNextId(
  db.orders,
  "SS"
);

    try {

      saveOrderHeader(db.orders, orderId, orderData);
      saveOrderItems(db.orderItems, orderId, orderData);
      saveOrUpdateCustomer(orderData);

      return {
        success: true,
        orderId: orderId,
        message: "Order Saved Successfully"
      };

    } catch (error) {

      rollbackOrder(db, orderId);
      throw error;

    }

  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

function validateOrderData(orderData) {

  if (!orderData.customer.customerName) throw new Error("Customer Name is required.");
  if (!orderData.customer.mobile) throw new Error("Mobile Number is required.");
  if (orderData.customer.mobile.length !== 10) throw new Error("Invalid Mobile Number.");
  if (!orderData.cart || orderData.cart.length === 0) throw new Error("Cart is empty.");
  if (orderData.totals.grandTotal <= 0) throw new Error("Grand Total should be greater than zero.");

  return true;

}


function saveOrderHeader(ordersSheet, orderId, orderData) {

  const now = new Date();

  const row = [

    // A
    orderId,

    // B
    Utilities.formatDate(
      new Date(orderData.meta.orderDate),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    ),

    // C
    orderData.customer.customerName,

    // D
    orderData.customer.mobile,

    // E
    orderData.customer.deliveryArea,

    // F
    orderData.customer.houseAddress,

    // G
    orderData.customer.deliverySlot,

    // H
    orderData.order.orderType,

    // I
    orderData.payment.paymentMode,

    // J
    orderData.payment.paymentStatus,

    // K
    Number(orderData.totals.grandTotal) || 0,

    // L
    "New",

    // M
    orderData.order.specialInstructions || "",

    // N
    now,

    // O
    Number(orderData.totals.deliveryCharge) || 0,

    // P
    Number(orderData.totals.discount) || 0

  ];

  ordersSheet.appendRow(row);

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
