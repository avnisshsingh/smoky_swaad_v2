/**
 * ==========================================
 * SMOKY SWAAD ERP
 * Save Order Module
 * ==========================================
 */

/**
 * Main Function (Optimized)
 */
function saveOrder() {

  // Validate Order
  const validation = validateOrder();

  if (!validation.valid) {

    showError(
      "Validation Failed\n\n" +
      validation.errors.join("\n")
    );

    return;

  }

  // Generate Order ID
  const orderID = generateOrderID();

  // Save Header
  saveOrderHeader(orderID);

  // Save Items
  saveOrderItems(orderID);

  // Update Customer
  updateCustomer();

  // Clear POS
  clearPOS();

  // Apply pending sheet changes
  SpreadsheetApp.flush();

  // Success Message
  showSuccess("Order Saved : " + orderID);

}


/**
 * Generate Order ID
 */
function generateOrderID() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const config = ss.getSheetByName(SHEETS.CONFIG);

  let lastNumber = Number(config.getRange("B1").getValue());

  lastNumber++;

  config.getRange("B1").setValue(lastNumber);

  return "SS" + lastNumber.toString().padStart(5, "0");

}

/**
 * Save Order Header (Optimized)
 */
function saveOrderHeader(orderID) {

  const pos = getSheet(SHEETS.POS);
  const orders = getSheet(SHEETS.ORDERS);

  const nextRow = getNextRow(orders);

  // Read entire form once
  const form = pos.getRange("B2:B30").getValues();

  const rowData = [[
    orderID,        // A Order ID
    form[2][0],     // B Order Date
    form[0][0],     // C Customer Name
    form[1][0],     // D Mobile
    form[3][0],     // E Delivery Area
    form[4][0],     // F House Address
    form[5][0],     // G Delivery Slot
    form[6][0],     // H Order Type
    form[7][0],     // I Payment Mode
    form[8][0],     // J Payment Status
    form[28][0],    // K Grand Total (B30)
    "New",          // L Status
    form[9][0],     // M Special Instructions
    new Date()      // N Created At
  ]];

  orders
    .getRange(nextRow, 1, 1, rowData[0].length)
    .setValues(rowData);

}

/**
 * Save Order Items (Optimized)
 */
function saveOrderItems(orderID) {

  const pos = getSheet(SHEETS.POS);
  const orderItems = getSheet(SHEETS.ORDER_ITEMS);

  const nextRow = getNextRow(orderItems);

  // Read all order rows in ONE API call
  const items = pos.getRange(
    POS.ITEM_START_ROW,
    POS.ITEM_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    4
  ).getValues();

  const data = [];

  items.forEach(row => {

    const item  = row[0];
    const price = row[1];
    const qty   = row[2];
    const total = row[3];

    if (item !== "" && item !== null) {

      data.push([
        orderID,
        item,
        qty,
        price,
        total
      ]);

    }

  });

  if (data.length > 0) {

    orderItems
      .getRange(nextRow, 1, data.length, data[0].length)
      .setValues(data);

  }

}

/**
 * Add or Update Customer (Optimized)
 */
function updateCustomer() {

  const pos = getSheet(SHEETS.POS);
  const customers = getSheet(SHEETS.CUSTOMERS);

  // Read POS data in one call
  const formData = pos.getRange("B2:B30").getValues();

  const customerName = formData[0][0];
  const mobile = formData[1][0];
  const orderDate = formData[2][0];
  const deliveryArea = formData[3][0];
  const houseAddress = formData[4][0];
  const grandTotal = Number(formData[28][0]);

  // Read Customers sheet once
  const lastRow = customers.getLastRow();

  if (lastRow > 1) {

    const customerData = customers.getRange(2, 1, lastRow - 1, 9).getValues();

    for (let i = 0; i < customerData.length; i++) {

      // Column C = Mobile Number
      if (customerData[i][2] == mobile) {

        customerData[i][1] = customerName;                     // Customer Name
        customerData[i][3] = deliveryArea;                     // Delivery Area
        customerData[i][4] = houseAddress;                     // House Address
        customerData[i][6] = orderDate;                        // Last Order Date
        customerData[i][7] = Number(customerData[i][7]) + 1;   // Total Orders
        customerData[i][8] = Number(customerData[i][8]) + grandTotal; // Lifetime Spend

        // Write only one row back
        customers
          .getRange(i + 2, 1, 1, 9)
          .setValues([customerData[i]]);

        return;

      }

    }

  }

  // New Customer
  const customerID = "CUS" + String(lastRow).padStart(5, "0");

  customers.appendRow([
    customerID,
    customerName,
    mobile,
    deliveryArea,
    houseAddress,
    orderDate,
    orderDate,
    1,
    grandTotal
  ]);

}

/**
 * Clear POS Screen (Optimized)
 */
function clearPOS() {

  const pos = getSheet(SHEETS.POS);

  // Customer Details
  pos.getRange("B2:B11").clearContent();

  // Menu Items (Column A)
  pos.getRange(
    POS.ITEM_START_ROW,
    POS.ITEM_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    1
  ).clearContent();

  // Quantity (Column C)
  pos.getRange(
    POS.ITEM_START_ROW,
    POS.QTY_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    1
  ).clearContent();

  // Reset Discount
  pos.getRange(POS.DISCOUNT).setValue(0);

  // Reset Delivery Charge
  pos.getRange(POS.DELIVERY_CHARGE).setValue(0);

}