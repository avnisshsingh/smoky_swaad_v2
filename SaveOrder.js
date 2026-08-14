/**
 * ==========================================
 * SMOKY SWAAD ERP
 * Save Order Module (Supports both Web UI payload & direct execution)
 * ==========================================
 */

/**
 * Main Function - Handles orders coming from Web UI (orderData) or spreadsheet cells.
 */
function saveOrder(orderData) {
  try {
    // If called from the Web UI with an order payload object
    if (orderData && (orderData.cart || orderData.customer)) {
      return saveOrderFromWebPayload(orderData);
    }

    // Fallback for legacy spreadsheet execution
    const validation = validateOrder();

    if (!validation.valid) {
      showError(
        "Validation Failed\n\n" +
        validation.errors.join("\n")
      );
      return { success: false, message: "Validation failed" };
    }

    const orderID = generateOrderID();
    saveOrderHeader(orderID);
    saveOrderItems(orderID);
    updateCustomer();
    clearPOS();

    SpreadsheetApp.flush();
    showSuccess("Order Saved : " + orderID);
    return { success: true, orderId: orderID };

  } catch (error) {
    console.error("Error in saveOrder: " + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Handles saving order when triggered from the Web UI with orderData
 */
function saveOrderFromWebPayload(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName(SHEETS.ORDERS) || ss.getSheetByName("Orders");
  const itemsSheet = ss.getSheetByName(SHEETS.ORDER_ITEMS) || ss.getSheetByName("OrderItems");
  
  if (!ordersSheet) throw new Error("Orders sheet not found.");

  // 1. Generate Order ID
  const orderID = generateOrderID();
  const orderDate = orderData.order && orderData.order.orderDate ? orderData.order.orderDate : new Date();

  // 2. Format Addons string
  const addonNames = (orderData.customAddons || []).map(a => a.name + " - ₹" + a.price).join(", ");

  // 3. Extract Customer & Order Details
  const customerName = orderData.customer ? orderData.customer.customerName : "";
  const mobile = orderData.customer ? orderData.customer.mobile : "";
  const deliveryArea = orderData.customer ? orderData.customer.deliveryArea : "";
  const houseAddress = orderData.customer ? orderData.customer.houseAddress : "";
  const deliverySlot = orderData.customer ? orderData.customer.deliverySlot : "";
  const orderType = orderData.order ? orderData.order.orderType : "";
  const paymentMode = orderData.payment ? orderData.payment.paymentMode : "";
  const paymentStatus = orderData.payment ? orderData.payment.paymentStatus : "";
  const grandTotal = orderData.totals ? orderData.totals.grandTotal : 0;
  const specialInstructions = orderData.order ? orderData.order.specialInstructions : "";
  const deliveryCharge = orderData.totals ? orderData.totals.deliveryCharge : 0;
  const discount = orderData.totals ? orderData.totals.discount : 0;
  const addonTotal = orderData.totals ? orderData.totals.addonTotal : 0;

  // 4. Save Header to Orders sheet
  const headerRow = [
    orderID,                // A: Order ID
    orderDate,              // B: Order Date
    customerName,           // C: Customer Name
    mobile,                 // D: Mobile
    deliveryArea,           // E: Delivery Area
    houseAddress,           // F: House Address
    deliverySlot,           // G: Delivery Slot
    orderType,              // H: Order Type
    paymentMode,            // I: Payment Mode
    paymentStatus,          // J: Payment Status
    grandTotal,             // K: Grand Total
    "New",                  // L: Status
    specialInstructions,    // M: Special Instructions
    new Date(),             // N: Created At
    deliveryCharge,         // O: Delivery Charge
    discount,               // P: Discount
    addonNames,             // Q: Add-ons
    addonTotal              // R: Add-on Total
  ];

  ordersSheet.appendRow(headerRow);

  // 5. Resolve Customer ID via Mobile
  const customerID = getCustomerIDByMobile(mobile);

  // 6. Save Order Items to OrderItems sheet (Columns A through G mapped correctly)
  if (itemsSheet && orderData.cart && orderData.cart.length > 0) {
    orderData.cart.forEach(function(item) {
      itemsSheet.appendRow([
        orderID,                  // Col A: Order ID
        item.itemName || item[0], // Col B: Menu Item
        item.qty || item[2],      // Col C: Qty
        item.price || item[1],    // Col D: Unit Price
        item.lineTotal || item[3],// Col E: Total
        customerID,               // Col F: Customer ID
        orderDate                 // Col G: Order Date
      ]);
    });
  }

  // 7. Update Customer database stats
  updateCustomerFromPayload(orderData);

  SpreadsheetApp.flush();
  return { success: true, orderId: orderID };
}

/**
 * Helper to update customer stats from web payload
 */
function updateCustomerFromPayload(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customers = ss.getSheetByName(SHEETS.CUSTOMERS) || ss.getSheetByName("Customers");
  if (!customers) return;

  const customerName = orderData.customer ? orderData.customer.customerName : "";
  const mobile = orderData.customer ? orderData.customer.mobile : "";
  const orderDate = orderData.order && orderData.order.orderDate ? orderData.order.orderDate : new Date();
  const deliveryArea = orderData.customer ? orderData.customer.deliveryArea : "";
  const houseAddress = orderData.customer ? orderData.customer.houseAddress : "";
  const grandTotal = orderData.totals ? Number(orderData.totals.grandTotal) || 0 : 0;

  const lastRow = customers.getLastRow();
  if (lastRow > 1) {
    const customerData = customers.getRange(2, 1, lastRow - 1, 9).getValues();
    for (let i = 0; i < customerData.length; i++) {
      if (String(customerData[i][2]).trim() === String(mobile).trim()) {
        customerData[i][1] = customerName;
        customerData[i][3] = deliveryArea;
        customerData[i][4] = houseAddress;
        customerData[i][6] = orderDate;
        customerData[i][7] = Number(customerData[i][7]) + 1;
        customerData[i][8] = Number(customerData[i][8]) + grandTotal;
        customers.getRange(i + 2, 1, 1, 9).setValues([customerData[i]]);
        return;
      }
    }
  }

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
 * Save Order Header (Optimized for spreadsheet sheet execution)
 */
function saveOrderHeader(orderID) {
  const pos = getSheet(SHEETS.POS);
  const orders = getSheet(SHEETS.ORDERS);
  const nextRow = getNextRow(orders);

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

  orders.getRange(nextRow, 1, 1, rowData[0].length).setValues(rowData);
}

/**
 * Save Order Items (Optimized for spreadsheet sheet execution)
 */
function saveOrderItems(orderID) {
  const pos = getSheet(SHEETS.POS);
  const orderItems = getSheet(SHEETS.ORDER_ITEMS);
  const nextRow = getNextRow(orderItems);

  const form = pos.getRange("B2:B30").getValues();
  const orderDate = form[2][0];
  const mobile = form[1][0];
  const customerID = getCustomerIDByMobile(mobile);

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
        orderID,     // Col A: Order ID
        item,        // Col B: Menu Item
        qty,         // Col C: Qty
        price,       // Col D: Unit Price
        total,       // Col E: Total
        customerID,  // Col F: Customer ID
        orderDate    // Col G: Order Date
      ]);
    }
  });

  if (data.length > 0) {
    orderItems.getRange(nextRow, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * Unified Helper function to find Customer ID from the Customers sheet using mobile number
 */
function getCustomerIDByMobile(mobile) {
  if (!mobile) return "";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customers = ss.getSheetByName(SHEETS.CUSTOMERS);
    if (!customers) return "";
    
    const lastRow = customers.getLastRow();
    if (lastRow < 2) return "";

    const customerData = customers.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < customerData.length; i++) {
      if (customerData[i][2] && String(customerData[i][2]).trim() === String(mobile).trim()) {
        return customerData[i][0];
      }
    }
  } catch (e) {
    // Fallback gracefully
  }
  return "";
}

/**
 * Add or Update Customer (Optimized)
 */
function updateCustomer() {
  const pos = getSheet(SHEETS.POS);
  const customers = getSheet(SHEETS.CUSTOMERS);

  const formData = pos.getRange("B2:B30").getValues();

  const customerName = formData[0][0];
  const mobile = formData[1][0];
  const orderDate = formData[2][0];
  const deliveryArea = formData[3][0];
  const houseAddress = formData[4][0];
  const grandTotal = Number(formData[28][0]);

  const lastRow = customers.getLastRow();

  if (lastRow > 1) {
    const customerData = customers.getRange(2, 1, lastRow - 1, 9).getValues();

    for (let i = 0; i < customerData.length; i++) {
      if (customerData[i][2] == mobile) {
        customerData[i][1] = customerName;
        customerData[i][3] = deliveryArea;
        customerData[i][4] = houseAddress;
        customerData[i][6] = orderDate;
        customerData[i][7] = Number(customerData[i][7]) + 1;
        customerData[i][8] = Number(customerData[i][8]) + grandTotal;

        customers.getRange(i + 2, 1, 1, 9).setValues([customerData[i]]);
        return;
      }
    }
  }

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

  pos.getRange("B2:B11").clearContent();

  pos.getRange(
    POS.ITEM_START_ROW,
    POS.ITEM_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    1
  ).clearContent();

  pos.getRange(
    POS.ITEM_START_ROW,
    POS.QTY_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    1
  ).clearContent();

  pos.getRange(POS.DISCOUNT).setValue(0);
  pos.getRange(POS.DELIVERY_CHARGE).setValue(0);
}

/**
 * Update Order From Web (Preserved fully)
 */
function updateOrderFromWeb(orderId, orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName("Orders");
    const itemsSheet = ss.getSheetByName("OrderItems");
    
    if (!ordersSheet) throw new Error("Orders sheet not found.");
    
    const ordersData = ordersSheet.getDataRange().getValues();
    let orderRowIndex = -1;
    let orderDate = "";
    for (let i = 1; i < ordersData.length; i++) {
      if (String(ordersData[i][0]).trim().toUpperCase() === String(orderId).trim().toUpperCase()) {
        orderRowIndex = i + 1;
        orderDate = ordersData[i][1];
        break;
      }
    }
    
    if (orderRowIndex === -1) {
      return { success: false, message: "Order ID not found for update." };
    }
    
    if (itemsSheet) {
      const itemsData = itemsSheet.getDataRange().getValues();
      for (let i = itemsData.length - 1; i >= 1; i--) {
        if (String(itemsData[i][0]).trim().toUpperCase() === String(orderId).trim().toUpperCase()) {
          itemsSheet.deleteRow(i + 1);
        }
      }
    }
    
    const mobile = orderData.customer && orderData.customer.mobile ? orderData.customer.mobile : "";
    const customerID = getCustomerIDByMobile(mobile);

    if (itemsSheet && orderData.cart && orderData.cart.length > 0) {
      orderData.cart.forEach(function(item) {
        itemsSheet.appendRow([
          orderId,
          item.itemName,
          item.qty,
          item.price,
          item.lineTotal,
          customerID,
          orderDate
        ]);
      });
    }
    
    const addonNames = (orderData.customAddons || []).map(a => a.name + " - ₹" + a.price).join(", ");
    
    ordersSheet.getRange(orderRowIndex, 3).setValue(orderData.customer.customerName);
    ordersSheet.getRange(orderRowIndex, 4).setValue(orderData.customer.mobile);
    ordersSheet.getRange(orderRowIndex, 5).setValue(orderData.customer.deliveryArea);
    ordersSheet.getRange(orderRowIndex, 6).setValue(orderData.customer.houseAddress);
    ordersSheet.getRange(orderRowIndex, 7).setValue(orderData.customer.deliverySlot);
    ordersSheet.getRange(orderRowIndex, 8).setValue(orderData.order.orderType);
    ordersSheet.getRange(orderRowIndex, 9).setValue(orderData.payment.paymentMode);
    ordersSheet.getRange(orderRowIndex, 10).setValue(orderData.payment.paymentStatus);
    ordersSheet.getRange(orderRowIndex, 11).setValue(orderData.totals.grandTotal);
    ordersSheet.getRange(orderRowIndex, 13).setValue(orderData.order.specialInstructions);
    ordersSheet.getRange(orderRowIndex, 15).setValue(orderData.totals.deliveryCharge);
    ordersSheet.getRange(orderRowIndex, 16).setValue(orderData.totals.discount);
    ordersSheet.getRange(orderRowIndex, 17).setValue(addonNames);
    ordersSheet.getRange(orderRowIndex, 18).setValue(orderData.totals.addonTotal);
    
    return { success: true, orderId: orderId };
  } catch (error) {
    return { success: false, message: error.message };
  }
}