/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * CUSTOMER SERVICE
 * ==========================================================
 */


/**
 * ==========================================================
 * Save Or Update Customer
 * ==========================================================
 */
function saveOrUpdateCustomer(orderData, orderId) {

  const customerSheet = getSheet(SHEETS.CUSTOMERS);

  const customer = orderData.customer;

  const rowNumber = findCustomerRowByMobile(
    customerSheet,
    customer.mobile
  );

  if (rowNumber === -1) {

    createCustomer(
      customerSheet,
      orderData,
      orderId
    );

  } else {

    updateExistingCustomer(
      customerSheet,
      rowNumber,
      orderData,
      orderId
    );

  }

}


/**
 * ==========================================================
 * Generate Customer ID
 * ==========================================================
 */
function generateCustomerID(customerSheet) {

  const lastRow = customerSheet.getLastRow();

  if (lastRow <= 1) {

    return "CUS00001";

  }

  const lastCustomerID = customerSheet
    .getRange(lastRow, 1)
    .getValue();

  const lastNumber = Number(
    String(lastCustomerID).replace("CUS", "")
  );

  const nextNumber = lastNumber + 1;

  return "CUS" + String(nextNumber).padStart(5, "0");

}


/**
 * ==========================================================
 * Find Customer Row By Mobile
 * Optimized using TextFinder
 * ==========================================================
 */
function findCustomerRowByMobile(customerSheet, mobile) {

  const lastRow = customerSheet.getLastRow();

  if (lastRow <= 1) {

    return -1;

  }

  const searchMobile = String(mobile).trim();

  if (!searchMobile) {

    return -1;

  }

  const match = customerSheet
    .getRange(2, 3, lastRow - 1, 1)
    .createTextFinder(searchMobile)
    .matchEntireCell(true)
    .findNext();

  if (!match) {

    return -1;

  }

  return match.getRow();

}


/**
 * ==========================================================
 * Create Customer
 * ==========================================================
 */
function createCustomer(customerSheet, orderData, orderId) {

  const customer = orderData.customer;

  const customerId = generateCustomerID(customerSheet);

  const parts = orderData.meta.orderDate.split("-");

  const orderDate = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
    12,
    0,
    0
  );

  const newRow = [[

    customerId,                                  // A Customer ID

    customer.customerName,                       // B Customer Name

    customer.mobile,                             // C Mobile

    customer.deliveryArea,                       // D Delivery Area

    customer.houseAddress,                       // E House Address

    orderDate,                                   // F First Order Date

    orderDate,                                   // G Last Order Date

    1,                                           // H Total Orders

    Number(orderData.totals.grandTotal) || 0,    // I Lifetime Spend

    orderId                                      // J Last Order ID

  ]];

  const nextRow = customerSheet.getLastRow() + 1;

  customerSheet
    .getRange(
      nextRow,
      1,
      1,
      newRow[0].length
    )
    .setValues(newRow);

}


/**
 * ==========================================================
 * Update Existing Customer
 * ==========================================================
 */
function updateExistingCustomer(
  customerSheet,
  rowNumber,
  orderData,
  orderId
) {

  const customer = orderData.customer;

  // Read existing customer data A:J
  const row = customerSheet
    .getRange(
      rowNumber,
      1,
      1,
      10
    )
    .getValues()[0];

  // Preserve First Order Date
  const firstOrderDate = row[5];

  // Build Last Order Date from business order date
  const parts = orderData.meta.orderDate.split("-");

  const lastOrderDate = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
    12,
    0,
    0
  );

  // Increment Total Orders
  const totalOrders =
    Number(row[7] || 0) + 1;

  // Current Order Total
  const grandTotal =
    Number(orderData.totals.grandTotal || 0);

  // Update Lifetime Spend
  const lifetimeSpend =
    Number(row[8] || 0) + grandTotal;

  // Update customer A:J in one batch write
  customerSheet
    .getRange(
      rowNumber,
      1,
      1,
      10
    )
    .setValues([[

      row[0],                    // A Customer ID

      customer.customerName,     // B Customer Name

      customer.mobile,           // C Mobile

      customer.deliveryArea,     // D Delivery Area

      customer.houseAddress,     // E House Address

      firstOrderDate,            // F First Order Date

      lastOrderDate,             // G Last Order Date

      totalOrders,               // H Total Orders

      lifetimeSpend,             // I Lifetime Spend

      orderId                    // J Last Order ID

    ]]);

}


/**
 * ==========================================================
 * Get Customer By Mobile
 * ==========================================================
 */
function getCustomerByMobileNumber(mobile) {

  const customerSheet = getSheet(SHEETS.CUSTOMERS);

  const rowNumber = findCustomerRowByMobile(
    customerSheet,
    mobile
  );

  if (rowNumber === -1) {

    return null;

  }

  const row = customerSheet
    .getRange(
      rowNumber,
      1,
      1,
      10
    )
    .getValues()[0];

  return {

    customerId: row[0],

    customerName: row[1],

    mobile: row[2],

    deliveryArea: row[3],

    houseAddress: row[4],

    firstOrderDate: row[5],

    lastOrderDate: row[6],

    totalOrders: row[7],

    lifetimeSpend: row[8],

    lastOrderId: row[9]

  };

}


/**
 * ==========================================================
 * Load All Customers for POS Cache
 * ==========================================================
 */
function getAllCustomers() {

  const sheet = getSheet(SHEETS.CUSTOMERS);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {

    return [];

  }

  const data = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      5
    )
    .getValues();

  return data.map(row => ({

    customerName: row[1],

    mobile: String(row[2]).trim(),

    deliveryArea: row[3],

    houseAddress: row[4]

  }));

}