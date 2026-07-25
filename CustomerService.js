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
function saveOrUpdateCustomer(orderData) {

  const customerSheet = getSheet(SHEETS.CUSTOMERS);

  const customer = orderData.customer;

  const rowNumber = findCustomerRowByMobile(
    customerSheet,
    customer.mobile
  );

  if (rowNumber === -1) {

    createCustomer(
      customerSheet,
      orderData
    );

  } else {

    updateCustomer(
      customerSheet,
      rowNumber,
      orderData
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
 * ==========================================================
 */
function findCustomerRowByMobile(customerSheet, mobile) {

  const lastRow = customerSheet.getLastRow();

  if (lastRow <= 1) {

    return -1;

  }

  const mobiles = customerSheet
    .getRange(2, 3, lastRow - 1, 1)
    .getValues();

  for (let i = 0; i < mobiles.length; i++) {

    if (String(mobiles[i][0]).trim() === String(mobile).trim()) {

      return i + 2;

    }

  }

  return -1;

}


/**
 * ==========================================================
 * Create Customer
 * ==========================================================
 */
function createCustomer(customerSheet, orderData) {

  const customer = orderData.customer;

  const customerId = generateCustomerID(customerSheet);

  const parts = orderData.meta.orderDate.split("-");

const orderDate = new Date(
    Number(parts[0]),      // Year
    Number(parts[1]) - 1,  // Month (JavaScript months start from 0)
    Number(parts[2]),      // Day
    12, 0, 0               // Noon (avoids timezone shifting)
);

  customerSheet.appendRow([

    customerId,

    customer.customerName,

    customer.mobile,

    customer.deliveryArea,

    customer.houseAddress,

    orderDate,

    orderDate,

    1,

    orderData.totals.grandTotal

  ]);

}


/**
 * ==========================================================
 * Update Customer
 * ==========================================================
 */
function updateCustomer(customerSheet, rowNumber, orderData) {

  const customer = orderData.customer;

  // Read the existing customer row
  const row = customerSheet
      .getRange(rowNumber, 1, 1, 9)
      .getValues()[0];

  // Preserve first order date
  const firstOrderDate = row[5];

  // Build the new last order date from the business date
  const parts = orderData.meta.orderDate.split("-");

  const lastOrderDate = new Date(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2]),
      12, 0, 0
  );

  const totalOrders = Number(row[7]) + 1;

  const lifetimeSpend =
      Number(row[8]) + Number(orderData.totals.grandTotal);

  customerSheet
      .getRange(rowNumber, 1, 1, 9)
      .setValues([[
          row[0],                    // Customer ID
          customer.customerName,
          customer.mobile,
          customer.deliveryArea,
          customer.houseAddress,
          firstOrderDate,
          lastOrderDate,
          totalOrders,
          lifetimeSpend
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
    .getRange(rowNumber, 1, 1, 9)
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

    lifetimeSpend: row[8]

  };

}