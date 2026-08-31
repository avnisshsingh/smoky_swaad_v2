/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * CUSTOMER SERVICE
 * ==========================================================
 */

/**
 * ==========================================
 * Get Customer Cache (Slimmed Payload)
 * ==========================================
 */
function getCustomerCacheOptimized() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_customers_slim_cache_v1";

   const cached = cache.get(cacheKey);
   if (cached) {
      try {
         return JSON.parse(cached);
      } catch (e) {}
   }

   const sheet = getSheet("Customers"); // Adjust constant if named differently in your Constants.js
   const lastRow = sheet.getLastRow();
   const customers = [];

   if (lastRow >= 2) {
      // Read only essential columns: ID(1), Name(2), Mobile(3), Address(4), Area(5)
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      
      data.forEach(function(row) {
         const mobile = String(row[2] || "").trim();
         if (mobile) {
            customers.push({
               customerId: String(row[0] || "").trim(),
               customerName: String(row[1] || "").trim(),
               mobile: mobile,
               houseAddress: String(row[3] || "").trim(),
               deliveryArea: String(row[4] || "").trim()
            });
         }
      });
   }

   try {
      cache.put(cacheKey, JSON.stringify(customers), 3600); // Cache for 1 hour
   } catch (e) {}

   return customers;
}

function clearCustomerCache() {
   try {
      CacheService.getScriptCache().remove("smoky_swaadv2_customers_slim_cache_v1");
   } catch (e) {}
}
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
// CLEAR THE CUSTOMER CACHE SO SUBSEQUENT POS LOOKUPS SEE THE UPDATE
  clearCustomersCache();
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


  // ========================================================
  // READ EXISTING CUSTOMER A:J
  // ========================================================

  const row = customerSheet
    .getRange(
      rowNumber,
      1,
      1,
      10
    )
    .getValues()[0];


  // ========================================================
  // BUILD CURRENT ORDER DATE
  // ========================================================

  const parts =
    String(orderData.meta.orderDate)
      .split("-");


  if (parts.length !== 3) {

    throw new Error(
      "Invalid Order Date."
    );

  }


  const currentOrderDate = new Date(

    Number(parts[0]),

    Number(parts[1]) - 1,

    Number(parts[2]),

    12,
    0,
    0

  );


  // ========================================================
  // FIRST ORDER DATE
  // ========================================================
  //
  // Normal existing customer:
  // Preserve original First Order Date.
  //
  // Customer whose entire order history was deleted:
  // First Order Date is blank and Total Orders = 0.
  // Therefore this new order becomes their new First Order.
  // ========================================================

  const existingTotalOrders =
    Number(row[7] || 0);


  let firstOrderDate =
    row[5];


  if (
    existingTotalOrders <= 0 ||
    !firstOrderDate
  ) {

    firstOrderDate =
      currentOrderDate;

  }


  // ========================================================
  // LAST ORDER DATE
  // ========================================================

  const lastOrderDate =
    currentOrderDate;


  // ========================================================
  // TOTAL ORDERS
  // ========================================================

  const totalOrders =
    existingTotalOrders + 1;


  // ========================================================
  // CURRENT ORDER GRAND TOTAL
  // ========================================================

  const grandTotal =
    Number(
      orderData.totals.grandTotal || 0
    );


  // ========================================================
  // LIFETIME SPEND
  // ========================================================

  const existingLifetimeSpend =
    Number(row[8] || 0);


  const lifetimeSpend =
    existingLifetimeSpend +
    grandTotal;


  // ========================================================
  // UPDATE CUSTOMER A:J
  // ========================================================

  customerSheet
    .getRange(
      rowNumber,
      1,
      1,
      10
    )
    .setValues([[

      // A Customer ID
      row[0],

      // B Customer Name
      customer.customerName,

      // C Mobile Number
      customer.mobile,

      // D Delivery Area
      customer.deliveryArea,

      // E House Address
      customer.houseAddress,

      // F First Order Date
      firstOrderDate,

      // G Last Order Date
      lastOrderDate,

      // H Total Orders
      totalOrders,

      // I Lifetime Spend
      lifetimeSpend,

      // J Last Order ID
      orderId

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
 * ==========================================
 * Load All Customers for POS Cache (Optimized with Server-Side Caching)
 * ==========================================
 */
function getAllCustomers() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_all_customers_cache_v1";

   // 1. Check if cached data exists in server memory
   const cachedData = cache.get(cacheKey);
   if (cachedData) {
      try {
         return JSON.parse(cachedData);
      } catch (e) {
         // Fallback to sheet reading if JSON parse fails
      }
   }

   // 2. Your exact original sheet fetching & mapping logic
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

   const customers = data.map(row => ({
      customerName: row[1],
      mobile: String(row[2]).trim(),
      deliveryArea: row[3],
      houseAddress: row[4]
   }));

   // 3. Store the result in cache for 6 hours (21600 seconds)
   try {
      cache.put(cacheKey, JSON.stringify(customers), 21600);
   } catch (e) {
      console.warn("Failed to write customers cache:", e);
   }

   return customers;
}

/**
 * ==========================================
 * Clear Customers Cache
 * ==========================================
 */
function clearCustomersCache() {
   try {
      const cache = CacheService.getScriptCache();
      cache.remove("smoky_swaad_all_customers_cache_v1");
   } catch (e) {
      console.warn("Failed to clear customers cache:", e);
   }
}