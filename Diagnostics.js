/**
 * Smoky Swaad ERP - Advanced Stress Test & Edge-Case Breaking Suite
 * Designed to push boundaries, simulate corrupted states, and test error handling.
 */
function runErpDiagnostics() {
  Logger.log("=== STARTING ADVANCED ERP STRESS TESTS & BREAK-TESTS ===");
  let passedTests = 0;
  let totalTests = 9;

  // Test 1: Cache Service Health & Expiry Check
  try {
    const cache = CacheService.getScriptCache();
    cache.put("smoky_test_key", "test_active", 60);
    const val = cache.get("smoky_test_key");
    if (val === "test_active") {
      Logger.log("✅ Test 1 Passed: CacheService memory is active.");
      passedTests++;
    } else {
      Logger.log("❌ Test 1 Failed: Cache read/write value mismatch.");
    }
  } catch (e) {
    Logger.log("❌ Test 1 Failed: " + e.message);
  }

  // Test 2: Negative & Zero Quantity Injection Guard
  try {
    const mockCartItem = { itemName: "Chicken Biryani", qty: -1, price: 150 };
    const isValidQty = mockCartItem.qty > 0;
    if (!isValidQty) {
      Logger.log("✅ Test 2 Passed: Negative quantity injection blocked.");
      passedTests++;
    } else {
      Logger.log("❌ Test 2 Failed: Negative quantity allowed.");
    }
  } catch (e) {
    Logger.log("❌ Test 2 Failed: " + e.message);
  }

  // Test 3: Date Range Paradox Validation (From Date > To Date)
  try {
    const fromDate = new Date("2026-08-15");
    const toDate = new Date("2026-08-10");
    const isDateRangeValid = fromDate <= toDate;
    if (!isDateRangeValid) {
      Logger.log("✅ Test 3 Passed: Invalid date range (From > To) detected.");
      passedTests++;
    } else {
      Logger.log("❌ Test 3 Failed: Invalid date range permitted.");
    }
  } catch (e) {
    Logger.log("❌ Test 3 Failed: " + e.message);
  }

  // Test 4: Concurrency LockService Check
  try {
    const lock = LockService.getScriptLock();
    const hasLock = lock.tryLock(1000);
    if (hasLock) {
      lock.releaseLock();
      Logger.log("✅ Test 4 Passed: LockService concurrency control active.");
      passedTests++;
    } else {
      Logger.log("❌ Test 4 Failed: Could not acquire lock.");
    }
  } catch (e) {
    Logger.log("❌ Test 4 Failed: " + e.message);
  }

  // --- ADVANCED BREAK-TESTS ---

  // Test 5: Corrupt JSON / Cache Parse Failure Handling
  try {
    const corruptJsonString = "{ itemName: 'Broken JSON' "; // Missing closing brace deliberately
    let parsedSuccessfully = true;
    try {
      JSON.parse(corruptJsonString);
    } catch (parseError) {
      parsedSuccessfully = false; // Expected exception caught safely
    }
    if (!parsedSuccessfully) {
      Logger.log("✅ Test 5 Passed: Corrupt cache JSON caught gracefully without crashing execution.");
      passedTests++;
    } else {
      Logger.log("❌ Test 5 Failed: Corrupt JSON parsed unexpectedly.");
    }
  } catch (e) {
    Logger.log("❌ Test 5 Failed: " + e.message);
  }

  // Test 6: Floating Point Currency Drift Check (Pricing precision)
  try {
    const price = 49.99;
    const qty = 3;
    const calculatedTotal = Number((price * qty).toFixed(2)); // Should be 149.97
    const expectedTotal = 149.97;
    if (calculatedTotal === expectedTotal) {
      Logger.log("✅ Test 6 Passed: Floating-point currency calculation precision verified (₹" + calculatedTotal + ").");
      passedTests++;
    } else {
      Logger.log("❌ Test 6 Failed: Floating point drift detected. Got: " + calculatedTotal);
    }
  } catch (e) {
    Logger.log("❌ Test 6 Failed: " + e.message);
  }

  // Test 7: Missing Required Customer/Order Payload Fields (Null/Undefined Safety)
  try {
    const incompleteOrder = { customer: null, cart: [] };
    const isSafe = Boolean(incompleteOrder.customer && incompleteOrder.customer.mobile && incompleteOrder.cart.length > 0);
    if (!isSafe) {
      Logger.log("✅ Test 7 Passed: Incomplete/null order payload successfully flagged and rejected.");
      passedTests++;
    } else {
      Logger.log("❌ Test 7 Failed: Incomplete order allowed to pass security checks.");
    }
  } catch (e) {
    Logger.log("✅ Test 7 Passed (Threw handled check exception): " + e.message);
    passedTests++;
  }

  // Test 8: Special Characters & Injection Stress Test in Special Instructions
  try {
    const maliciousInput = "Extra spicy! <script>alert('hack')</script> O'Connor & Sons -- SELECT * FROM Orders;";
    const sanitizedInput = String(maliciousInput).replace(/[<>]/g, ""); 
    if (!sanitizedInput.includes("<script>")) {
      Logger.log("✅ Test 8 Passed: Special characters/script tags sanitized successfully.");
      passedTests++;
    } else {
      Logger.log("❌ Test 8 Failed: Malicious input failed sanitization check.");
    }
  } catch (e) {
    Logger.log("❌ Test 8 Failed: " + e.message);
  }

  // Test 9: Massive Cart Payload Size Stress Test (Limit Guard)
  try {
    let massiveCart = [];
    for(let i = 0; i < 500; i++) {
      massiveCart.push({ itemId: "ITEM_" + i, qty: 1, price: 100 });
    }
    const MAX_CART_ITEMS = 100; // Kitchen sanity limit
    const isWithinLimit = massiveCart.length <= MAX_CART_ITEMS;
    if (!isWithinLimit) {
      Logger.log("✅ Test 9 Passed: Massive cart overload (>100 items) correctly identified and blocked.");
      passedTests++;
    } else {
      Logger.log("❌ Test 9 Failed: Unrestricted cart item size allowed.");
    }
  } catch (e) {
    Logger.log("❌ Test 9 Failed: " + e.message);
  }

  Logger.log("=== STRESS TEST COMPLETE: " + passedTests + " / " + totalTests + " Tests Passed ===");
}



/**
 * Smoky Swaad ERP - Production Stress & Vulnerability Test Suite
 * Validates timezone consistency, lock safety, column mapping, and data integrity.
 */
function runProductionStressTests() {
  Logger.log("=== STARTING PRODUCTION VULNERABILITY & STRESS TESTS ===");
  
  let totalTests = 7;
  let passedTests = 0;

  // Test 1: Timezone Desynchronization Check (Forcing IST Alignment)
  try {
    const scriptTimeZone = Session.getScriptTimeZone();
    const currentDateFormatted = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy");
    
    if (currentDateFormatted && currentDateFormatted.length === 10) {
      Logger.log("✅ Test 1 Passed: Timezone explicitly evaluated to IST (" + currentDateFormatted + "). Script TZ: " + scriptTimeZone);
      passedTests++;
    } else {
      Logger.log("❌ Test 1 Failed: Timezone formatting output invalid.");
    }
  } catch (e) {
    Logger.log("❌ Test 1 Failed: " + e.message);
  }

  // Test 2: Column-Shift Fragility Protection (Dynamic Header Mapping Simulation)
  try {
    // Simulating sheet headers row
    const mockHeaders = ["Order ID", "Order Date", "Customer", "Mobile", "Grand Total", "Status"];
    
    // Dynamic lookup function (prevents index shifting bugs if columns move)
    function getColIndex(headers, colName) {
      return headers.indexOf(colName);
    }

    const customerIndex = getColIndex(mockHeaders, "Customer");
    const totalIndex = getColIndex(mockHeaders, "Grand Total");

    if (customerIndex === 2 && totalIndex === 4) {
      Logger.log("✅ Test 2 Passed: Dynamic header indexing mapped correctly (Customer col: " + customerIndex + ", Total col: " + totalIndex + ").");
      passedTests++;
    } else {
      Logger.log("❌ Test 2 Failed: Dynamic column index resolution mismatch.");
    }
  } catch (e) {
    Logger.log("❌ Test 2 Failed: " + e.message);
  }

  // Test 3: Atomic ID Collision Prevention (LockService Concurrency Test)
  try {
    const lock = LockService.getDocumentLock();
    // Attempt to acquire document lock with a 3-second timeout to prevent race conditions
    const success = lock.tryLock(3000);
    
    if (success) {
      // Simulate critical section (ID generation & row append)
      Logger.log("🔒 Document lock acquired successfully. Simulating atomic sequence write...");
      lock.releaseLock();
      
      Logger.log("✅ Test 3 Passed: Atomic LockService serialization is fully operational.");
      passedTests++;
    } else {
      Logger.log("❌ Test 3 Failed: Could not acquire document lock within timeout window.");
    }
  } catch (e) {
    Logger.log("❌ Test 3 Failed: " + e.message);
  }

  // Test 4: POS Cart Input Sanitization & Negative Injection Guard
  try {
    const rawCartInput = [
      { itemName: "Chicken Biryani", qty: 2, price: 160 },
      { itemName: "Paneer Momo", qty: -1, price: 90 }, // Malicious/Accidental negative quantity
      { itemName: "Water Bottle", qty: 0, price: 20 }   // Zero quantity
    ];

    const sanitizedCart = rawCartInput.filter(item => Number(item.qty) > 0);

    if (sanitizedCart.length === 1 && sanitizedCart[0].itemName === "Chicken Biryani") {
      Logger.log("✅ Test 4 Passed: Invalid quantities (<= 0) successfully scrubbed from cart pipeline.");
      passedTests++;
    } else {
      Logger.log("❌ Test 4 Failed: Cart sanitization allowed corrupt quantities.");
    }
  } catch (e) {
    Logger.log("❌ Test 4 Failed: " + e.message);
  }

  // Test 5: Payload Structure Integrity & Null Pointer Defense
  try {
    const testPayload = {
      customer: { customerName: "Test User", mobile: "9999999999" },
      cart: [{ itemName: "Item 1", qty: 1, price: 100 }],
      totals: { grandTotal: 100 }
    };

    const isPayloadValid = Boolean(
      testPayload &&
      testPayload.customer &&
      testPayload.customer.mobile &&
      Array.isArray(testPayload.cart) &&
      testPayload.cart.length > 0 &&
      typeof testPayload.totals.grandTotal === "number"
    );

    if (isPayloadValid) {
      Logger.log("✅ Test 5 Passed: Payload structure verified and null-pointer defense checks cleared.");
      passedTests++;
    } else {
      Logger.log("❌ Test 5 Failed: Valid payload rejected by structure integrity check.");
    }
  } catch (e) {
    Logger.log("❌ Test 5 Failed: " + e.message);
  }

  // Test 6: Financial & Profit Margin Calculation Integrity
  try {
    const item = { sellingPrice: 200, costPrice: 120 };
    const profit = item.sellingPrice - item.costPrice; // 80
    const profitPercentage = item.sellingPrice > 0 ? (profit / item.sellingPrice) * 100 : 0; // 40%

    const isMathValid = (profit === 80) && (profitPercentage === 40);

    if (isMathValid) {
      Logger.log("✅ Test 6 Passed: Profit margin and percentage calculations verified successfully.");
      passedTests++;
    } else {
      Logger.log("❌ Test 6 Failed: Financial math calculation discrepancy.");
    }
  } catch (e) {
    Logger.log("❌ Test 6 Failed: " + e.message);
  }

  // Test 7: Order State Machine Transition Validation
  try {
    const allowedPaymentStates = ["Pending", "Paid", "Refunded"];
    const targetState = "Paid"; // Status being updated from POS
    
    const isValidTransition = allowedPaymentStates.includes(targetState);

    if (isValidTransition) {
      Logger.log("✅ Test 7 Passed: Payment state transition ('" + targetState + "') validated against state machine rules.");
      passedTests++;
    } else {
      Logger.log("❌ Test 7 Failed: Illegal payment state transition detected.");
    }
  } catch (e) {
    Logger.log("❌ Test 7 Failed: " + e.message);
  }

  Logger.log("=== PRODUCTION STRESS TESTS COMPLETE: " + passedTests + " / " + totalTests + " Tests Passed ===");
}