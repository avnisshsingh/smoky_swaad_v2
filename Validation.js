/**
 * ===========================================
 * SMOKY SWAAD ERP
 * Validation Module
 * ===========================================
 */

/**
 * Reset all highlighted cells
 */
function resetValidationColors(pos) {

  const INPUT_COLOR = "#d9eaf7";

  // Customer Details
  pos.getRange("B2:B11").setBackground(INPUT_COLOR);

  // Order Items
  pos.getRange(
    POS.ITEM_START_ROW,
    POS.ITEM_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    3
  ).setBackground(INPUT_COLOR);

}


/**
 * Highlight a cell
 */
function highlight(pos, cell) {

  pos.getRange(cell).setBackground("#f4cccc");

}


/**
 * Validate POS Screen
 */
function validateOrder() {

  const pos = getSheet(SHEETS.POS);

  let errors = [];

  // Reset previous validation colors
  resetValidationColors(pos);

  //---------------------------------------
  // Read Customer Details (ONE API CALL)
  //---------------------------------------

  const formData = pos.getRange("B2:B11").getValues();

  const customerName   = String(formData[0][0]).trim();
  const mobile         = String(formData[1][0]).trim();
  const orderDate      = formData[2][0];
  const deliveryArea   = String(formData[3][0]).trim();
  const houseAddress   = String(formData[4][0]).trim();
  const deliverySlot   = String(formData[5][0]).trim();
  const orderType      = String(formData[6][0]).trim();
  const paymentMode    = String(formData[7][0]).trim();
  const paymentStatus  = String(formData[8][0]).trim();

  //---------------------------------------
  // Mandatory Fields
  //---------------------------------------

  if (!customerName) {

    highlight(pos, POS.CUSTOMER_NAME);
    errors.push("Customer Name is required.");

  }

  if (!mobile) {

    highlight(pos, POS.MOBILE);
    errors.push("Mobile Number is required.");

  }

  if (!orderDate) {

    highlight(pos, POS.ORDER_DATE);
    errors.push("Order Date is required.");

  }

  if (!deliveryArea) {

    highlight(pos, POS.DELIVERY_AREA);
    errors.push("Delivery Area is required.");

  }

  // House Address is OPTIONAL

  if (!deliverySlot) {

    highlight(pos, POS.DELIVERY_SLOT);
    errors.push("Delivery Slot is required.");

  }

  if (!orderType) {

    highlight(pos, POS.ORDER_TYPE);
    errors.push("Order Type is required.");

  }

  if (!paymentMode) {

    highlight(pos, POS.PAYMENT_MODE);
    errors.push("Payment Mode is required.");

  }

  if (!paymentStatus) {

    highlight(pos, POS.PAYMENT_STATUS);
    errors.push("Payment Status is required.");

  }

  //---------------------------------------
  // Mobile Validation
  //---------------------------------------

  if (mobile && !/^[6-9]\d{9}$/.test(mobile)) {

    highlight(pos, POS.MOBILE);

    errors.push(
      "Mobile Number must be a valid 10 digit Indian mobile number."
    );

  }

  //---------------------------------------
  // Read Order Items (ONE API CALL)
  //---------------------------------------

  const items = pos.getRange(
    POS.ITEM_START_ROW,
    POS.ITEM_COL,
    POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
    4
  ).getValues();

  let itemFound = false;

  items.forEach((row, index) => {

    const item = String(row[0]).trim();
    const qty  = Number(row[2]);

    if (item) {

      itemFound = true;

      if (qty <= 0) {

        pos.getRange(
          POS.ITEM_START_ROW + index,
          POS.QTY_COL
        ).setBackground("#f4cccc");

        errors.push("Quantity missing for '" + item + "'.");

      }

    }

  });

  if (!itemFound) {

    pos.getRange(
      POS.ITEM_START_ROW,
      POS.ITEM_COL,
      POS.ITEM_END_ROW - POS.ITEM_START_ROW + 1,
      1
    ).setBackground("#f4cccc");

    errors.push("Please select at least one Menu Item.");

  }

  //---------------------------------------
  // Grand Total
  //---------------------------------------

  const grandTotal = Number(
    pos.getRange(POS.GRAND_TOTAL).getValue()
  );

  if (grandTotal <= 0) {

    highlight(pos, POS.GRAND_TOTAL);

    errors.push("Grand Total must be greater than zero.");

  }

  //---------------------------------------
  // Result
  //---------------------------------------

  return {

    valid: errors.length === 0,

    errors: errors

  };

}


/**
 * Test Validation
 */
function testValidation() {

  const result = validateOrder();

  if (result.valid) {

    SpreadsheetApp.getUi().alert(
      "✅ Validation Passed.\n\nOrder is ready to save."
    );

  } else {

    SpreadsheetApp.getUi().alert(
      "Please correct the following:\n\n• " +
      result.errors.join("\n• ")
    );

  }

}