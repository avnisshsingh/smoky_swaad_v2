/**
 * ==========================================
 * Get Invoice
 * ==========================================
 */
function getInvoice(orderId) {

  Logger.log("=====================================");
  Logger.log("getInvoice() Started");
  Logger.log("Input Order ID : " + orderId);

  try {

    // -------------------------------
    // Validate Input
    // -------------------------------
    if (!orderId) {

      Logger.log("Order ID is blank.");

      return {
        success: false,
        message: "Order ID is required."
      };

    }

    orderId = String(orderId).trim().toUpperCase();

    Logger.log("Searching for Order ID : " + orderId);

    // -------------------------------
    // Get Sheets
    // -------------------------------
    Logger.log("Opening Orders sheet...");

    const ordersSheet = getSheet(SHEETS.ORDERS);

    Logger.log("Orders sheet loaded.");

    Logger.log("Opening OrderItems sheet...");

    const itemsSheet = getSheet(SHEETS.ORDER_ITEMS);

    Logger.log("OrderItems sheet loaded.");

    // -------------------------------
    // Read Data
    // -------------------------------
    const orders = ordersSheet.getDataRange().getValues();
    const items = itemsSheet.getDataRange().getValues();

    Logger.log("Orders Rows : " + orders.length);
    Logger.log("Order Items Rows : " + items.length);

    if (orders.length <= 1) {

      Logger.log("Orders sheet is empty.");

      return {
        success: false,
        message: "Orders sheet is empty."
      };

    }

    orders.shift();
    items.shift();

    // -------------------------------
    // Find Order
    // -------------------------------
    Logger.log("Searching order...");

    const orderRow = orders.find(function (row) {

      return String(row[0]).trim().toUpperCase() === orderId;

    });

    if (!orderRow) {

      Logger.log("Order NOT FOUND.");

      return {

        success: false,

        message: "Invoice not found."

      };

    }

    Logger.log("Order Found.");

    // -------------------------------
    // Find Order Items
    // -------------------------------
    Logger.log("Searching invoice items...");

    const invoiceItems = items
      .filter(function (row) {

        return String(row[0]).trim().toUpperCase() === orderId;

      })
      .map(function (row) {

        return {

          itemName: row[1],
          qty: Number(row[2]),
          price: Number(row[3]),
          total: Number(row[4])

        };

      });

    Logger.log("Items Found : " + invoiceItems.length);

    // -------------------------------
    // Success
    // -------------------------------
// -------------------------------
// Success
// -------------------------------
const response = {

  success: true,

  order: {

    orderId: String(orderRow[0] || ""),

    orderDate: orderRow[1]
      ? Utilities.formatDate(
          new Date(orderRow[1]),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        )
      : "",

    customerName: String(orderRow[2] || ""),
    mobile: String(orderRow[3] || ""),
    deliveryArea: String(orderRow[4] || ""),
    houseAddress: String(orderRow[5] || ""),
    deliverySlot: String(orderRow[6] || ""),
    orderType: String(orderRow[7] || ""),
    paymentMode: String(orderRow[8] || ""),
    paymentStatus: String(orderRow[9] || ""),

    grandTotal: Number(orderRow[10]) || 0,

    status: String(orderRow[11] || ""),

    specialInstructions: String(orderRow[12] || ""),

    createdAt: orderRow[13]
      ? Utilities.formatDate(
          new Date(orderRow[13]),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm:ss"
        )
      : "",

    deliveryCharge: Number(orderRow[14]) || 0,

    discount: Number(orderRow[15]) || 0

  },

  items: invoiceItems

};

Logger.log(JSON.stringify(response));

return response;

  } catch (err) {

    Logger.log("=====================================");
    Logger.log("ERROR OCCURRED");
    Logger.log(err.toString());

    if (err.stack) {

      Logger.log(err.stack);

    }

    return {

      success: false,

      message: err.toString()

    };

  }

}



/**
 * ==========================================
 * Generate Invoice PDF
 * Step 2.2
 * ==========================================
 */
function generateInvoicePdf(orderId) {

  try {

    Logger.log("==================================");
    Logger.log("Generating PDF");
    Logger.log("Order ID : " + orderId);

    // ------------------------------------
    // Fetch Invoice
    // ------------------------------------

    const invoice = getInvoice(orderId);

    if (!invoice.success) {

      return invoice;

    }

    // ------------------------------------
    // Load HTML Template
    // ------------------------------------

    const template = HtmlService.createTemplateFromFile("InvoicePdf");

    template.order = invoice.order;
    template.items = invoice.items;

    // ------------------------------------
    // Convert HTML
    // ------------------------------------

    const html = template
      .evaluate()
      .getContent();

    Logger.log("HTML Generated");

    // ------------------------------------
    // Generate PDF Blob
    // ------------------------------------

    const pdfBlob = Utilities.newBlob(html, "text/html")
      .getAs(MimeType.PDF);

    pdfBlob.setName(
      orderId + ".pdf"
    );

    Logger.log("PDF Blob Created");

    Logger.log("Name : " + pdfBlob.getName());

    Logger.log("Size : " + pdfBlob.getBytes().length);

    return {

      success: true,

      message: "PDF Blob Generated Successfully.",

      blobName: pdfBlob.getName(),

      blobSize: pdfBlob.getBytes().length

    };

  }

  catch(err){

    Logger.log(err);

    return {

      success:false,

      message:err.toString()

    };

  }

}