/**
 * ==========================================
 * Get Invoice - PERFORMANCE OPTIMIZED
 * ==========================================
 */
function getInvoice(orderId) {

  try {

    // ==========================================
    // VALIDATE
    // ==========================================

    if (!orderId) {

      return {
        success: false,
        message: "Order ID is required."
      };

    }

    orderId =
      String(orderId)
        .trim()
        .toUpperCase();


    // ==========================================
    // LOAD SHEETS
    // ==========================================

    const ordersSheet =
      getSheet(SHEETS.ORDERS);

    const itemsSheet =
      getSheet(SHEETS.ORDER_ITEMS);


    // ==========================================
    // FIND ORDER
    //
    // IMPORTANT:
    // We search only Column A first.
    // We no longer load the complete Orders
    // sheet into memory.
    // ==========================================

    const orderLastRow =
      ordersSheet.getLastRow();

    if (orderLastRow <= 1) {

      return {
        success: false,
        message: "Orders sheet is empty."
      };

    }


    const orderIdRange =
      ordersSheet.getRange(
        2,
        1,
        orderLastRow - 1,
        1
      );


    const orderFinder =
      orderIdRange.createTextFinder(orderId)
        .matchEntireCell(true)
        .matchCase(false);


    const orderCell =
      orderFinder.findNext();


    if (!orderCell) {

      return {
        success: false,
        message: "Invoice not found."
      };

    }


    // ==========================================
    // READ ONLY THE MATCHING ORDER ROW
    // A:R = 18 columns
    // ==========================================

    const orderRowNumber =
      orderCell.getRow();


    const orderRow =
      ordersSheet
        .getRange(
          orderRowNumber,
          1,
          1,
          18
        )
        .getValues()[0];


    // ==========================================
    // FIND ORDER ITEMS
    // ==========================================

    const invoiceItems = [];

    const itemLastRow =
      itemsSheet.getLastRow();


    if (itemLastRow > 1) {

      const itemIdRange =
        itemsSheet.getRange(
          2,
          1,
          itemLastRow - 1,
          1
        );


      const itemFinder =
        itemIdRange.createTextFinder(orderId)
          .matchEntireCell(true)
          .matchCase(false);


      const matches =
        itemFinder.findAll();


      if (matches && matches.length) {

        // Usually an invoice has only a few items.
        // Read only those matching rows instead of
        // loading the complete OrderItems sheet.

        for (
          let i = 0;
          i < matches.length;
          i++
        ) {

          const rowNumber =
            matches[i].getRow();


          const row =
            itemsSheet
              .getRange(
                rowNumber,
                1,
                1,
                5
              )
              .getValues()[0];


          invoiceItems.push({

            itemName:
              String(row[1] || ""),

            qty:
              Number(row[2]) || 0,

            price:
              Number(row[3]) || 0,

            total:
              Number(row[4]) || 0

          });

        }

      }

    }


    // ==========================================
    // FORMAT DATES
    // ==========================================

    const timezone =
      Session.getScriptTimeZone();


    const orderDate =
      orderRow[1]
        ? Utilities.formatDate(
            new Date(orderRow[1]),
            timezone,
            "dd/MM/yyyy"
          )
        : "";


    const createdAt =
      orderRow[13]
        ? Utilities.formatDate(
            new Date(orderRow[13]),
            timezone,
            "dd/MM/yyyy HH:mm:ss"
          )
        : "";


    // ==========================================
    // RETURN
    // ==========================================

    return {

      success: true,

      order: {

        orderId:
          String(orderRow[0] || ""),

        orderDate:
          orderDate,

        customerName:
          String(orderRow[2] || ""),

        mobile:
          String(orderRow[3] || ""),

        deliveryArea:
          String(orderRow[4] || ""),

        houseAddress:
          String(orderRow[5] || ""),

        deliverySlot:
          String(orderRow[6] || ""),

        orderType:
          String(orderRow[7] || ""),

        paymentMode:
          String(orderRow[8] || ""),

        paymentStatus:
          String(orderRow[9] || ""),

        grandTotal:
          Number(orderRow[10]) || 0,

        status:
          String(orderRow[11] || ""),

        specialInstructions:
          String(orderRow[12] || ""),

        createdAt:
          createdAt,

        deliveryCharge:
          Number(orderRow[14]) || 0,

        discount:
          Number(orderRow[15]) || 0,

        addons:
          String(orderRow[16] || ""),

        addonTotal:
          Number(orderRow[17]) || 0

      },

      items:
        invoiceItems

    };


  } catch (error) {

    console.error(
      "getInvoice Error:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

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










/**
 * ==========================================
 * Get Smoky Swaad Logo
 * ==========================================
 *
 * Loads the Smoky Swaad logo from Google Drive
 * and returns it as a Base64 data URL.
 *
 * Used by Invoice.html and PDF generation.
 * ==========================================
 */
function getSmokySwaadLogo() {

  try {

    const LOGO_FILE_ID =
      "19nlgGbhvi8vU3DbE52F8DNlwotcHoPuz";

    const file =
      DriveApp.getFileById(LOGO_FILE_ID);

    const blob =
      file.getBlob();

    const contentType =
      blob.getContentType() || "image/png";

    const base64 =
      Utilities.base64Encode(
        blob.getBytes()
      );

    return {
      success: true,

      dataUrl:
        "data:" +
        contentType +
        ";base64," +
        base64
    };

  } catch (error) {

    Logger.log(
      "Logo Load Error: " +
      error.toString()
    );

    return {
      success: false,
      dataUrl: "",
      message: error.toString()
    };

  }

}