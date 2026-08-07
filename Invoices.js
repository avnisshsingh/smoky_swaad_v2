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










/**
 * ==========================================================
 * BULK INVOICE - SEARCH ORDERS
 * ==========================================================
 *
 * Completely independent from getInvoice().
 *
 * Supported filters:
 * - From Date
 * - To Date
 * - Mobile Number
 * - Order ID
 * - Payment Status
 *
 * Default behaviour:
 * If From Date and To Date are both blank,
 * orders from the CURRENT MONTH are returned.
 *
 * Orders Sheet:
 * A = Order ID
 * B = Order Date
 * C = Customer Name
 * D = Mobile
 * J = Payment Status
 * K = Grand Total
 * L = Order Status
 * ==========================================================
 */
function searchBulkInvoiceOrdersFromWeb(filters) {

  try {

    // ======================================================
    // NORMALIZE INPUT
    // ======================================================

    filters = filters || {};


    const fromDateText =
      String(
        filters.fromDate || ""
      ).trim();


    const toDateText =
      String(
        filters.toDate || ""
      ).trim();


    const mobileFilter =
      String(
        filters.mobile || ""
      )
      .replace(/\D/g, "")
      .trim();


    const orderIdFilter =
      String(
        filters.orderId || ""
      )
      .trim()
      .toUpperCase();


    const paymentStatusFilter =
      String(
        filters.paymentStatus || ""
      )
      .trim()
      .toLowerCase();


    // ======================================================
    // VALIDATE PAYMENT STATUS
    // ======================================================

    if (
      paymentStatusFilter &&
      paymentStatusFilter !== "paid" &&
      paymentStatusFilter !== "pending"
    ) {

      return {

        success: false,

        message:
          "Invalid payment status."

      };

    }


    // ======================================================
    // VALIDATE DATE INPUT
    // ======================================================

    if (
      (fromDateText && !toDateText) ||
      (!fromDateText && toDateText)
    ) {

      return {

        success: false,

        message:
          "Please select both From Date and To Date."

      };

    }


    // ======================================================
    // CREATE DATE RANGE
    // ======================================================

    let fromDate;
    let toDate;
    let usingCurrentMonth = false;


    // ------------------------------------------------------
    // User entered a date range
    // ------------------------------------------------------

    if (
      fromDateText &&
      toDateText
    ) {

      fromDate =
        new Date(
          fromDateText + "T00:00:00"
        );


      toDate =
        new Date(
          toDateText + "T23:59:59.999"
        );


      if (
        isNaN(fromDate.getTime()) ||
        isNaN(toDate.getTime())
      ) {

        return {

          success: false,

          message:
            "Invalid date range."

        };

      }


      if (
        fromDate.getTime() >
        toDate.getTime()
      ) {

        return {

          success: false,

          message:
            "From Date cannot be after To Date."

        };

      }

    }

    // ------------------------------------------------------
    // No dates selected:
    // automatically use current month
    // ------------------------------------------------------

    else {

      const today =
        new Date();


      fromDate =
        new Date(
          today.getFullYear(),
          today.getMonth(),
          1,
          0,
          0,
          0,
          0
        );


      toDate =
        new Date(
          today.getFullYear(),
          today.getMonth() + 1,
          0,
          23,
          59,
          59,
          999
        );


      usingCurrentMonth = true;

    }


    // ======================================================
    // LOAD ORDERS SHEET
    // ======================================================

    const ordersSheet =
      getSheet(
        SHEETS.ORDERS
      );


    const lastRow =
      ordersSheet.getLastRow();


    if (
      lastRow <= 1
    ) {

      return {

        success: true,

        orders: [],

        count: 0,

        totalAmount: 0,

        usingCurrentMonth:
          usingCurrentMonth,

        message:
          "No orders found."

      };

    }


    // ======================================================
    // READ ORDERS
    //
    // A:R = 18 columns
    //
    // For bulk searching, one bulk read is intentional.
    // Multiple TextFinder calls would be slower when several
    // optional filters must be evaluated together.
    // ======================================================

    const rows =
      ordersSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          18
        )
        .getValues();


    // ======================================================
    // SEARCH
    // ======================================================

    const results = [];

    let totalAmount = 0;


    const timezone =
      Session.getScriptTimeZone();


    for (
      let i = 0;
      i < rows.length;
      i++
    ) {

      const row =
        rows[i];


      // ----------------------------------------------------
      // ORDER ID
      // Column A
      // ----------------------------------------------------

      const orderId =
        String(
          row[0] || ""
        )
        .trim()
        .toUpperCase();


      if (
        !orderId
      ) {

        continue;

      }


      // ----------------------------------------------------
      // ORDER DATE
      // Column B
      // ----------------------------------------------------

      if (
        !row[1]
      ) {

        continue;

      }


      const orderDate =
        new Date(
          row[1]
        );


      if (
        isNaN(
          orderDate.getTime()
        )
      ) {

        continue;

      }


      // ----------------------------------------------------
      // DATE FILTER
      // ----------------------------------------------------

      if (
        orderDate.getTime() <
          fromDate.getTime() ||
        orderDate.getTime() >
          toDate.getTime()
      ) {

        continue;

      }


      // ----------------------------------------------------
      // ORDER ID FILTER
      // ----------------------------------------------------

      if (
        orderIdFilter &&
        orderId !==
          orderIdFilter
      ) {

        continue;

      }


      // ----------------------------------------------------
      // MOBILE
      // Column D
      //
      // Strip spaces, +91, hyphens etc.
      // ----------------------------------------------------

      const mobile =
        String(
          row[3] || ""
        )
        .replace(/\D/g, "");


      if (
        mobileFilter
      ) {

        /*
         * endsWith() allows:
         *
         * Sheet:
         * 919938234400
         *
         * Search:
         * 9938234400
         *
         * to still match.
         */

        if (
          !mobile.endsWith(
            mobileFilter
          )
        ) {

          continue;

        }

      }


      // ----------------------------------------------------
      // PAYMENT STATUS
      // Column J
      // ----------------------------------------------------

      const paymentStatus =
        String(
          row[9] || ""
        )
        .trim();


      if (
        paymentStatusFilter &&
        paymentStatus
          .toLowerCase() !==
          paymentStatusFilter
      ) {

        continue;

      }


      // ----------------------------------------------------
      // CUSTOMER
      // Column C
      // ----------------------------------------------------

      const customerName =
        String(
          row[2] || ""
        ).trim();


      // ----------------------------------------------------
      // GRAND TOTAL
      // Column K
      // ----------------------------------------------------

      const grandTotal =
        Number(
          row[10]
        ) || 0;


      // ----------------------------------------------------
      // ORDER STATUS
      // Column L
      // ----------------------------------------------------

      const orderStatus =
        String(
          row[11] || ""
        ).trim();


      // ====================================================
      // ADD RESULT
      // ====================================================

      results.push({

        orderId:
          orderId,

        orderDate:
          Utilities.formatDate(
            orderDate,
            timezone,
            "dd/MM/yyyy"
          ),

        customerName:
          customerName,

        mobile:
          String(
            row[3] || ""
          ),

        paymentMode:
          String(
            row[8] || ""
          ),

        paymentStatus:
          paymentStatus,

        grandTotal:
          grandTotal,

        orderStatus:
          orderStatus

      });


      totalAmount +=
        grandTotal;

    }


    // ======================================================
    // SORT
    //
    // Latest order first.
    // ======================================================

    results.sort(
      function(a, b) {

        const aParts =
          a.orderDate.split("/");

        const bParts =
          b.orderDate.split("/");


        const aDate =
          new Date(
            Number(aParts[2]),
            Number(aParts[1]) - 1,
            Number(aParts[0])
          );


        const bDate =
          new Date(
            Number(bParts[2]),
            Number(bParts[1]) - 1,
            Number(bParts[0])
          );


        return (
          bDate.getTime() -
          aDate.getTime()
        );

      }
    );


    // ======================================================
    // RETURN
    // ======================================================

    return {

      success: true,

      orders:
        results,

      count:
        results.length,

      totalAmount:
        totalAmount,

      usingCurrentMonth:
        usingCurrentMonth,

      dateRange: {

        from:
          Utilities.formatDate(
            fromDate,
            timezone,
            "dd/MM/yyyy"
          ),

        to:
          Utilities.formatDate(
            toDate,
            timezone,
            "dd/MM/yyyy"
          )

      }

    };


  } catch (error) {

    console.error(
      "searchBulkInvoiceOrdersFromWeb Error:",
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
 * ==========================================================
 * GET BULK INVOICE DATA
 * ==========================================================
 *
 * Loads complete data for multiple selected orders.
 *
 * IMPORTANT:
 * - Independent from getInvoice()
 * - Does NOT call getInvoice() repeatedly
 * - Reads Orders once
 * - Reads Order Items once
 * - Validates that all selected orders belong
 *   to the same mobile number
 * ==========================================================
 */
function getBulkInvoiceData(orderIds) {

  try {

    // ======================================================
    // VALIDATE INPUT
    // ======================================================

    if (
      !Array.isArray(orderIds) ||
      orderIds.length === 0
    ) {

      return {
        success: false,
        message: "No orders were selected."
      };

    }


    // ======================================================
    // NORMALIZE ORDER IDS
    // ======================================================

    const normalizedOrderIds = [];

    const orderIdSet = new Set();


    for (
      let i = 0;
      i < orderIds.length;
      i++
    ) {

      const orderId =
        String(
          orderIds[i] || ""
        )
        .trim()
        .toUpperCase();


      if (
        orderId &&
        !orderIdSet.has(orderId)
      ) {

        normalizedOrderIds.push(
          orderId
        );

        orderIdSet.add(
          orderId
        );

      }

    }


    if (
      normalizedOrderIds.length === 0
    ) {

      return {
        success: false,
        message: "No valid Order IDs were selected."
      };

    }


    // ======================================================
    // LOAD SHEETS
    // ======================================================

    const ordersSheet =
      getSheet(
        SHEETS.ORDERS
      );


    const itemsSheet =
      getSheet(
        SHEETS.ORDER_ITEMS
      );


    const orderLastRow =
      ordersSheet.getLastRow();


    if (
      orderLastRow <= 1
    ) {

      return {
        success: false,
        message: "Orders sheet is empty."
      };

    }


    // ======================================================
    // READ ORDERS ONCE
    //
    // A:R = 18 columns
    // ======================================================

    const orderRows =
      ordersSheet
        .getRange(
          2,
          1,
          orderLastRow - 1,
          18
        )
        .getValues();


    const timezone =
      Session.getScriptTimeZone();


    const selectedOrders = [];


    // ======================================================
    // FIND SELECTED ORDERS
    // ======================================================

    for (
      let i = 0;
      i < orderRows.length;
      i++
    ) {

      const row =
        orderRows[i];


      const orderId =
        String(
          row[0] || ""
        )
        .trim()
        .toUpperCase();


      if (
        !orderIdSet.has(orderId)
      ) {

        continue;

      }


      // ----------------------------------------------------
      // Order Date
      // ----------------------------------------------------

      let orderDate = "";

      let orderDateValue = 0;


      if (row[1]) {

        const date =
          new Date(
            row[1]
          );


        if (
          !isNaN(date.getTime())
        ) {

          orderDateValue =
            date.getTime();


          orderDate =
            Utilities.formatDate(
              date,
              timezone,
              "dd/MM/yyyy"
            );

        }

      }


      // ----------------------------------------------------
      // Store order
      // ----------------------------------------------------

      selectedOrders.push({

        orderId:
          orderId,

        orderDate:
          orderDate,

        orderDateValue:
          orderDateValue,

        customerName:
          String(
            row[2] || ""
          ).trim(),

        mobile:
          String(
            row[3] || ""
          ).trim(),

        deliveryArea:
          String(
            row[4] || ""
          ).trim(),

        houseAddress:
          String(
            row[5] || ""
          ).trim(),

        deliverySlot:
          String(
            row[6] || ""
          ).trim(),

        orderType:
          String(
            row[7] || ""
          ).trim(),

        paymentMode:
          String(
            row[8] || ""
          ).trim(),

        paymentStatus:
          String(
            row[9] || ""
          ).trim(),

        grandTotal:
          Number(
            row[10]
          ) || 0,

        status:
          String(
            row[11] || ""
          ).trim(),

        specialInstructions:
          String(
            row[12] || ""
          ).trim(),

        deliveryCharge:
          Number(
            row[14]
          ) || 0,

        discount:
          Number(
            row[15]
          ) || 0,

        addons:
          String(
            row[16] || ""
          ).trim(),

        addonTotal:
          Number(
            row[17]
          ) || 0,

        items: []

      });

    }


    // ======================================================
    // MAKE SURE ALL SELECTED ORDERS WERE FOUND
    // ======================================================

    if (
      selectedOrders.length !==
      normalizedOrderIds.length
    ) {

      const foundIds =
        new Set(
          selectedOrders.map(
            function(order) {

              return order.orderId;

            }
          )
        );


      const missingIds =
        normalizedOrderIds.filter(
          function(orderId) {

            return !foundIds.has(
              orderId
            );

          }
        );


      return {

        success: false,

        message:
          "The following orders could not be found: " +
          missingIds.join(", ")

      };

    }


    // ======================================================
    // VALIDATE SAME CUSTOMER / MOBILE
    // ======================================================

    const normalizeMobile =
      function(value) {

        let mobile =
          String(
            value || ""
          ).replace(/\D/g, "");


        // Keep last 10 digits for Indian mobile numbers.

        if (
          mobile.length > 10
        ) {

          mobile =
            mobile.slice(-10);

        }


        return mobile;

      };


    const firstMobile =
      normalizeMobile(
        selectedOrders[0].mobile
      );


    if (!firstMobile) {

      return {

        success: false,

        message:
          "The selected customer's mobile number is missing."

      };

    }


    for (
      let i = 1;
      i < selectedOrders.length;
      i++
    ) {

      const mobile =
        normalizeMobile(
          selectedOrders[i].mobile
        );


      if (
        mobile !== firstMobile
      ) {

        return {

          success: false,

          message:
            "Selected orders belong to different customers. Bulk invoice can only contain orders for one mobile number."

        };

      }

    }


    // ======================================================
    // CREATE ORDER LOOKUP
    // ======================================================

    const orderLookup = {};


    for (
      let i = 0;
      i < selectedOrders.length;
      i++
    ) {

      orderLookup[
        selectedOrders[i].orderId
      ] = selectedOrders[i];

    }


    // ======================================================
    // LOAD ORDER ITEMS ONCE
    // ======================================================

    const itemLastRow =
      itemsSheet.getLastRow();


    if (
      itemLastRow > 1
    ) {

      const itemRows =
        itemsSheet
          .getRange(
            2,
            1,
            itemLastRow - 1,
            5
          )
          .getValues();


      for (
        let i = 0;
        i < itemRows.length;
        i++
      ) {

        const row =
          itemRows[i];


        const orderId =
          String(
            row[0] || ""
          )
          .trim()
          .toUpperCase();


        const order =
          orderLookup[
            orderId
          ];


        if (!order) {

          continue;

        }


        order.items.push({

          itemName:
            String(
              row[1] || ""
            ).trim(),

          qty:
            Number(
              row[2]
            ) || 0,

          price:
            Number(
              row[3]
            ) || 0,

          total:
            Number(
              row[4]
            ) || 0

        });

      }

    }


    // ======================================================
    // SORT SELECTED ORDERS BY DATE
    //
    // Oldest -> Newest makes the consolidated invoice
    // easier to read.
    // ======================================================

    selectedOrders.sort(
      function(a, b) {

        if (
          a.orderDateValue !==
          b.orderDateValue
        ) {

          return (
            a.orderDateValue -
            b.orderDateValue
          );

        }


        return String(
          a.orderId
        ).localeCompare(
          String(b.orderId),
          undefined,
          {
            numeric: true
          }
        );

      }
    );


    // ======================================================
    // CALCULATE BULK TOTALS
    // ======================================================

    let grandTotal = 0;

    let paidAmount = 0;

    let pendingAmount = 0;

    let deliveryChargeTotal = 0;

    let discountTotal = 0;

    let addonTotal = 0;


    for (
      let i = 0;
      i < selectedOrders.length;
      i++
    ) {

      const order =
        selectedOrders[i];


      grandTotal +=
        Number(
          order.grandTotal || 0
        );


      deliveryChargeTotal +=
        Number(
          order.deliveryCharge || 0
        );


      discountTotal +=
        Number(
          order.discount || 0
        );


      addonTotal +=
        Number(
          order.addonTotal || 0
        );


      const paymentStatus =
        String(
          order.paymentStatus || ""
        )
        .trim()
        .toLowerCase();


      if (
        paymentStatus === "paid"
      ) {

        paidAmount +=
          Number(
            order.grandTotal || 0
          );

      }


      if (
        paymentStatus === "pending"
      ) {

        pendingAmount +=
          Number(
            order.grandTotal || 0
          );

      }

    }


    // ======================================================
    // BILLING PERIOD
    // ======================================================

    let billingFrom = "";

    let billingTo = "";


    if (
      selectedOrders.length
    ) {

      billingFrom =
        selectedOrders[0].orderDate;


      billingTo =
        selectedOrders[
          selectedOrders.length - 1
        ].orderDate;

    }


    // ======================================================
    // REMOVE INTERNAL SORT VALUE
    // ======================================================

    for (
      let i = 0;
      i < selectedOrders.length;
      i++
    ) {

      delete selectedOrders[i]
        .orderDateValue;

    }


    // ======================================================
    // RETURN BULK INVOICE DATA
    // ======================================================

    return {

      success: true,

      customer: {

        customerName:
          selectedOrders[0]
            .customerName,

        mobile:
          selectedOrders[0]
            .mobile,

        deliveryArea:
          selectedOrders[0]
            .deliveryArea,

        houseAddress:
          selectedOrders[0]
            .houseAddress

      },


      billingPeriod: {

        from:
          billingFrom,

        to:
          billingTo

      },


      orderCount:
        selectedOrders.length,


      orders:
        selectedOrders,


      totals: {

        grandTotal:
          grandTotal,

        paidAmount:
          paidAmount,

        pendingAmount:
          pendingAmount,

        deliveryCharge:
          deliveryChargeTotal,

        discount:
          discountTotal,

        addonTotal:
          addonTotal

      }

    };


  } catch (error) {

    console.error(
      "getBulkInvoiceData Error:",
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
 * Generate Bulk Invoice PDF
 * ==========================================
 */
function generateBulkInvoicePdf(orderIds) {
  try {
    const bulkData = getBulkInvoiceData(orderIds);

    if (!bulkData.success) {
      return bulkData;
    }

    const template = HtmlService.createTemplateFromFile("BulkInvoicePdf");

    // Fetch the logo from Google Drive as Base64
    const logoResponse = getSmokySwaadLogo();
    template.logoDataUrl = logoResponse.success ? logoResponse.dataUrl : "";

    // Attach data to template
    template.customer = bulkData.customer;
    template.billingPeriod = bulkData.billingPeriod;
    template.orderCount = bulkData.orderCount;
    template.orders = bulkData.orders;
    template.totals = bulkData.totals;

    const html = template.evaluate().getContent();

    const pdfBlob = Utilities.newBlob(html, "text/html").getAs(MimeType.PDF);
    const fileName = "Bulk_Invoice_" + bulkData.customer.mobile + ".pdf";
    pdfBlob.setName(fileName);

    return {
      success: true,
      message: "Bulk Invoice PDF generated successfully.",
      blobName: pdfBlob.getName(),
      blobSize: pdfBlob.getBytes().length,
      // Pass base64 data for client download/sharing
      base64Data: Utilities.base64Encode(pdfBlob.getBytes()),
      customerMobile: bulkData.customer.mobile,
      totals: bulkData.totals
    };

  } catch (err) {
    console.error("generateBulkInvoicePdf Error:", err);
    return {
      success: false,
      message: err.toString()
    };
  }
}
