/**
 * ==========================================
 * SMOKY SWAAD ERP
 * Web Application Entry Point
 * ==========================================
 */

function doGet() {

  return HtmlService
      .createTemplateFromFile("Index")
      .evaluate()
      .addMetaTag(
          "viewport",
          "width=device-width, initial-scale=1, viewport-fit=cover"
      )
      .setTitle("Smoky Swaad")
      .setXFrameOptionsMode(
          HtmlService.XFrameOptionsMode.ALLOWALL
      );

}

/**
 * Include HTML files
 */
function include(filename){

  return HtmlService
      .createHtmlOutputFromFile(filename)
      .getContent();

}


/**
 * ==========================================
 * Load Screen
 * ==========================================
 */
function loadScreen(screen) {
  // Changing this to createTemplateFromFile tells Apps Script to process the <?!= ?> tags
  return HtmlService.createTemplateFromFile(screen).evaluate().getContent();
}



function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}




/**
 * ==========================================
 * ORDER HISTORY BACKEND SERVICE
 * ==========================================
 */
function getHeaderIndex(headers, possibleNames) {
  for (let name of possibleNames) {
    const idx = headers.indexOf(name.toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}


function getOrderHistoryData(targetMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Orders");
    if (!sheet) {
      return { success: false, message: "Orders sheet not found." };
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return { success: true, orders: [] };
    }

    const headers = rows[0].map(h => String(h).trim().toLowerCase());
    
    const idxOrderId = getHeaderIndex(headers, ["order id", "orderid"]);
    const idxDate = getHeaderIndex(headers, ["order date", "date"]);
    const idxCustomer = getHeaderIndex(headers, ["customer name", "customer", "client name", "name"]);
    const idxMobile = getHeaderIndex(headers, ["mobile", "mobile number", "phone"]);
    const idxType = getHeaderIndex(headers, ["order type", "type"]);
    const idxSlot = getHeaderIndex(headers, ["delivery slot", "slot"]);
    const idxPaymentMode = getHeaderIndex(headers, ["payment mode", "payment"]);
    const idxPaymentStatus = getHeaderIndex(headers, ["payment status", "status"]);
    const idxGrandTotal = getHeaderIndex(headers, ["grand total", "total", "amount"]);

    // Find the items sheet dynamically (checks OrderItems, Order Items, Items, InvoiceItems)
    const possibleItemSheetNames = ["OrderItems", "Order Items", "Items", "InvoiceItems", "Invoice Items"];
    let itemsSheet = null;
    for (let sName of possibleItemSheetNames) {
      itemsSheet = ss.getSheetByName(sName);
      if (itemsSheet) break;
    }

    const itemMap = {};
    
    if (itemsSheet && itemsSheet.getLastRow() > 1) {
      const itemRows = itemsSheet.getDataRange().getValues();
      const itemHeaders = itemRows[0].map(h => String(h).trim().toLowerCase());
      const iOrderId = getHeaderIndex(itemHeaders, ["order id", "orderid", "id"]);
      const iItemName = getHeaderIndex(itemHeaders, ["item name", "item", "name", "menu item", "product"]);
      const iQty = getHeaderIndex(itemHeaders, ["qty", "quantity", "count"]);

      for (let j = 1; j < itemRows.length; j++) {
        const iRow = itemRows[j];
        const oId = iOrderId !== -1 ? String(iRow[iOrderId] || "").trim() : "";
        if (!oId) continue;
        
        const itemName = iItemName !== -1 ? String(iRow[iItemName] || "").trim() : "";
        const qty = iQty !== -1 ? Number(iRow[iQty] || 1) : 1;

        if (itemName) {
          if (!itemMap[oId]) {
            itemMap[oId] = [];
          }
          itemMap[oId].push({ name: itemName, qty: qty });
        }
      }
    }

    const orders = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const orderId = idxOrderId !== -1 ? String(row[idxOrderId] || "").trim() : "";
      if (!orderId) continue;

      const dateVal = idxDate !== -1 ? row[idxDate] : "";
      let dateStr = "";
      let dateRaw = "";

      if (dateVal instanceof Date) {
         const y = dateVal.getFullYear();
         const m = String(dateVal.getMonth() + 1).padStart(2, "0");
         const d = String(dateVal.getDate()).padStart(2, "0");
         dateRaw = `${y}-${m}-${d}`;
         dateStr = `${d}/${m}/${y}`;
      } else {
         dateRaw = String(dateVal).trim();
         dateStr = dateRaw;
      }

      if (targetMonth && !dateRaw.startsWith(targetMonth)) {
         continue;
      }

      orders.push({
        orderId: orderId,
        orderDate: dateStr,
        orderDateRaw: dateRaw,
        customerName: idxCustomer !== -1 ? String(row[idxCustomer] || "") : "",
        mobile: idxMobile !== -1 ? String(row[idxMobile] || "") : "",
        orderType: idxType !== -1 ? String(row[idxType] || "") : "Delivery",
        deliverySlot: idxSlot !== -1 ? String(row[idxSlot] || "") : "",
        paymentMode: idxPaymentMode !== -1 ? String(row[idxPaymentMode] || "") : "Cash",
        paymentStatus: idxPaymentStatus !== -1 ? String(row[idxPaymentStatus] || "") : "Pending",
        grandTotal: idxGrandTotal !== -1 ? Number(row[idxGrandTotal] || 0) : 0,
        items: itemMap[orderId] || []
      });
    }

    return { success: true, orders: orders };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  }
}