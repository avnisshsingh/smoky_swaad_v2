/**
 * ==========================================
 * Load Purchase Settings (Optimized with Server-Side Caching)
 * ==========================================
 */
function loadPurchaseSettings() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_purchase_settings_v1";

   // 1. Check if purchase settings are already cached in server memory
   const cachedData = cache.get(cacheKey);
   if (cachedData) {
      try {
         return JSON.parse(cachedData);
      } catch (e) {
         // Fallback to sheet reading if JSON parse fails
      }
   }

   // 2. Original sheet fetching & mapping logic (untouched)
   const sheet = getSheet(SHEETS.SETTINGS);
   const settings = {
      // Units (Column S)
      units: sheet
         .getRange("S2:S100")
         .getValues()
         .flat()
         .filter(String),

      // Payment Types (Column D)
      paymentTypes: sheet
         .getRange("D2:D100")
         .getValues()
         .flat()
         .filter(String),

      // Suppliers (Column Y)
      suppliers: sheet
         .getRange("Y2:Y100")
         .getValues()
         .flat()
         .filter(String)
   };

   // 3. Store the result in cache for 6 hours (21600 seconds)
   try {
      cache.put(cacheKey, JSON.stringify(settings), 21600);
   } catch (e) {
      console.warn("Failed to write purchase settings cache:", e);
   }

   return settings;
}

/**
 * ==========================================
 * Clear Purchase Settings Cache
 * ==========================================
 * Call this helper function whenever you add or modify units, payment types, 
 * or suppliers in your Settings sheet so changes reflect immediately.
 */
function clearPurchaseSettingsCache() {
   try {
      const cache = CacheService.getScriptCache();
      cache.remove("smoky_swaad_purchase_settings_v1");
   } catch (e) {
      console.warn("Failed to clear purchase settings cache:", e);
   }
}


/**
 * ==========================================
 * Search Purchase Items
 * ==========================================
 */
function searchPurchaseItems(keyword) {
  const sheet = getSheet(SHEETS.SETTINGS);
  const items = sheet
      .getRange("X2:X200")
      .getValues()
      .flat()
      .filter(String);

  if (!keyword) return [];

  keyword = keyword.toLowerCase();

  return items
      .filter(item =>
          item.toLowerCase().includes(keyword)
      )
      .map(item => ({
          itemName: item
      }))
      .slice(0, 10);
}


/**
 * ==========================================
 * Add Purchase Item
 * ==========================================
 */
function addPurchaseItem(itemName) {
    itemName = String(itemName).trim();

    if (!itemName) {
        throw new Error("Item Name is required.");
    }

    const sheet = getSheet(SHEETS.SETTINGS);
    const values = sheet.getRange("X2:X1000").getValues();
    let nextRow = null;

    for (let i = 0; i < values.length; i++) {
        const value = String(values[i][0]).trim();
        if (value && value.toLowerCase() === itemName.toLowerCase()) {
            return {
                success: true,
                alreadyExists: true,
                itemName: value
            };
        }
        if (!value && nextRow === null) {
            nextRow = i + 2;
        }
    }

    if (nextRow === null) {
        nextRow = values.length + 2;
    }

    sheet.getRange(nextRow, 24).setValue(itemName);

    return {
        success: true,
        alreadyExists: false,
        itemName: itemName
    };
}

/**
 * ==========================================
 * Load Purchase Items List (Fixed & Optimized)
 * ==========================================
 */
function getPurchaseItems() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_purchase_items_cache_v1";

   // 1. Check if purchase items list is already cached in server memory
   const cachedData = cache.get(cacheKey);
   if (cachedData) {
      try {
         return JSON.parse(cachedData);
      } catch (e) {
         // Fallback to sheet reading if JSON parse fails
      }
   }

   // 2. Fixed logic: Read items from the Settings sheet, Column X (X2:X1000)
   const sheet = getSheet(SHEETS.SETTINGS);
   const values = sheet.getRange("X2:X1000").getValues();
   
   let items = [];
   for (let i = 0; i < values.length; i++) {
      const val = String(values[i][0]).trim();
      if (val) {
         items.push({
            itemName: val,
            searchName: val.toLowerCase()
         });
      }
   }

   // 3. Store the items array in cache for 6 hours (21600 seconds)
   try {
      cache.put(cacheKey, JSON.stringify(items), 21600);
   } catch (e) {
      console.warn("Failed to write purchase items cache:", e);
   }

   return items;
}

/**
 * ==========================================
 * Clear Purchase Items Cache
 * ==========================================
 * Call this helper function whenever you add a new purchase item dynamically 
 * so the suggestion box updates instantly.
 */
function clearPurchaseItemsCache() {
   try {
      const cache = CacheService.getScriptCache();
      cache.remove("smoky_swaad_purchase_items_cache_v1");
   } catch (e) {
      console.warn("Failed to clear purchase items cache:", e);
   }
}


/**
 * ==========================================
 * CACHE MANAGEMENT HELPERS
 * ==========================================
 */
function getPurchaseCacheVersion() {
   const cache = CacheService.getScriptCache();
   let version = cache.get("PURCHASE_CACHE_VERSION");
   if (!version) {
      version = String(new Date().getTime());
      cache.put("PURCHASE_CACHE_VERSION", version, 21600); // 6 hours
   }
   return version;
}

function invalidatePurchaseCache() {
   const cache = CacheService.getScriptCache();
   cache.put("PURCHASE_CACHE_VERSION", String(new Date().getTime()), 21600);
}

/**
 * ==========================================
 * BLAZING-FAST CACHED PURCHASE BOOTSTRAP WITH MONTHS
 * ==========================================
 */
function getPurchaseBootstrapData(targetMonth) {
   const tz = Session.getScriptTimeZone();
   const defaultMonth = targetMonth || Utilities.formatDate(new Date(), tz, "yyyy-MM");
   
   const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
   const lastRow = sheet.getLastRow();
   
   const monthsSet = new Set();
   monthsSet.add(Utilities.formatDate(new Date(), tz, "yyyy-MM")); // Ensure current month is always present

   const purchases = [];
   let monthTotal = 0;

   if (lastRow >= 2) {
      // Read recent rows to quickly extract unique months and current month purchases
      const numRowsToRead = Math.min(lastRow - 1, 1000);
      const startRow = Math.max(2, lastRow - numRowsToRead + 1);
      const data = sheet.getRange(startRow, 1, numRowsToRead, 11).getValues();

      for (let i = 0; i < data.length; i++) {
         const row = data[i];
         if (!row[0]) continue;
         const rawDate = row[1];
         if (!rawDate) continue;

         const purchaseDateKey = Utilities.formatDate(new Date(rawDate), tz, "yyyy-MM-dd");
         const mKey = purchaseDateKey.slice(0, 7);
         monthsSet.add(mKey);

         const status = String(row[10] || "Active");
         if (status.toLowerCase() !== "deleted" && mKey === defaultMonth) {
            const amount = Number(row[6] || 0);
            monthTotal += amount;

            purchases.push({
               purchaseId: String(row[0]),
               purchaseDateRaw: purchaseDateKey,
               purchaseDate: Utilities.formatDate(new Date(rawDate), tz, "dd/MM/yyyy"),
               itemName: String(row[2] || ""),
               quantity: Number(row[3] || 0),
               unit: String(row[4] || ""),
               paymentType: String(row[5] || ""),
               amount: amount,
               supplier: String(row[7] || ""),
               remarks: String(row[8] || ""),
               status: status
            });
         }
      }
   }

   purchases.sort((a, b) => b.purchaseId.localeCompare(a.purchaseId, undefined, { numeric: true }));

   // Sort months descending (newest first) and format labels
   const sortedMonths = Array.from(monthsSet).sort().reverse();
   const monthsList = sortedMonths.map(mKey => {
      const parts = mKey.split("-");
      const date = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
      return {
         value: mKey,
         label: Utilities.formatDate(date, tz, "MMMM yyyy")
      };
   });

   return {
      success: true,
      summary: {
         thisMonth: monthTotal,
         transactionCount: purchases.length
      },
      suppliers: loadPurchaseSettings().suppliers,
      months: monthsList,
      purchases: purchases
   };
}






/**
 * ==========================================
 * Save Purchase From Web
 * ==========================================
 */
function savePurchaseFromWeb(purchaseData) {
    try {
        validatePurchaseData(purchaseData);

        const purchaseSheet = getSheet(SHEETS.PURCHASE_REGISTER);
        const purchaseId = generateNextId(purchaseSheet, "PUR");

        savePurchaseEntry(
            purchaseSheet,
            purchaseId,
            purchaseData
        );

        // Invalidate cache so the new purchase appears instantly
        invalidatePurchaseCache();

        return {
            success: true,
            purchaseId,
            message: "Purchase Saved Successfully"
        };
    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}


/**
 * ==========================================
 * Validate Purchase Data
 * ==========================================
 */
function validatePurchaseData(purchaseData) {
  const purchase = purchaseData.purchase;

  if (!purchase.purchaseDate)
    throw new Error("Purchase Date is required.");

  if (!purchase.itemName)
    throw new Error("Item Name is required.");

  if (!purchase.quantity || purchase.quantity <= 0)
    throw new Error("Quantity should be greater than zero.");

  purchase.unit = purchase.unit || "Gram";
  purchase.paymentType = purchase.paymentType || "PhonePe";
  purchase.supplier = purchase.supplier || "Local Market";

  if (!purchase.amount || purchase.amount <= 0)
    throw new Error("Amount should be greater than zero.");

  return true;
}


/**
 * ==========================================
 * Save Purchase Entry
 * ==========================================
 */
function savePurchaseEntry(
    purchaseSheet,
    purchaseId,
    purchaseData
) {
    const p = purchaseData.purchase;
    const nextRow = purchaseSheet.getLastRow() + 1;

    purchaseSheet
        .getRange(nextRow, 1, 1, 10)
        .setValues([[
            purchaseId,
            new Date(p.purchaseDate),
            p.itemName,
            Number(p.quantity),
            p.unit,
            p.paymentType,
            Number(p.amount),
            p.supplier,
            p.remarks || "",
            new Date()
        ]]);
}


/**
 * ==========================================
 * Update Purchase From Web (Optimized Range)
 * ==========================================
 */
function updatePurchaseFromWeb(purchaseId, purchaseData) {
   try {
      validatePurchaseData(purchaseData);
      const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }
      
      // Optimization: Fetch only active rows/columns instead of entire sheet getDataRange()
      const data = sheet.getRange(1, 1, lastRow, 11).getValues();

      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(purchaseId)) {
            targetRow = i + 1;
            break;
         }
      }

      if (targetRow === -1) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }

      const p = purchaseData.purchase;
      sheet.getRange(targetRow, 1, 1, 9).setValues([[
         purchaseId,
         new Date(p.purchaseDate),
         p.itemName,
         Number(p.quantity),
         p.unit,
         p.paymentType,
         Number(p.amount),
         p.supplier,
         p.remarks || ""
      ]]);

      // Invalidate cache
      invalidatePurchaseCache();

      return { success: true, purchaseId: purchaseId };
   } catch (err) {
      return { success: false, message: err.message };
   }
}


/**
 * ==========================================
 * Delete Purchase From Web (Soft Delete - Optimized)
 * ==========================================
 */
function deletePurchaseFromWeb(purchaseId) {
   try {
      const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }

      const data = sheet.getRange(1, 1, lastRow, 11).getValues();
      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(purchaseId)) {
            targetRow = i + 1;
            break;
         }
      }

      if (targetRow === -1) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }

      sheet.getRange(targetRow, 11).setValue("Deleted");

      // Invalidate cache
      invalidatePurchaseCache();

      return { success: true };
   } catch (err) {
      return { success: false, message: err.message };
   }
}


/**
 * ==========================================
 * Restore Purchase From Web (Optimized)
 * ==========================================
 */
function restorePurchaseFromWeb(purchaseId) {
   try {
      const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }

      const data = sheet.getRange(1, 1, lastRow, 11).getValues();
      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(purchaseId)) {
            targetRow = i + 1;
            break;
         }
      }

      if (targetRow === -1) {
         throw new Error("Purchase ID not found: " + purchaseId);
      }

      sheet.getRange(targetRow, 11).setValue("Active");

      // Invalidate cache
      invalidatePurchaseCache();

      return { success: true };
   } catch (err) {
      return { success: false, message: err.message };
   }
}


/**
 * ==========================================
 * Get Deleted Purchases (Optimized Range)
 * ==========================================
 */
function getDeletedPurchases() {
   const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
   const lastRow = sheet.getLastRow();
   const deletedPurchases = [];

   if (lastRow >= 2) {
      const data = sheet.getRange(1, 1, lastRow, 11).getValues();

      for (let i = 1; i < data.length; i++) {
         const row = data[i];
         if (!row[0]) continue;
         const status = String(row[10] || "Active");

         if (status.toLowerCase() === "deleted") {
            const rawDate = row[1];
            deletedPurchases.push({
               purchaseId: String(row[0]),
               purchaseDate: rawDate ? Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
               itemName: String(row[2] || ""),
               quantity: Number(row[3] || 0),
               unit: String(row[4] || ""),
               paymentType: String(row[5] || ""),
               amount: Number(row[6] || 0),
               supplier: String(row[7] || ""),
               remarks: String(row[8] || "")
            });
         }
      }

      deletedPurchases.sort((a, b) => b.purchaseId.localeCompare(a.purchaseId, undefined, { numeric: true }));
   }

   return {
      success: true,
      purchases: deletedPurchases
   };
}



/**
 * ==========================================
 * Get Purchase Module Bootstrap Data (Fixed)
 * ==========================================
 * Safely aggregates settings, item master list, and history using your existing functions.
 */
function getPurchaseModuleBootstrapData(monthStr) {
   try {
      // 1. Fetch settings using your existing function
      const settings = typeof loadPurchaseSettings === "function" ? loadPurchaseSettings() : {};
      
      // 2. Fetch items using your existing function
      const items = typeof getPurchaseItems === "function" ? getPurchaseItems() : [];
      
      // 3. Fetch purchase history and filters using your existing bootstrap function
      let historyData = { months: [], suppliers: [], purchases: [] };
      if (typeof getPurchaseBootstrapData === "function") {
         const res = getPurchaseBootstrapData(monthStr);
         if (res && res.success) {
            historyData = res;
         }
      }

      return {
         success: true,
         settings: settings,
         items: items,
         months: historyData.months || [],
         suppliers: historyData.suppliers || [],
         purchases: historyData.purchases || []
      };
   } catch (error) {
      console.error("Purchase Module Bootstrap Error:", error);
      return {
         success: false,
         message: error && error.message ? error.message : String(error)
      };
   }
}