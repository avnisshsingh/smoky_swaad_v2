/**
 * Load Complete Active Menu (Optimized with Server-Side Caching)
 */
function loadMenu() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_menu_cache_v1";

   // 1. Check if the active menu is already cached in server memory
   const cachedData = cache.get(cacheKey);
   if (cachedData) {
      try {
         return JSON.parse(cachedData);
      } catch (e) {
         // Fallback to sheet reading if JSON parse fails
      }
   }

   // 2. Original sheet fetching & mapping logic (untouched)
   const menu = getSheet(SHEETS.MENU);
   const lastRow = menu.getLastRow();
   
   if (lastRow < 2) return [];

   const data = menu
      .getRange(2, 1, lastRow - 1, 9)
      .getValues();

   const menuItems = data
      .filter(row => row[8] === "Yes")
      .map(row => ({
         itemId: row[0],
         itemName: row[1],
         category: row[2],
         cost: Number(row[3]),
         price: Number(row[4]),
         prepTime: row[7]
      }));

   // 3. Store the menu array in cache for 6 hours (21600 seconds)
   try {
      cache.put(cacheKey, JSON.stringify(menuItems), 21600);
   } catch (e) {
      console.warn("Failed to write menu cache:", e);
   }

   return menuItems;
}

/**
 * CRITICAL: Clear Menu Cache
 * Call this helper function in your backend code whenever a menu item 
 * is added, edited, or deleted in your Google Sheet so changes appear instantly.
 */
function clearMenuCache() {
   try {
      const cache = CacheService.getScriptCache();
      cache.remove("smoky_swaad_menu_cache_v1");
   } catch (e) {
      console.warn("Failed to clear menu cache:", e);
   }
}






/**
 * ==========================================
 * Load Complete POS Cache
 * ==========================================
 */
function loadPOSCache() {
  return getCachedData("CACHE_POS_PAYLOAD", function() {
    return {
      menu: loadMenu(),          
      settings: loadAppSettings(), 
      customers: getAllCustomers() // Corrected function name
    };
  });
}



/**
 * Add a food item to the existing Menu sheet.
 * Menu columns used by the current POS:
 * A Item ID, B Item Name, C Category, D Cost, E Selling Price,
 * H Preparation Time, I Active ("Yes"/"No").
 */
function addFoodItemToMenu(payload) {
   const lock = LockService.getDocumentLock();
   if (!lock.tryLock(10000)) {
      return { success: false, message: "Another menu update is in progress. Please try again." };
   }

   try {
      payload = payload || {};
      const itemName = String(payload.itemName || "").trim();
      const category = String(payload.category || "").trim();
      const price = Number(payload.price);
      const cost = Number(payload.cost || 0);
      const prepTime = Number(payload.prepTime || 0);
      const isActive = payload.isActive === true;

      if (!itemName) return { success: false, message: "Food item name is required." };
      if (!category) return { success: false, message: "Please select a category." };
      if (!Number.isFinite(price) || price <= 0) return { success: false, message: "Selling price must be greater than zero." };
      if (!Number.isFinite(cost) || cost < 0) return { success: false, message: "Cost price cannot be negative." };
      if (!Number.isFinite(prepTime) || prepTime < 0) return { success: false, message: "Preparation time cannot be negative." };

      const sheet = getSheet(SHEETS.MENU);
      const lastRow = sheet.getLastRow();
      const data = lastRow >= 2
         ? sheet.getRange(2, 1, lastRow - 1, 9).getValues()
         : [];

      const normalizedName = itemName.toLocaleLowerCase().replace(/\s+/g, " ");
      const duplicate = data.some(function(row) {
         return String(row[1] || "").trim().toLocaleLowerCase().replace(/\s+/g, " ") === normalizedName;
      });
      if (duplicate) {
         return { success: false, message: "A food item with this name already exists in the Menu sheet." };
      }

      // Generate a unique menu ID based on existing IDs; do not depend on the last row's ID.
      let maxId = 0;
      data.forEach(function(row) {
         const match = String(row[0] || "").match(/^MENU[-_]?(\d+)$/i);
         if (match) maxId = Math.max(maxId, Number(match[1]));
      });
      let itemId = "MENU" + String(maxId + 1).padStart(5, "0");
      const existingIds = new Set(data.map(function(row) { return String(row[0] || ""); }));
      while (existingIds.has(itemId)) {
         maxId++;
         itemId = "MENU" + String(maxId + 1).padStart(5, "0");
      }

      const newRow = lastRow + 1;
      // Preserve existing columns F and G as blank; H is prep time and I is Active.
      sheet.getRange(newRow, 1, 1, 9).setValues([[
         itemId, itemName, category, cost, price, "", "", prepTime, isActive ? "Yes" : "No"
      ]]);

      // Invalidate both the menu cache and the combined POS bootstrap cache.
      clearMenuCache();
      try {
         CacheService.getScriptCache().remove("CACHE_POS_PAYLOAD");
         CacheService.getScriptCache().remove("CACHE_MENU_PRICING");
      } catch (cacheError) {
         console.warn("Menu cache invalidation warning:", cacheError);
      }

      return {
         success: true,
         message: isActive
            ? "Food item saved and made available in POS."
            : "Food item saved as inactive. Enable it in the Menu sheet before using it in POS.",
         item: {
            itemId: itemId,
            itemName: itemName,
            category: category,
            cost: cost,
            price: price,
            prepTime: prepTime,
            isActive: isActive
         }
      };
   } catch (error) {
      console.error("addFoodItemToMenu failed:", error);
      return { success: false, message: error && error.message ? error.message : String(error) };
   } finally {
      lock.releaseLock();
   }
}


/** Return all menu items, including inactive ones, for the admin menu manager. */
function getMenuItemsForManagement() {
  const sheet = getSheet(SHEETS.MENU);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 9).getValues()
    .map(function(row, index) {
      return {
        rowNumber: index + 2,
        itemId: String(row[0] || ""),
        itemName: String(row[1] || ""),
        category: String(row[2] || ""),
        cost: Number(row[3]) || 0,
        price: Number(row[4]) || 0,
        prepTime: Number(row[7]) || 0,
        isActive: String(row[8] || "").trim().toLowerCase() === "yes"
      };
    })
    .filter(function(item) { return item.itemName; });
}

/** Update category and prices for an existing menu item by its stable Item ID. */
function updateMenuItemDetails(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return { success: false, message: "Another menu update is in progress. Please try again." };
  try {
    payload = payload || {};
    const itemId = String(payload.itemId || "").trim();
    const category = String(payload.category || "").trim();
    const price = Number(payload.price);
    const cost = Number(payload.cost);
    const isActive = payload.isActive === true;
    if (!itemId) return { success: false, message: "Missing item ID." };
    if (!category) return { success: false, message: "Category is required." };
    if (!Number.isFinite(price) || price <= 0) return { success: false, message: "Selling price must be greater than zero." };
    if (!Number.isFinite(cost) || cost < 0) return { success: false, message: "Cost price cannot be negative." };

    const sheet = getSheet(SHEETS.MENU);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No menu items were found." };
    const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const index = rows.findIndex(function(row) { return String(row[0] || "").trim() === itemId; });
    if (index < 0) return { success: false, message: "Menu item not found. Refresh the list and try again." };

    const rowNumber = index + 2;
    sheet.getRange(rowNumber, 3).setValue(category);
    sheet.getRange(rowNumber, 4, 1, 2).setValues([[cost, price]]);
    sheet.getRange(rowNumber, 9).setValue(isActive ? "Yes" : "No");
    clearMenuCache();
    try {
      const cache = CacheService.getScriptCache();
      cache.remove("CACHE_POS_PAYLOAD");
      cache.remove("CACHE_MENU_PRICING");
    } catch (cacheError) { console.warn("Menu cache invalidation warning:", cacheError); }
    return { success: true, message: "Menu item updated successfully." };
  } catch (error) {
    console.error("updateMenuItemDetails failed:", error);
    return { success: false, message: error && error.message ? error.message : String(error) };
  } finally {
    lock.releaseLock();
  }
}

/** Remove an item from POS by marking it inactive; preserve historical order data. */
function deactivateMenuItem(itemId) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return { success: false, message: "Another menu update is in progress. Please try again." };
  try {
    itemId = String(itemId || "").trim();
    if (!itemId) return { success: false, message: "Missing item ID." };
    const sheet = getSheet(SHEETS.MENU);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No menu items were found." };
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const index = ids.findIndex(function(row) { return String(row[0] || "").trim() === itemId; });
    if (index < 0) return { success: false, message: "Menu item not found. Refresh the list and try again." };
    sheet.getRange(index + 2, 9).setValue("No");
    clearMenuCache();
    try {
      const cache = CacheService.getScriptCache();
      cache.remove("CACHE_POS_PAYLOAD");
      cache.remove("CACHE_MENU_PRICING");
    } catch (cacheError) { console.warn("Menu cache invalidation warning:", cacheError); }
    return { success: true, message: "Item removed from POS. Existing order history has been preserved." };
  } catch (error) {
    console.error("deactivateMenuItem failed:", error);
    return { success: false, message: error && error.message ? error.message : String(error) };
  } finally {
    lock.releaseLock();
  }
}


/** Save edits from the management table in one batch; inactive is a safe soft-delete. */
function saveMenuItemsBulk(items) {
  if (!Array.isArray(items) || !items.length) return { success: false, message: "No menu changes were supplied." };
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return { success: false, message: "Another menu update is in progress. Please try again." };
  try {
    const sheet = getSheet(SHEETS.MENU);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "The Menu sheet has no items." };
    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const rowById = new Map();
    values.forEach(function(row, i) { if (row[0] !== "") rowById.set(String(row[0]), i + 2); });
    const seen = new Set();
    items.forEach(function(item) {
      if (!item || !String(item.itemId || "").trim() || seen.has(String(item.itemId))) throw new Error("Invalid or duplicate menu item ID in request.");
      seen.add(String(item.itemId));
      if (!rowById.has(String(item.itemId))) throw new Error("Menu item not found: " + item.itemId);
      if (!String(item.category || "").trim()) throw new Error("Category is required for item " + item.itemId + ".");
      if (!Number.isFinite(Number(item.cost)) || Number(item.cost) < 0) throw new Error("Invalid cost price for item " + item.itemId + ".");
      if (!Number.isFinite(Number(item.price)) || Number(item.price) <= 0) throw new Error("Invalid selling price for item " + item.itemId + ".");
    });
    // Only change category, cost, selling price and active flag; keep item IDs/names and other columns intact.
    items.forEach(function(item) {
      const row = rowById.get(String(item.itemId));
      sheet.getRange(row, 3, 1, 3).setValues([[String(item.category).trim(), Number(item.cost), Number(item.price)]]);
      sheet.getRange(row, 9).setValue(item.isActive === true ? "Yes" : "No");
    });
    clearMenuCache();
    try { CacheService.getScriptCache().remove("CACHE_POS_PAYLOAD"); CacheService.getScriptCache().remove("CACHE_MENU_PRICING"); } catch (e) {}
    return { success: true, message: "Saved changes for " + items.length + " menu item(s)." };
  } catch (error) {
    console.error("saveMenuItemsBulk failed:", error);
    return { success: false, message: error && error.message ? error.message : String(error) };
  } finally { lock.releaseLock(); }
}
