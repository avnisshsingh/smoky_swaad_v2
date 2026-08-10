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

