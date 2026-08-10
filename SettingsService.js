/**
 * Load Application Settings (Optimized with CacheService)
 */
function loadAppSettings() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_app_settings_v1";

   // 1. Try to retrieve settings from server cache first
   const cachedData = cache.get(cacheKey);
   if (cachedData) {
      try {
         return JSON.parse(cachedData);
      } catch (e) {
         // Proceed to fetch from sheet if JSON parsing fails
      }
   }

   // 2. Fetch fresh data from the spreadsheet if cache misses
   const sheet = getSheet(SHEETS.SETTINGS);

   const settings = {
      deliveryAreas: sheet
         .getRange("V2:W100")
         .getValues()
         .filter(r => r[0] != "")
         .map(r => ({
            area: r[0],
            charge: Number(r[1])
         })),

      paymentModes: sheet
         .getRange("D2:D100")
         .getValues()
         .flat()
         .filter(String),

      paymentStatus: sheet
         .getRange("G2:G100")
         .getValues()
         .flat()
         .filter(String),

      orderTypes: sheet
         .getRange("M2:M100")
         .getValues()
         .flat()
         .filter(String),

      deliverySlots: sheet
         .getRange("P2:P100")
         .getValues()
         .flat()
         .filter(String)
   };

   // 3. Store the result in script cache for 6 hours (21600 seconds)
   try {
      cache.put(cacheKey, JSON.stringify(settings), 21600);
   } catch (e) {
      console.warn("Failed to write app settings to cache:", e);
   }

   return settings;
}

/**
 * Clear App Settings Cache 
 * (Tip: Call this helper function anywhere in your code if you ever update 
 * delivery areas or settings via an admin panel so changes reflect instantly)
 */
function clearAppSettingsCache() {
   try {
      const cache = CacheService.getScriptCache();
      cache.remove("smoky_swaad_app_settings_v1");
   } catch (e) {
      console.warn("Failed to clear settings cache:", e);
   }
}