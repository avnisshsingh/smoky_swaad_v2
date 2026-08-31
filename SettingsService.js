/**
 * Load Application Settings (Optimized with CacheService)
 */
function loadAppSettings() {
   const cache = CacheService.getScriptCache();
   const cacheKey = "smoky_swaad_app_settings_cache_v7"; // Version bumped to clear old cache
   
   const cached = cache.get(cacheKey);
   if (cached) {
      try {
         return JSON.parse(cached);
      } catch (e) {}
   }

   const sheet = getSheet(SHEETS.SETTINGS);
   const lastRow = sheet.getLastRow();
   
   const settings = {
      deliveryAreas: [],
      paymentModes: ["Cash", "PhonePe", "Google Pay", "Paytm", "Due"],
      paymentStatus: ["Pending", "Paid", "Partial"],
      orderTypes: ["Delivery", "Pickup", "Dining"],
      deliverySlots: ["Lunch", "Evening", "Dinner"]
   };

   if (lastRow >= 2) {
      // Column V is 22 (Delivery Areas), Column W is 23 (Delivery Charges)
      const data = sheet.getRange(2, 22, lastRow - 1, 2).getValues();
      
      const areas = [];
      data.forEach(function(row) {
         const areaName = String(row[0] || "").trim();
         const charge = Number(row[1] || 0);
         if (areaName) {
            areas.push({ area: areaName, charge: charge });
         }
      });
      if (areas.length > 0) {
         settings.deliveryAreas = areas;
      }
   }

   try {
      cache.put(cacheKey, JSON.stringify(settings), 7200);
   } catch (e) {}

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