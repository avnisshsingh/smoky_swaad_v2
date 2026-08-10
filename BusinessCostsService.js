/**
 * ==========================================
 * Get Business Costs (Optimized with Caching)
 * ==========================================
 */
function getBusinessCosts() {

    const cache = CacheService.getScriptCache();
    const cacheKey = "smoky_swaadv2_business_costs_cache_v1";

    // 1. Check if business costs are cached in server memory
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            // Fallback to sheet reading if JSON parse fails
        }
    }

    // 2. Original Sheet Reading & Mapping Logic
    const sheet = getSheet(SHEETS.BUSINESS_COSTS);
    const data = sheet.getDataRange().getValues();

    data.shift(); // Remove Header

    const businessCosts = [];
    let totalBusinessCost = 0;

    data.forEach(function(row, index) {

        const cost = {

            rowNumber: index + 2,

            costHead: row[0],

            amount: Number(row[1]) || 0,

            unit: row[2],

            active: row[3],

            remarks: row[4]

        };

        if (String(cost.active).trim().toUpperCase() === "YES") {

            totalBusinessCost += cost.amount;

        }

        businessCosts.push(cost);

    });

    const result = {

        success: true,

        totalBusinessCost: totalBusinessCost,

        businessCosts: businessCosts

    };

    // 3. Store result in cache for 6 hours (21600 seconds)
    try {
        cache.put(cacheKey, JSON.stringify(result), 21600);
    } catch (e) {
        console.warn("Failed to write business costs cache:", e);
    }

    return result;

}


/**
 * ==========================================
 * Save / Update Business Cost
 * ==========================================
 */
function saveBusinessCost(cost) {

    const sheet = getSheet(SHEETS.BUSINESS_COSTS);

    const data = sheet.getDataRange().getValues();

    const rowNumber = Number(cost.rowNumber);

    // Duplicate Check
    for (let i = 1; i < data.length; i++) {

        if (
            i + 1 !== rowNumber &&
            String(data[i][0]).trim().toUpperCase() ===
            cost.costHead.trim().toUpperCase()
        ) {

            return {

                success: false,

                message: "Cost Head already exists."

            };

        }

    }

    // ==========================================
    // UPDATE
    // ==========================================

    if (rowNumber > 0) {

        sheet.getRange(rowNumber, 1, 1, 5).setValues([[
            cost.costHead,
            cost.amount,
            cost.unit,
            cost.active,
            cost.remarks
        ]]);

    }

    // ==========================================
    // INSERT
    // ==========================================

    else {

        sheet.appendRow([
            cost.costHead,
            cost.amount,
            cost.unit,
            cost.active,
            cost.remarks
        ]);

    }

    // CRITICAL: Clear the business costs cache on save/update
    clearBusinessCostsCache();

    return {

        success: true

    };

}


/**
 * ==========================================
 * Get Business Cost By Row
 * ==========================================
 */
function getBusinessCost(rowNumber) {

    const sheet = getSheet(SHEETS.BUSINESS_COSTS);

    const row = sheet
        .getRange(rowNumber, 1, 1, 5)
        .getValues()[0];

    return {

        rowNumber: rowNumber,

        costHead: row[0],

        amount: Number(row[1]) || 0,

        unit: row[2],

        active: row[3],

        remarks: row[4]

    };

}


/**
 * ==========================================
 * Clear Business Costs Cache Helper
 * ==========================================
 */
function clearBusinessCostsCache() {
    try {
        const cache = CacheService.getScriptCache();
        cache.remove("smoky_swaadv2_business_costs_cache_v1");
    } catch (e) {
        console.warn("Failed to clear business costs cache:", e);
    }
}