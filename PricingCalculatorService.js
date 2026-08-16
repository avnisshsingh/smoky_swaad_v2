/**
 * ==========================================
 * Get Latest Ingredient Purchase
 * ==========================================
 */
function getLatestIngredientPurchase(ingredientName) {

    const sheet = getSheet(SHEETS.PURCHASE_REGISTER);

    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {

        if (
            String(data[i][2]).trim().toUpperCase() ===
            String(ingredientName).trim().toUpperCase()
        ) {

            return {

                success: true,

                purchaseQty: Number(data[i][3]) || 0,

                purchaseUnit: data[i][4],

                purchasePrice: Number(data[i][6]) || 0

            };

        }

    }

    return {

        success: false,

        message: "Ingredient not found."

    };

}




/**
 * ==========================================
 * Get Menu Items (Optimized with Cache)
 * ==========================================
 */
function getMenuItems() {
    return getCachedData("CACHE_MENU_PRICING", function() {
        const sheet = getSheet(SHEETS.MENU);
        const data = sheet.getDataRange().getValues();
        data.shift();

        const menuItems = [];

        data.forEach(function(row){
            if (!row[1]) {
                return;
            }
            menuItems.push({
                menuItem: row[1],
                category: row[2],
                sellingPrice: Number(row[4]) || 0
            });
        });

        return {
            success: true,
            menuItems: menuItems
        };
    });
}





/**
 * ==========================================
 * Get Ingredients (Live from Purchase Register)
 * ==========================================
 */
function getIngredients() {
    try {
        const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
        const data = sheet.getDataRange().getValues();
        data.shift(); // Remove header row

        const ingredientMap = {};

        data.forEach(function(row){
            const ingredient = String(row[2]).trim();
            if (!ingredient) {
                return;
            }
            const key = ingredient.toUpperCase();
            if (!ingredientMap[key]) {
                ingredientMap[key] = ingredient;
            }
        });

        return {
            success: true,
            ingredients: Object.values(ingredientMap).sort()
        };
    } catch (error) {
        console.error("Error fetching ingredients: " + error.toString());
        return {
            success: false,
            ingredients: []
        };
    }
}







/**
 * ==========================================
 * Get Latest Purchase
 * ==========================================
 */
function getLatestPurchase(itemName) {

    const sheet = getSheet(SHEETS.PURCHASE_REGISTER);

    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {

        if (

            String(data[i][2]).trim().toUpperCase() ===

            String(itemName).trim().toUpperCase()

        ) {

            return {

                success: true,

                quantity: Number(data[i][3]) || 0,

                unit: data[i][4],

                amount: Number(data[i][6]) || 0

            };

        }

    }

    return {

        success: false,

        message: "Ingredient not found."

    };

}











/**
 * ==========================================
 * Save Recipe & Pricing Data to Sheet (Enhanced)
 * ==========================================
 */
function saveRecipeData(payload) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        
        // 1. Find or create the Recipes tab
        let recipeSheet = ss.getSheetByName("Recipes");
        if (!recipeSheet) {
            recipeSheet = ss.insertSheet("Recipes");
            recipeSheet.appendRow([
                "Timestamp", "Menu Item", "Category", "Ingredient Name", 
                "Quantity", "Unit", "Rate", "Wastage %", "Line Cost", 
                "Ingredient Total", "Packaging Cost", "Business Cost", "Total Cost", 
                "Selling Price", "Net Profit", "Profit %", "Food Cost %"
            ]);
        }

        const menuItem = payload.menuItem;
        const timestamp = new Date();

        // 2. Remove old entries for this Menu Item to prevent duplicates on update
        const dataRange = recipeSheet.getDataRange();
        const values = dataRange.getValues();
        
        for (let i = values.length - 1; i >= 1; i--) {
            if (String(values[i][1]).trim().toUpperCase() === String(menuItem).trim().toUpperCase()) {
                recipeSheet.deleteRow(i + 1);
            }
        }

        // 3. Append new ingredient breakdown rows with enhanced metrics
        payload.ingredients.forEach(function(item) {
            recipeSheet.appendRow([
                timestamp,
                payload.menuItem,
                payload.category,
                item.name,
                item.qty,
                item.unit,
                item.rate,
                item.wastage || 0,
                item.cost,
                payload.ingredientTotal,
                payload.packagingCost || 0,
                payload.businessCost,
                payload.totalCost,
                payload.sellingPrice,
                payload.profit,
                payload.profitPercentage,
                payload.foodCostPercentage
            ]);
        });

        // 4. Invalidate cache
        const cache = CacheService.getScriptCache();
        cache.remove("CACHE_MENU_PRICING");

        return {
            success: true,
            message: "Recipe saved successfully with packaging and cost metrics!"
        };

    } catch (error) {
        console.error("Save Recipe Error: " + error.toString());
        return {
            success: false,
            message: error.toString()
        };
    }
}