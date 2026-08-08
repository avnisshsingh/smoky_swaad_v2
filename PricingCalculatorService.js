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
 * Get Ingredients (Optimized with Cache)
 * ==========================================
 */
function getIngredients() {
    return getCachedData("CACHE_UNIQUE_INGREDIENTS", function() {
        const sheet = getSheet(SHEETS.PURCHASE_REGISTER);
        const data = sheet.getDataRange().getValues();
        data.shift();

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
    });
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