/**
 * ==========================================
 * Convert Quantity to Base Unit
 * ==========================================
 *
 * Weight  -> Gram
 * Volume  -> ml
 * Count   -> Unit
 *
 */
function convertToBaseUnit(quantity, unit) {

    quantity = Number(quantity) || 0;

    switch (String(unit).trim().toUpperCase()) {

        // -----------------------------
        // Weight
        // -----------------------------

        case "KG":
            return quantity * 1000;

        case "GRAM":
        case "G":
            return quantity;

        // -----------------------------
        // Volume
        // -----------------------------

        case "LITRE":
        case "LITER":
        case "L":
            return quantity * 1000;

        case "ML":
            return quantity;

        // -----------------------------
        // Count
        // -----------------------------

        case "UNIT":
        case "NOS":
        case "NO":
        case "PIECE":
        case "PCS":
            return quantity;

        default:
            return quantity;

    }

}



/**
 * ==========================================
 * Get Base Unit
 * ==========================================
 */
function getBaseUnit(unit) {

    switch (String(unit).trim().toUpperCase()) {

        case "KG":
        case "GRAM":
        case "G":
            return "Gram";

        case "LITRE":
        case "LITER":
        case "L":
        case "ML":
            return "ml";

        default:
            return "Unit";

    }

}




/**
 * ==========================================
 * Calculate Rate Per Base Unit
 * ==========================================
 */
function calculateRate(amount, quantity, unit) {

    const baseQty = convertToBaseUnit(quantity, unit);

    if (baseQty <= 0) {

        return 0;

    }

    return amount / baseQty;

}


/**
 * ==========================================
 * Calculate Ingredient Cost
 * ==========================================
 */
function calculateIngredientCost(ratePerBaseUnit, recipeQty, recipeUnit) {

    const recipeBaseQty =
        convertToBaseUnit(recipeQty, recipeUnit);

    return ratePerBaseUnit * recipeBaseQty;

}
