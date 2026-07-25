
/**
 * ==========================================
 * Get Business Costs
 * ==========================================
 */
function getBusinessCosts() {

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

    return {

        success: true,

        totalBusinessCost: totalBusinessCost,

        businessCosts: businessCosts

    };

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