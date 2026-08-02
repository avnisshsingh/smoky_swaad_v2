/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * PERSONAL EXPENSES SERVICE
 * ==========================================================
 *
 * Sheet:
 * PersonalExpenses
 *
 * Columns:
 *
 * A  Expense ID
 * B  Expense Date
 * C  Category
 * D  Subcategory
 * E  Description
 * F  Amount
 * G  Payment Mode
 * H  Paid To
 * I  Expense For
 * J  Nature
 * K  Recurring
 * L  Reference No
 * M  Remarks
 * N  Created At
 * O  Updated At
 * P  Status
 * Q  Unit
 * R  Quantity
 *
 * ==========================================================
 */


/**
 * ==========================================================
 * SAVE PERSONAL EXPENSE
 * ==========================================================
 */

/**
 * ==========================================================
 * PERSONAL EXPENSE PERFORMANCE CACHE
 * ==========================================================
 */
const PERSONAL_EXPENSE_CACHE_KEY = "PERSONAL_EXPENSE_BOOTSTRAP_V4";
const PERSONAL_EXPENSE_CACHE_SECONDS = 300;


function clearPersonalExpenseCache_() {

  try {
    CacheService
      .getScriptCache()
      .remove(
        PERSONAL_EXPENSE_CACHE_KEY
      );
  } catch (error) {
    console.warn(
      "Personal Expense cache clear warning:",
      error
    );
  }

}


/**
 * ==========================================================
 * GET PERSONAL EXPENSE BOOTSTRAP DATA
 * ==========================================================
 * One call supplies:
 * - dashboard summary
 * - active history
 * - units
 * - smart description memory
 *
 * Script Cache avoids repeated Sheet reads when reopening
 * the screen. Cache is invalidated after mutations.
 * ==========================================================
 */
function getPersonalExpenseBootstrapData() {

  try {

    const cache =
      CacheService.getScriptCache();

    const cached =
      cache.get(
        PERSONAL_EXPENSE_CACHE_KEY
      );

    if (cached) {
      return JSON.parse(cached);
    }

    const expenseSheet =
      getSheet(
        SHEETS.PERSONAL_EXPENSES
      );

    const settingsSheet =
      getSheet(
        SHEETS.SETTINGS
      );

    const expenseLastRow =
      expenseSheet.getLastRow();

    const rows =
      expenseLastRow >= 2
        ? expenseSheet
            .getRange(
              2,
              1,
              expenseLastRow - 1,
              18
            )
            .getValues()
        : [];

    // One Settings Q:S read supplies Personal Expense Categories + Units.
    const settingsLastRow =
      settingsSheet.getLastRow();

    let categories = [];
    let units = [];

    if (settingsLastRow >= 2) {

      const settingsData =
        settingsSheet
          .getRange(
            2,
            17,
            settingsLastRow - 1,
            3
          )
          .getValues();

      const seenCategories = {};
      const seenUnits = {};

      settingsData.forEach(function(row) {

        const category =
          String(row[0] || "").trim();

        const unit =
          String(row[2] || "").trim();

        if (category) {

          const categoryKey =
            category.toLowerCase();

          if (!seenCategories[categoryKey]) {
            seenCategories[categoryKey] = true;
            categories.push(category);
          }

        }

        if (unit) {

          const unitKey =
            unit.toLowerCase();

          if (!seenUnits[unitKey]) {
            seenUnits[unitKey] = true;
            units.push(unit);
          }

        }

      });

    }

    // Essential defaults are merged in-memory so older Settings
    // sheets still expose the complete requested category list.
    const defaultCategories = [
      "Grocery",
      "Fuel",
      "Baby",
      "Shopping",
      "Recharge",
      "House",
      "Food",
      "Medical",
      "Vehicle",
      "Bills",
      "Staff",
      "Transfer",
      "Education",
      "Entertainment",
      "ATM Cash",
      "Maintenance",
      "Naasta",
      "Vegetables",
      "Tea",
      "Xerox",
      "Fruits",
      "Milk"
    ];

    const categorySeen = {};

    categories.forEach(function(category) {
      categorySeen[
        category.toLowerCase()
      ] = true;
    });

    defaultCategories.forEach(function(category) {

      const key =
        category.toLowerCase();

      if (!categorySeen[key]) {
        categorySeen[key] = true;
        categories.push(category);
      }

    });

    categories.sort(function(a, b) {
      return a.localeCompare(
        b,
        undefined,
        {
          sensitivity: "base",
          numeric: true
        }
      );
    });

    const now =
      new Date();

    const todayKey =
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

    const monthKey =
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "yyyy-MM"
      );

    let todayTotal = 0;
    let monthTotal = 0;
    let monthCount = 0;

    const history = [];
    const latestByDescription = {};

    rows.forEach(function(row, index) {

      const status =
        String(row[15] || "").trim();

      if (
        status.toLowerCase() !== "active"
      ) {
        return;
      }

      const date =
        row[1] instanceof Date
          ? row[1]
          : new Date(row[1]);

      const validDate =
        !isNaN(date.getTime());

      const dateKey =
        validDate
          ? Utilities.formatDate(
              date,
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : "";

      const rowMonthKey =
        dateKey
          ? dateKey.substring(0, 7)
          : "";

      const amount =
        Number(row[5]) || 0;

      const nature =
        String(row[9] || "").trim();

      if (
        nature.toLowerCase() === "expense"
      ) {

        if (dateKey === todayKey) {
          todayTotal += amount;
        }

        if (rowMonthKey === monthKey) {
          monthTotal += amount;
          monthCount++;
        }

      }

      const description =
        String(row[4] || "").trim();

      const category =
        String(row[2] || "").trim();

      const unit =
        String(row[16] || "").trim();

      const quantity =
        row[17] === null ||
        row[17] === undefined
          ? ""
          : String(row[17]).trim();

      const updatedAt =
        row[14] instanceof Date &&
        !isNaN(row[14].getTime())
          ? row[14].getTime()
          : 0;

      const createdAt =
        row[13] instanceof Date &&
        !isNaN(row[13].getTime())
          ? row[13].getTime()
          : 0;

      const lastUsed =
        updatedAt ||
        createdAt ||
        index;

      if (
        description &&
        category
      ) {

        const key =
          description.toLowerCase();

        const existing =
          latestByDescription[key];

        if (
          !existing ||
          lastUsed > existing.lastUsed ||
          (
            lastUsed === existing.lastUsed &&
            index > existing.rowIndex
          )
        ) {

          latestByDescription[key] = {
            description: description,
            category: category,
            unit: unit,
            quantity: quantity,
            lastUsed: lastUsed,
            rowIndex: index
          };

        }

      }

      history.push({
        expenseId:
          String(row[0] || ""),
        expenseDate:
          validDate
            ? Utilities.formatDate(
                date,
                Session.getScriptTimeZone(),
                "dd-MM-yyyy"
              )
            : "",
        expenseDateKey:
          dateKey,
        category:
          category,
        subcategory:
          String(row[3] || ""),
        description:
          description,
        unit:
          unit,
        quantity:
          quantity,
        amount:
          amount,
        paymentMode:
          String(row[6] || ""),
        paidTo:
          String(row[7] || ""),
        expenseFor:
          String(row[8] || ""),
        nature:
          nature,
        recurring:
          String(row[10] || ""),
        referenceNo:
          String(row[11] || ""),
        remarks:
          String(row[12] || ""),
        createdAt:
          row[13] instanceof Date
            ? Utilities.formatDate(
                row[13],
                Session.getScriptTimeZone(),
                "dd-MM-yyyy HH:mm:ss"
              )
            : String(row[13] || ""),
        updatedAt:
          row[14] instanceof Date
            ? Utilities.formatDate(
                row[14],
                Session.getScriptTimeZone(),
                "dd-MM-yyyy HH:mm:ss"
              )
            : String(row[14] || ""),
        status:
          status
      });

    });

    history.sort(function(a, b) {

      return String(b.expenseDateKey || "")
        .localeCompare(
          String(a.expenseDateKey || "")
        );

    });

    const descriptions =
      Object.keys(latestByDescription)
        .map(function(key) {

          const item =
            latestByDescription[key];

          return {
            description:
              item.description,
            category:
              item.category,
            unit:
              item.unit,
            quantity:
              item.quantity,
            lastUsed:
              item.lastUsed
          };

        })
        .sort(function(a, b) {

          return Number(b.lastUsed || 0) -
                 Number(a.lastUsed || 0);

        });

    const response = {
      success: true,
      summary: {
        today: todayTotal,
        thisMonth: monthTotal,
        transactionCount: monthCount
      },
      history: history,
      categories: categories,
      units: units,
      descriptions: descriptions
    };

    // Cache only if payload is within CacheService practical limits.
    try {

      const json =
        JSON.stringify(response);

      if (json.length < 90000) {

        cache.put(
          PERSONAL_EXPENSE_CACHE_KEY,
          json,
          PERSONAL_EXPENSE_CACHE_SECONDS
        );

      }

    } catch (cacheError) {
      console.warn(
        "Personal Expense cache write warning:",
        cacheError
      );
    }

    return response;

  } catch (error) {

    console.error(
      "Personal Expense Bootstrap Error:",
      error
    );

    return {
      success: false,
      message:
        error && error.message
          ? error.message
          : String(error)
    };

  }

}


function savePersonalExpenseFromWeb(expenseData) {

  try {

    // ------------------------------------------
    // Validate
    // ------------------------------------------

    validatePersonalExpenseData(expenseData);


    // ------------------------------------------
    // Get Sheet
    // ------------------------------------------

    const expenseSheet =
      getSheet(SHEETS.PERSONAL_EXPENSES);


    // ------------------------------------------
    // Generate Expense ID
    // ------------------------------------------

    const expenseId =
      generatePersonalExpenseId(expenseSheet);


    // ------------------------------------------
    // Build Expense Date
    // ------------------------------------------

    const expenseDate =
      buildPersonalExpenseDate(
        expenseData.expenseDate
      );


    // ------------------------------------------
    // System timestamps
    // ------------------------------------------

    const now = new Date();


    // ------------------------------------------
    // Prepare row A:R
    // ------------------------------------------

    const row = [[

      expenseId,                              // A Expense ID

      expenseDate,                            // B Expense Date

      cleanPersonalExpenseText(
        expenseData.category
      ),                                      // C Category

      cleanPersonalExpenseText(
        expenseData.subcategory
      ),                                      // D Subcategory

      cleanPersonalExpenseText(
        expenseData.description
      ),                                      // E Description

      Number(expenseData.amount),             // F Amount

      cleanPersonalExpenseText(
        expenseData.paymentMode
      ),                                      // G Payment Mode

      cleanPersonalExpenseText(
        expenseData.paidTo
      ),                                      // H Paid To

      cleanPersonalExpenseText(
        expenseData.expenseFor
      ) || "Self",                            // I Expense For

      cleanPersonalExpenseText(
        expenseData.nature
      ) || "Expense",                         // J Nature

      cleanPersonalExpenseText(
        expenseData.recurring
      ) || "No",                              // K Recurring

      cleanPersonalExpenseText(
        expenseData.referenceNo
      ),                                      // L Reference No

      cleanPersonalExpenseText(
        expenseData.remarks
      ),                                      // M Remarks

      now,                                    // N Created At

      now,                                    // O Updated At

      "Active",                               // P Status

      cleanPersonalExpenseText(
        expenseData.unit
      ),                                       // Q Unit

      cleanPersonalExpenseQuantity(
        expenseData.quantity
      )                                        // R Quantity

    ]];


    // ------------------------------------------
    // Save in one write
    // ------------------------------------------

    const nextRow =
      expenseSheet.getLastRow() + 1;


    expenseSheet
      .getRange(
        nextRow,
        1,
        1,
        18
      )
      .setValues(row);


    // ------------------------------------------
    // Apply formats to new row
    // ------------------------------------------

    expenseSheet
      .getRange(nextRow, 2)
      .setNumberFormat("dd/MM/yyyy");


    expenseSheet
      .getRange(nextRow, 6)
      .setNumberFormat("₹#,##0.00");


    expenseSheet
      .getRange(nextRow, 14, 1, 2)
      .setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );


    // ------------------------------------------
    // Success
    // ------------------------------------------

    clearPersonalExpenseCache_();

    return {

      success: true,

      expenseId: expenseId,

      message:
        "Expense Saved Successfully"

    };


  } catch (error) {

    console.error(
      "Save Personal Expense Error:",
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : String(error)

    };

  }

}


/**
 * ==========================================================
 * VALIDATE PERSONAL EXPENSE
 * ==========================================================
 */
function validatePersonalExpenseData(expenseData) {

  if (
    !expenseData ||
    typeof expenseData !== "object"
  ) {

    throw new Error(
      "Expense data is missing."
    );

  }


  // ------------------------------------------
  // Expense Date
  // ------------------------------------------

  if (
    !String(
      expenseData.expenseDate || ""
    ).trim()
  ) {

    throw new Error(
      "Expense Date is required."
    );

  }


  // ------------------------------------------
  // Validate Date Format
  // Expected from HTML:
  // YYYY-MM-DD
  // ------------------------------------------

  const dateText =
    String(
      expenseData.expenseDate
    ).trim();


  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(
      dateText
    )
  ) {

    throw new Error(
      "Invalid Expense Date."
    );

  }


  // ------------------------------------------
  // Category
  // ------------------------------------------

  if (
    !String(
      expenseData.category || ""
    ).trim()
  ) {

    throw new Error(
      "Category is required."
    );

  }


  // ------------------------------------------
  // Description
  // ------------------------------------------

  if (
    !String(
      expenseData.description || ""
    ).trim()
  ) {

    throw new Error(
      "Description is required."
    );

  }


  // ------------------------------------------
  // Amount
  // ------------------------------------------

  const amount =
    Number(expenseData.amount);


  if (
    !Number.isFinite(amount) ||
    amount <= 0
  ) {

    throw new Error(
      "Amount should be greater than zero."
    );

  }


  // ------------------------------------------
  // Payment Mode
  // ------------------------------------------

  if (
    !String(
      expenseData.paymentMode || ""
    ).trim()
  ) {

    throw new Error(
      "Payment Mode is required."
    );

  }


  // ------------------------------------------
  // Nature
  // ------------------------------------------

  const allowedNature = [

    "Expense",
    "Transfer",
    "Refund"

  ];


  const nature =
    String(
      expenseData.nature || "Expense"
    ).trim();


  if (
    !allowedNature.includes(nature)
  ) {

    throw new Error(
      "Invalid transaction nature."
    );

  }


  // ------------------------------------------
  // Recurring
  // ------------------------------------------

  const recurring =
    String(
      expenseData.recurring || "No"
    ).trim();


  if (
    recurring !== "Yes" &&
    recurring !== "No"
  ) {

    throw new Error(
      "Invalid recurring value."
    );

  }


  return true;

}


/**
 * ==========================================================
 * GENERATE PERSONAL EXPENSE ID
 * ==========================================================
 *
 * Examples:
 *
 * EXP00001
 * EXP00002
 * EXP00003
 *
 * This searches all existing Expense IDs and
 * finds the highest number.
 *
 * Therefore sorting the PersonalExpenses sheet
 * does not break Expense ID generation.
 *
 * Soft-deleted records also retain their IDs.
 *
 * ==========================================================
 */
function generatePersonalExpenseId(
  expenseSheet
) {

  const lastRow =
    expenseSheet.getLastRow();


  // ------------------------------------------
  // No expense rows yet
  // ------------------------------------------

  if (lastRow <= 1) {

    return "EXP00001";

  }


  // ------------------------------------------
  // Read Expense IDs
  // Column A
  // ------------------------------------------

  const ids =
    expenseSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  let highestNumber = 0;


  ids.forEach(function(row) {

    const id =
      String(
        row[0] || ""
      ).trim();


    // Accept only proper Expense IDs
    // such as EXP00001

    const match =
      id.match(/^EXP(\d+)$/i);


    if (!match) {

      return;

    }


    const number =
      Number(match[1]);


    if (
      Number.isFinite(number) &&
      number > highestNumber
    ) {

      highestNumber =
        number;

    }

  });


  // ------------------------------------------
  // Generate next ID
  // ------------------------------------------

  const nextNumber =
    highestNumber + 1;


  return (
    "EXP" +
    String(nextNumber)
      .padStart(5, "0")
  );

}


/**
 * ==========================================================
 * BUILD PERSONAL EXPENSE DATE
 * ==========================================================
 *
 * HTML date field sends:
 *
 * YYYY-MM-DD
 *
 * We create the Date at 12:00 PM instead of
 * midnight to reduce timezone/date-shift issues.
 *
 * ==========================================================
 */
function buildPersonalExpenseDate(
  dateText
) {

  const parts =
    String(dateText)
      .trim()
      .split("-");


  if (parts.length !== 3) {

    throw new Error(
      "Invalid Expense Date."
    );

  }


  const year =
    Number(parts[0]);

  const month =
    Number(parts[1]);

  const day =
    Number(parts[2]);


  // ------------------------------------------
  // Basic numeric validation
  // ------------------------------------------

  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {

    throw new Error(
      "Invalid Expense Date."
    );

  }


  const expenseDate =
    new Date(
      year,
      month - 1,
      day,
      12,
      0,
      0
    );


  // ------------------------------------------
  // Catch impossible dates
  //
  // Example:
  // 2026-02-31
  // ------------------------------------------

  if (
    expenseDate.getFullYear() !== year ||
    expenseDate.getMonth() !==
      month - 1 ||
    expenseDate.getDate() !== day
  ) {

    throw new Error(
      "Invalid Expense Date."
    );

  }


  return expenseDate;

}


/**
 * ==========================================================
 * CLEAN PERSONAL EXPENSE TEXT
 * ==========================================================
 */
function cleanPersonalExpenseText(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(value).trim();

}


/**
 * ==========================================================
 * FIND PERSONAL EXPENSE ROW BY ID
 * ==========================================================
 *
 * This helper is intentionally included now.
 *
 * Edit / Delete / Restore will later locate
 * transactions using Expense ID rather than
 * depending on physical Google Sheet row numbers.
 *
 * ==========================================================
 */
function findPersonalExpenseRowById(
  expenseSheet,
  expenseId
) {

  const lastRow =
    expenseSheet.getLastRow();


  if (lastRow <= 1) {

    return -1;

  }


  const searchId =
    String(
      expenseId || ""
    )
      .trim()
      .toUpperCase();


  if (!searchId) {

    return -1;

  }


  // ------------------------------------------
  // Search Column A
  // ------------------------------------------

  const match =
    expenseSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .createTextFinder(searchId)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();


  if (!match) {

    return -1;

  }


  return match.getRow();

}







/**
 * ==========================================================
 * GET PERSONAL EXPENSES DASHBOARD DATA
 * ==========================================================
 *
 * Returns:
 * - Today's spending
 * - This month's spending
 * - This month's transaction count
 * - Recent active transactions
 *
 * IMPORTANT:
 * Only:
 *   Status = Active
 *   Nature = Expense
 *
 * are included in spending totals.
 *
 * Transfers and Refunds are stored but are NOT counted
 * as personal spending.
 *
 * ==========================================================
 */
function getPersonalExpensesDashboardData() {

  try {

    const sheet =
      getSheet(SHEETS.PERSONAL_EXPENSES);

    const lastRow =
      sheet.getLastRow();

    // ------------------------------------------
    // Empty Sheet
    // ------------------------------------------

    if (lastRow <= 1) {
return {

        success: true,

        summary: {
          today: 0,
          thisMonth: 0,
          transactionCount: 0
      },

        expenses: []

      };

    }


    // ------------------------------------------
    // Read A:P in one operation
    // ------------------------------------------

    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          18
        )
        .getValues();


    const now =
      new Date();


    const todayYear =
      now.getFullYear();

    const todayMonth =
      now.getMonth();

    const todayDate =
      now.getDate();


    let todayTotal = 0;

    let monthTotal = 0;

    let monthTransactionCount = 0;

    const expenses = [];


    // ------------------------------------------
    // Process transactions
    // ------------------------------------------

    data.forEach(function(row) {

      const expenseId =
        String(row[0] || "").trim();

      const expenseDate =
        row[1];

      const category =
        String(row[2] || "").trim();

      const subcategory =
        String(row[3] || "").trim();

      const description =
        String(row[4] || "").trim();

      const amount =
        Number(row[5]) || 0;

      const paymentMode =
        String(row[6] || "").trim();

      const paidTo =
        String(row[7] || "").trim();

      const expenseFor =
        String(row[8] || "").trim();

      const nature =
        String(row[9] || "").trim();

      const recurring =
        String(row[10] || "").trim();

      const referenceNo =
        String(row[11] || "").trim();

      const remarks =
        String(row[12] || "").trim();

      const createdAt =
        row[13];

      const updatedAt =
        row[14];

      const status =
        String(row[15] || "").trim();

      const unit =
        String(row[16] || "").trim();

      const quantity =
        row[17] === null ||
        row[17] === undefined
          ? ""
          : String(row[17]).trim();


      // ------------------------------------------
      // Skip invalid rows
      // ------------------------------------------

      if (!expenseId) {
        return;
      }


      // ------------------------------------------
      // Only Active transactions appear
      // in normal Recent Expenses
      // ------------------------------------------

      if (
        status.toLowerCase() !==
        "active"
      ) {

        return;

      }


      // ------------------------------------------
      // Validate Expense Date
      // ------------------------------------------

      const dateObject =
        expenseDate instanceof Date
          ? expenseDate
          : new Date(expenseDate);


      if (
        isNaN(
          dateObject.getTime()
        )
      ) {

        return;

      }


      const expenseYear =
        dateObject.getFullYear();

      const expenseMonth =
        dateObject.getMonth();

      const expenseDay =
        dateObject.getDate();


      const isThisMonth =
        expenseYear === todayYear &&
        expenseMonth === todayMonth;


      const isToday =
        isThisMonth &&
        expenseDay === todayDate;


      // ------------------------------------------
      // SUMMARY CALCULATIONS
      //
      // Only actual Expense transactions
      // contribute to spending.
      //
      // Transfer and Refund are excluded.
      // ------------------------------------------

      if (
        nature.toLowerCase() ===
        "expense"
      ) {

        if (isToday) {

          todayTotal +=
            amount;

        }


        if (isThisMonth) {

          monthTotal +=
            amount;

          monthTransactionCount++;

        }

      }


      // ------------------------------------------
      // Add Active transaction to list
      // ------------------------------------------

      expenses.push({

        expenseId:
          expenseId,

        expenseDate:
          Utilities.formatDate(
            dateObject,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          ),

        expenseDateRaw:
          Utilities.formatDate(
            dateObject,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          ),

        category:
          category,

        subcategory:
          subcategory,

        description:
          description,

        unit:
          unit,

        quantity:
          quantity,

        amount:
          amount,

        paymentMode:
          paymentMode,

        paidTo:
          paidTo,

        expenseFor:
          expenseFor,

        nature:
          nature,

        recurring:
          recurring,

        referenceNo:
          referenceNo,

        remarks:
          remarks,

        createdAt:
          createdAt
            ? Utilities.formatDate(
                new Date(createdAt),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy HH:mm:ss"
              )
            : "",

        updatedAt:
          updatedAt
            ? Utilities.formatDate(
                new Date(updatedAt),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy HH:mm:ss"
              )
            : "",

        status:
          status

      });

    });


    // ------------------------------------------
    // Latest transactions first
    //
    // We use Expense ID as a secondary sort
    // because EXP numbers increase over time.
    // ------------------------------------------

    expenses.sort(function(a, b) {

      if (
        a.expenseDateRaw !==
        b.expenseDateRaw
      ) {

        return (
          b.expenseDateRaw >
          a.expenseDateRaw
            ? 1
            : -1
        );

      }


      return (
        b.expenseId >
        a.expenseId
          ? 1
          : -1
      );

    });


    // ------------------------------------------
    // Limit Recent Expenses
    //
    // Keeps mobile loading fast as the sheet grows.
    // ------------------------------------------

    const recentExpenses =
      expenses.slice(0, 50);


    return {

      success: true,

      summary: {

        today:
          todayTotal,

        thisMonth:
          monthTotal,

        transactionCount:
          monthTransactionCount

      },

      expenses:
        recentExpenses

    };


  } catch (error) {

    console.error(
      "Get Personal Expenses Dashboard Error:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}














/**
 * ==========================================================
 * UPDATE PERSONAL EXPENSE
 * ==========================================================
 */
function updatePersonalExpenseFromWeb(expenseId, expenseData) {

  try {

    expenseId =
      String(expenseId || "")
        .trim()
        .toUpperCase();


    if (!expenseId) {

      throw new Error(
        "Expense ID is required."
      );

    }


    validatePersonalExpenseData(
      expenseData
    );


    const sheet =
      getSheet(
        SHEETS.PERSONAL_EXPENSES
      );


    const rowNumber =
      findPersonalExpenseRowById(
        sheet,
        expenseId
      );


    if (rowNumber === -1) {

      throw new Error(
        "Expense not found: " +
        expenseId
      );

    }


    // ------------------------------------------
    // Read existing A:R
    // ------------------------------------------

    const existingRow =
      sheet
        .getRange(
          rowNumber,
          1,
          1,
          18
        )
        .getValues()[0];


    const status =
      String(
        existingRow[15] || ""
      ).trim();


    if (
      status.toLowerCase() ===
      "deleted"
    ) {

      throw new Error(
        "Deleted expense cannot be edited. Restore it first."
      );

    }


    // ------------------------------------------
    // Preserve Created At
    // ------------------------------------------

    const createdAt =
      existingRow[13];


    const expenseDate =
      buildPersonalExpenseDate(
        expenseData.expenseDate
      );


    const now =
      new Date();


    // ------------------------------------------
    // Build complete replacement A:R
    // ------------------------------------------

    const updatedRow = [[

      expenseId,                              // A

      expenseDate,                            // B

      cleanPersonalExpenseText(
        expenseData.category
      ),                                      // C

      cleanPersonalExpenseText(
        expenseData.subcategory
      ),                                      // D

      cleanPersonalExpenseText(
        expenseData.description
      ),                                      // E

      Number(
        expenseData.amount
      ),                                      // F

      cleanPersonalExpenseText(
        expenseData.paymentMode
      ),                                      // G

      cleanPersonalExpenseText(
        expenseData.paidTo
      ),                                      // H

      cleanPersonalExpenseText(
        expenseData.expenseFor
      ) || "Self",                            // I

      cleanPersonalExpenseText(
        expenseData.nature
      ) || "Expense",                         // J

      cleanPersonalExpenseText(
        expenseData.recurring
      ) || "No",                              // K

      cleanPersonalExpenseText(
        expenseData.referenceNo
      ),                                      // L

      cleanPersonalExpenseText(
        expenseData.remarks
      ),                                      // M

      createdAt,                              // N

      now,                                    // O

      "Active",                               // P

      cleanPersonalExpenseText(
        expenseData.unit
      ),                                       // Q

      cleanPersonalExpenseQuantity(
        expenseData.quantity
      )                                        // R

    ]];


    sheet
      .getRange(
        rowNumber,
        1,
        1,
        18
      )
      .setValues(
        updatedRow
      );


    sheet
      .getRange(
        rowNumber,
        2
      )
      .setNumberFormat(
        "dd/MM/yyyy"
      );


    sheet
      .getRange(
        rowNumber,
        6
      )
      .setNumberFormat(
        "₹#,##0.00"
      );


    sheet
      .getRange(
        rowNumber,
        14,
        1,
        2
      )
      .setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );


    clearPersonalExpenseCache_();

    return {

      success: true,

      expenseId:
        expenseId,

      message:
        "Expense Updated Successfully"

    };


  } catch (error) {

    console.error(
      "Update Personal Expense Error:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}


/**
 * ==========================================================
 * SOFT DELETE PERSONAL EXPENSE
 * ==========================================================
 */
function deletePersonalExpenseFromWeb(
  expenseId
) {

  try {

    expenseId =
      String(expenseId || "")
        .trim()
        .toUpperCase();


    if (!expenseId) {

      throw new Error(
        "Expense ID is required."
      );

    }


    const sheet =
      getSheet(
        SHEETS.PERSONAL_EXPENSES
      );


    const rowNumber =
      findPersonalExpenseRowById(
        sheet,
        expenseId
      );


    if (rowNumber === -1) {

      throw new Error(
        "Expense not found: " +
        expenseId
      );

    }


    // ------------------------------------------
    // Read current status
    // ------------------------------------------

    const currentStatus =
      String(
        sheet
          .getRange(
            rowNumber,
            16
          )
          .getValue() || ""
      ).trim();


    if (
      currentStatus.toLowerCase() ===
      "deleted"
    ) {

      return {

        success: true,

        expenseId:
          expenseId,

        message:
          "Expense is already deleted."

      };

    }


    const now =
      new Date();


    // ------------------------------------------
    // O = Updated At
    // P = Status
    //
    // ONE batch write
    // ------------------------------------------

    sheet
      .getRange(
        rowNumber,
        15,
        1,
        2
      )
      .setValues([[

        now,

        "Deleted"

      ]]);


    sheet
      .getRange(
        rowNumber,
        15
      )
      .setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );


    clearPersonalExpenseCache_();

    return {

      success: true,

      expenseId:
        expenseId,

      message:
        "Expense Deleted Successfully"

    };


  } catch (error) {

    console.error(
      "Delete Personal Expense Error:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}


/**
 * ==========================================================
 * RESTORE PERSONAL EXPENSE
 * ==========================================================
 */
function restorePersonalExpenseFromWeb(
  expenseId
) {

  try {

    expenseId =
      String(expenseId || "")
        .trim()
        .toUpperCase();


    if (!expenseId) {

      throw new Error(
        "Expense ID is required."
      );

    }


    const sheet =
      getSheet(
        SHEETS.PERSONAL_EXPENSES
      );


    const rowNumber =
      findPersonalExpenseRowById(
        sheet,
        expenseId
      );


    if (rowNumber === -1) {

      throw new Error(
        "Expense not found: " +
        expenseId
      );

    }


    const currentStatus =
      String(
        sheet
          .getRange(
            rowNumber,
            16
          )
          .getValue() || ""
      ).trim();


    if (
      currentStatus.toLowerCase() ===
      "active"
    ) {

      return {

        success: true,

        expenseId:
          expenseId,

        message:
          "Expense is already active."

      };

    }


    const now =
      new Date();


    sheet
      .getRange(
        rowNumber,
        15,
        1,
        2
      )
      .setValues([[

        now,

        "Active"

      ]]);


    sheet
      .getRange(
        rowNumber,
        15
      )
      .setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );


    clearPersonalExpenseCache_();

    return {

      success: true,

      expenseId:
        expenseId,

      message:
        "Expense Restored Successfully"

    };


  } catch (error) {

    console.error(
      "Restore Personal Expense Error:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}


/**
 * ==========================================================
 * GET DELETED PERSONAL EXPENSES
 * ==========================================================
 */
function getDeletedPersonalExpenses() {

  try {

    const sheet =
      getSheet(
        SHEETS.PERSONAL_EXPENSES
      );


    const lastRow =
      sheet.getLastRow();


    if (lastRow <= 1) {

      return {

        success: true,

        expenses: []

      };

    }


    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          18
        )
        .getValues();


    const expenses = [];


    data.forEach(
      function(row) {

        const status =
          String(
            row[15] || ""
          ).trim();


        if (
          status.toLowerCase() !==
          "deleted"
        ) {

          return;

        }


        const dateObject =
          row[1] instanceof Date
            ? row[1]
            : new Date(row[1]);


        expenses.push({

          expenseId:
            String(
              row[0] || ""
            ),

          expenseDate:
            isNaN(
              dateObject.getTime()
            )
              ? ""
              : Utilities.formatDate(
                  dateObject,
                  Session.getScriptTimeZone(),
                  "dd/MM/yyyy"
                ),

          expenseDateRaw:
            isNaN(
              dateObject.getTime()
            )
              ? ""
              : Utilities.formatDate(
                  dateObject,
                  Session.getScriptTimeZone(),
                  "yyyy-MM-dd"
                ),

          category:
            String(
              row[2] || ""
            ),

          subcategory:
            String(
              row[3] || ""
            ),

          description:
            String(
              row[4] || ""
            ),

          unit:
            String(
              row[16] || ""
            ),

          quantity:
            row[17] === null ||
            row[17] === undefined
              ? ""
              : String(row[17]).trim(),

          amount:
            Number(
              row[5]
            ) || 0,

          paymentMode:
            String(
              row[6] || ""
            ),

          paidTo:
            String(
              row[7] || ""
            ),

          expenseFor:
            String(
              row[8] || ""
            ),

          nature:
            String(
              row[9] || ""
            ),

          recurring:
            String(
              row[10] || ""
            ),

          referenceNo:
            String(
              row[11] || ""
            ),

          remarks:
            String(
              row[12] || ""
            ),

          status:
            status

        });

      }
    );


    expenses.reverse();


    return {

      success: true,

      expenses:
        expenses.slice(0, 50)

    };


  } catch (error) {

    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}


/**
 * ==========================================================
 * GET PERSONAL EXPENSE HISTORY
 * ==========================================================
 *
 * Simple household-expense history.
 * Filters:
 * - month: YYYY-MM
 * - category
 * - search: description / paid to / reference / expense ID
 *
 * Only Active records are returned.
 * Maximum 100 matching rows are sent to the UI.
 * ==========================================================
 */
function getPersonalExpenseHistory(filters) {

  try {

    filters = filters || {};

    const monthFilter =
      String(filters.month || "").trim();

    const categoryFilter =
      String(filters.category || "").trim().toLowerCase();

    const searchFilter =
      String(filters.search || "").trim().toLowerCase();

    const sheet =
      getSheet(SHEETS.PERSONAL_EXPENSES);

    const lastRow =
      sheet.getLastRow();

    if (lastRow <= 1) {

      return {
        success: true,
        expenses: [],
        count: 0,
        total: 0
      };

    }

    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          18
        )
        .getValues();

    const matches = [];
    let total = 0;

    data.forEach(function(row) {

      const expenseId =
        String(row[0] || "").trim();

      const expenseDate =
        row[1];

      const category =
        String(row[2] || "").trim();

      const subcategory =
        String(row[3] || "").trim();

      const description =
        String(row[4] || "").trim();

      const amount =
        Number(row[5]) || 0;

      const paymentMode =
        String(row[6] || "").trim();

      const paidTo =
        String(row[7] || "").trim();

      const expenseFor =
        String(row[8] || "").trim();

      const nature =
        String(row[9] || "").trim();

      const recurring =
        String(row[10] || "").trim();

      const referenceNo =
        String(row[11] || "").trim();

      const remarks =
        String(row[12] || "").trim();

      const createdAt =
        row[13];

      const updatedAt =
        row[14];

      const status =
        String(row[15] || "").trim();

      const unit =
        String(row[16] || "").trim();

      const quantity =
        row[17] === null ||
        row[17] === undefined
          ? ""
          : String(row[17]).trim();

      if (
        !expenseId ||
        status.toLowerCase() !== "active"
      ) {
        return;
      }

      const dateObject =
        expenseDate instanceof Date
          ? expenseDate
          : new Date(expenseDate);

      if (isNaN(dateObject.getTime())) {
        return;
      }

      const rawDate =
        Utilities.formatDate(
          dateObject,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );

      if (
        monthFilter &&
        rawDate.slice(0, 7) !== monthFilter
      ) {
        return;
      }

      if (
        categoryFilter &&
        category.toLowerCase() !== categoryFilter
      ) {
        return;
      }

      if (searchFilter) {

        const searchableText = [
          expenseId,
          description,
          paidTo,
          referenceNo,
          subcategory
        ]
          .join(" ")
          .toLowerCase();

        if (
          searchableText.indexOf(searchFilter) === -1
        ) {
          return;
        }

      }

      /*
       * Total represents actual spending only.
       * Transfers / refunds remain visible in history
       * but do not inflate household spend.
       */
      if (
        nature.toLowerCase() === "expense"
      ) {
        total += amount;
      }

      matches.push({

        expenseId: expenseId,

        expenseDate:
          Utilities.formatDate(
            dateObject,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          ),

        expenseDateRaw: rawDate,

        category: category,
        subcategory: subcategory,
        description: description,
        unit: unit,
        quantity: quantity,
        amount: amount,
        paymentMode: paymentMode,
        paidTo: paidTo,
        expenseFor: expenseFor,
        nature: nature,
        recurring: recurring,
        referenceNo: referenceNo,
        remarks: remarks,

        createdAt:
          createdAt
            ? Utilities.formatDate(
                new Date(createdAt),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy HH:mm:ss"
              )
            : "",

        updatedAt:
          updatedAt
            ? Utilities.formatDate(
                new Date(updatedAt),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy HH:mm:ss"
              )
            : "",

        status: status

      });

    });

    matches.sort(function(a, b) {

      if (
        a.expenseDateRaw !== b.expenseDateRaw
      ) {
        return a.expenseDateRaw < b.expenseDateRaw
          ? 1
          : -1;
      }

      return a.expenseId < b.expenseId
        ? 1
        : -1;

    });

    const count =
      matches.length;

    return {

      success: true,

      expenses:
        matches.slice(0, 100),

      count: count,

      total: total

    };

  } catch (error) {

    console.error(
      "Get Personal Expense History Error:",
      error
    );

    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : String(error)

    };

  }

}



/**
 * ==========================================================
 * CLEAN PERSONAL EXPENSE QUANTITY
 * Optional, numeric and non-negative
 * ==========================================================
 */
function cleanPersonalExpenseQuantity(value) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ""
  ) {
    return "";
  }

  const quantity =
    Number(value);

  if (
    !isFinite(quantity) ||
    quantity < 0
  ) {
    throw new Error(
      "Quantity must be a valid non-negative number."
    );
  }

  return quantity;

}


/**
 * ==========================================================
 * GET PERSONAL EXPENSE SETUP DATA
 * ==========================================================
 * Units: Settings column S
 * Smart Description: most recent Category + Quantity + Unit wins
 * Performance: loaded once; client searches locally while typing
 * ==========================================================
 */
function getPersonalExpenseSetupData() {

  try {

    const settingsSheet =
      getSheet(SHEETS.SETTINGS);

    const expenseSheet =
      getSheet(SHEETS.PERSONAL_EXPENSES);

    // Units from Settings!S2:S
    const settingsLastRow =
      settingsSheet.getLastRow();

    let units = [];

    if (settingsLastRow >= 2) {

      units =
        settingsSheet
          .getRange(
            2,
            19,
            settingsLastRow - 1,
            1
          )
          .getValues()
          .flat()
          .map(function(value) {
            return String(value || "").trim();
          })
          .filter(function(value) {
            return value !== "";
          });

      const seenUnits = {};

      units =
        units.filter(function(unit) {

          const key =
            unit.toLowerCase();

          if (seenUnits[key]) {
            return false;
          }

          seenUnits[key] = true;
          return true;

        });

    }

    // Learn Description -> latest Category from Active records.
    const expenseLastRow =
      expenseSheet.getLastRow();

    const latestByDescription = {};

    if (expenseLastRow >= 2) {

      const rows =
        expenseSheet
          .getRange(
            2,
            1,
            expenseLastRow - 1,
            18
          )
          .getValues();

      rows.forEach(function(row, index) {

        const description =
          String(row[4] || "").trim();

        const category =
          String(row[2] || "").trim();

        const status =
          String(row[15] || "")
            .trim()
            .toLowerCase();

        if (
          !description ||
          !category ||
          status !== "active"
        ) {
          return;
        }

        const updatedAt =
          row[14] instanceof Date &&
          !isNaN(row[14].getTime())
            ? row[14].getTime()
            : 0;

        const createdAt =
          row[13] instanceof Date &&
          !isNaN(row[13].getTime())
            ? row[13].getTime()
            : 0;

        const time =
          updatedAt || createdAt || 0;

        const key =
          description.toLowerCase();

        const existing =
          latestByDescription[key];

        if (
          !existing ||
          time > existing.time ||
          (
            time === existing.time &&
            index > existing.rowIndex
          )
        ) {

          latestByDescription[key] = {
            description: description,
            category: category,
            unit: String(row[16] || "").trim(),
            quantity:
              row[17] === null ||
              row[17] === undefined
                ? ""
                : String(row[17]).trim(),
            time: time,
            rowIndex: index
          };

        }

      });

    }

    const descriptions =
      Object.keys(latestByDescription)
        .map(function(key) {

          return {
            description:
              latestByDescription[key].description,
            category:
              latestByDescription[key].category,
            unit:
              latestByDescription[key].unit,
            quantity:
              latestByDescription[key].quantity,
            lastUsed:
              latestByDescription[key].time
          };

        })
        .sort(function(a, b) {

          return Number(b.lastUsed || 0) -
                 Number(a.lastUsed || 0);

        });

    return {
      success: true,
      units: units,
      descriptions: descriptions
    };

  } catch (error) {

    console.error(
      "Get Personal Expense Setup Error:",
      error
    );

    return {
      success: false,
      message:
        error && error.message
          ? error.message
          : String(error)
    };

  }

}

/**
 * ==========================================================
 * ADD PERSONAL EXPENSE CATEGORY
 * ==========================================================
 * Persists a user-created category in Settings column Q.
 * Uses LockService to avoid duplicate concurrent writes.
 * ==========================================================
 */
function addPersonalExpenseCategory(category) {

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(5000);

    category =
      String(category || "")
        .trim()
        .replace(/\s+/g, " ");

    if (!category) {
      return {
        success: false,
        message: "Category is required."
      };
    }

    if (category.length > 50) {
      return {
        success: false,
        message: "Category is too long."
      };
    }

    const sheet =
      getSheet(SHEETS.SETTINGS);

    const lastRow =
      Math.max(
        sheet.getLastRow(),
        2
      );

    const values =
      sheet
        .getRange(
          2,
          17,
          Math.max(lastRow - 1, 1),
          1
        )
        .getValues();

    let firstBlankRow = -1;
    let existingCategory = "";

    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      const current =
        String(values[i][0] || "").trim();

      if (
        !current &&
        firstBlankRow === -1
      ) {
        firstBlankRow = i + 2;
      }

      if (
        current &&
        current.toLowerCase() ===
          category.toLowerCase()
      ) {
        existingCategory = current;
        break;
      }

    }

    if (!existingCategory) {

      const targetRow =
        firstBlankRow !== -1
          ? firstBlankRow
          : lastRow + 1;

      sheet
        .getRange(
          targetRow,
          17
        )
        .setValue(category);

    } else {

      category =
        existingCategory;

    }

    clearPersonalExpenseCache_();

    // Return the fresh merged list from bootstrap.
    const bootstrap =
      getPersonalExpenseBootstrapData();

    return {
      success: true,
      category: category,
      categories:
        bootstrap.categories || []
    };

  } catch (error) {

    return {
      success: false,
      message:
        error && error.message
          ? error.message
          : String(error)
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (ignore) {}

  }

}


