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
 *
 * ==========================================================
 */


/**
 * ==========================================================
 * SAVE PERSONAL EXPENSE
 * ==========================================================
 */
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
    // Prepare row A:P
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

      "Active"                                // P Status

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
        16
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
          16
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
    // Read existing A:P
    // ------------------------------------------

    const existingRow =
      sheet
        .getRange(
          rowNumber,
          1,
          1,
          16
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
    // Build complete replacement A:P
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

      "Active"                                // P

    ]];


    sheet
      .getRange(
        rowNumber,
        1,
        1,
        16
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
          16
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