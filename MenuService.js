


/**
 * ==========================================
 * Menu Service
 * ==========================================
 */
/**
 * Search Menu Items
 */
function searchMenu(keyword) {

  if (!keyword) return [];

  const menu = getSheet(SHEETS.MENU);

  const lastRow = menu.getLastRow();

  if (lastRow < 2) return [];

  const data = menu
      .getRange(2, 1, lastRow - 1, 9)
      .getValues();

  keyword = keyword.toLowerCase();

  return data
      .filter(row =>

          row[8] === "Yes" &&

          String(row[1]).toLowerCase().includes(keyword)

      )

      .map(row => ({

          itemId: row[0],

          itemName: row[1],

          category: row[2],

          cost: Number(row[3]),

          price: Number(row[4]),

          prepTime: row[7]

      }))

      .slice(0,8);

}

/**
 * Load Complete Active Menu
 */
function loadMenu() {

  const menu = getSheet(SHEETS.MENU);

  const lastRow = menu.getLastRow();

  if (lastRow < 2) return [];

  const data = menu
      .getRange(2, 1, lastRow - 1, 9)
      .getValues();

  return data

      .filter(row => row[8] === "Yes")

      .map(row => ({

          itemId: row[0],

          itemName: row[1],

          category: row[2],

          cost: Number(row[3]),

          price: Number(row[4]),

          prepTime: row[7]

      }));

}







/**
 * ==========================================
 * Load Complete POS Cache
 * ==========================================
 */
function loadPOSCache() {

    return {

        menu: loadMenu(),

        settings: loadAppSettings(),

        customers: getAllCustomers()

    };

}



