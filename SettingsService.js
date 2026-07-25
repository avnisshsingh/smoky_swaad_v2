/**
 * Load Application Settings
 */
function loadAppSettings() {

  const sheet = getSheet(SHEETS.SETTINGS);

  return {

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

}