/**
 * ==========================================================
 * Storage Service
 * ==========================================================
 *
 * Responsible for:
 *  - PDF Storage
 *  - Backup Storage
 *  - Report Storage
 *  - Customer Export
 *  - Purchase Export
 *
 * ==========================================================
 */


/**
 * ==========================================================
 * Save PDF
 * ==========================================================
 *
 * @param {String} fileName
 * @param {String} base64Data
 * @returns {Object}
 *
 * ==========================================================
 */
function savePdf(fileName, base64Data) {

  try {

    Logger.log("STEP 1");

    base64Data = base64Data.replace(
      /^data:application\/pdf;base64,/,
      ""
    );

    Logger.log("STEP 2");

    const bytes = Utilities.base64Decode(base64Data);

    Logger.log("STEP 3");

    const blob = Utilities.newBlob(
      bytes,
      MimeType.PDF,
      fileName
    );

    Logger.log("STEP 4");

    const folder = getInvoiceFolder();

    Logger.log("STEP 5");

    const file = folder.createFile(blob);

    Logger.log("STEP 6");

    return {

      success: true,

      fileUrl: file.getUrl()

    };

  }

  catch(err){

    Logger.log(err);

    return {

      success: false,

      message: err.toString()

    };

  }

}


/**
 * ==========================================
 * Save Invoice PDF
 * ==========================================
 */
function saveInvoicePdf(fileName, base64Pdf) {

  return savePdf(fileName, base64Pdf);

}