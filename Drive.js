/**
 * ==========================================
 * Google Drive Configuration
 * ==========================================
 */

const INVOICE_FOLDER_ID =
"1t5GjyLWAS8wfc03KVc5HJ81AfE2fHJ3K";


/**
 * ==========================================
 * Get Invoice Folder
 * ==========================================
 */
function getInvoiceFolder() {

  return DriveApp.getFolderById(INVOICE_FOLDER_ID);

}