/**
 * ==========================================
 * SMOKY SWAAD ERP
 * Web Application Entry Point
 * ==========================================
 */

function doGet() {

  return HtmlService
      .createTemplateFromFile("Index")
      .evaluate()
      .setTitle("Smoky Swaad")
      .setXFrameOptionsMode(
          HtmlService.XFrameOptionsMode.ALLOWALL
      );

}

/**
 * Include HTML files
 */
function include(filename){

  return HtmlService
      .createHtmlOutputFromFile(filename)
      .getContent();

}


/**
 * ==========================================
 * Load Screen
 * ==========================================
 */
function loadScreen(screen) {

  return HtmlService
    .createHtmlOutputFromFile(screen)
    .getContent();

}

