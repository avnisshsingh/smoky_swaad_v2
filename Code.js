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
      .addMetaTag(
          "viewport",
          "width=device-width, initial-scale=1, viewport-fit=cover"
      )
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
  // Changing this to createTemplateFromFile tells Apps Script to process the <?!= ?> tags
  return HtmlService.createTemplateFromFile(screen).evaluate().getContent();
}



function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

