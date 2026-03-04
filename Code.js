// Adds Tiller Tools menu to Google Sheets
// Allows launching the Amazon Orders import dialog and Quick Search sidebar
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tiller Tools")
    .addItem("Amazon Orders Import", "importAmazonCSV_LocalUpload")
    .addItem("Quick Search", "openQuickSearchSidebar")
    .addToUi();
}