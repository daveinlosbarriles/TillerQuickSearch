// Adds Tiller Tools menu to Google Sheets (both tools in this project)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tiller Tools")
    .addItem("Amazon Orders Import", "importAmazonCSV_LocalUpload")
    .addItem("Quick Search", "openQuickSearchSidebar")
    .addToUi();
}