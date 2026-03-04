// Adds Tiller Tools menu to Google Sheets (Quick Search sidebar)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tiller Tools")
    .addItem("Quick Search", "openQuickSearchSidebar")
    .addToUi();
}