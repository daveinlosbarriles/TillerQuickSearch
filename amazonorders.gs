// Amazon Orders Import: CSV upload and append to Tiller Transactions sheet.
// Independent of Code.gs; menu in Code.gs calls importAmazonCSV_LocalUpload() defined here.

const AMAZON_CONFIG = {
  COLUMNS: {
    ORDER_DATE: "Order Date",
    ORDER_ID: "Order ID",
    PRODUCT_NAME: "Product Name",
    TOTAL_AMOUNT: "Total Amount",
    ASIN: "ASIN"
  }
};

const TILLER_CONFIG = {
  SHEET_NAME: "Transactions",
  COLUMNS: {
    DATE: "Date",
    DESCRIPTION: "Description",
    AMOUNT: "Amount",
    TRANSACTION_ID: "Transaction ID",
    FULL_DESCRIPTION: "Full Description",
    DATE_ADDED: "Date Added",
    MONTH: "Month",
    WEEK: "Week",
    ACCOUNT: "Account",
    ACCOUNT_NUMBER: "Account #",
    INSTITUTION: "Institution",
    ACCOUNT_ID: "Account ID",
    METADATA: "Metadata"
  },
  STATIC_VALUES: {
    ACCOUNT: "Chase Amazon Visa",
    ACCOUNT_NUMBER: "xxxx8534",
    INSTITUTION: "Chase",
    ACCOUNT_ID: "636838acde7b2a0033ff46d5"
  }
};

function generateGuid() {
  return Utilities.getUuid();
}

function getWeekStartDate(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getTillerColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[h.trim()] = i + 1;
  });
  return map;
}

/**
 * Opens the Amazon Orders import dialog (HTML from AmazonOrdersDialog.html).
 */
function importAmazonCSV_LocalUpload() {
  const html = HtmlService.createHtmlOutputFromFile("AmazonOrdersDialog")
    .setWidth(600)
    .setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, "Import Amazon Orders");
}

/**
 * Imports Amazon CSV data into the Tiller Transactions sheet.
 * Called from AmazonOrdersDialog.html via google.script.run.importAmazonRecent(text, months).
 * @param {string} csvText - Raw CSV file content
 * @param {number|null} months - Optional months lookback; null = all rows
 * @returns {string} Newline-separated summary and timing lines for the dialog log
 */
function importAmazonRecent(csvText, months) {
  const t0 = Date.now();
  const timing = [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(TILLER_CONFIG.SHEET_NAME);

  const tillerCols = getTillerColumnMap(sheet);

  const tParseStart = Date.now();
  const csv = Utilities.parseCsv(csvText);
  const tParseEnd = Date.now();
  timing.push("Server: parse CSV: " + ((tParseEnd - tParseStart) / 1000).toFixed(2) + " s");

  const headers = csv[0];
  const col = {};
  headers.forEach((h, i) => col[h.trim()] = i);

  let cutoff = null;
  if (months) {
    cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
  }

  const lastRow = sheet.getLastRow();

  const existingFullDescSet = new Set();
  const tDupStart = Date.now();
  if (lastRow > 1) {
    const fullDescs = sheet.getRange(
      2,
      tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION],
      lastRow - 1,
      1
    ).getValues();

    for (let i = 0; i < fullDescs.length; i++) {
      const val = fullDescs[i][0];
      if (val) {
        existingFullDescSet.add(String(val));
      }
    }
  }
  const tDupEnd = Date.now();
  timing.push(
    "Server: read Full Description column + build duplicate set (" +
    existingFullDescSet.size + " entries): " +
    ((tDupEnd - tDupStart) / 1000).toFixed(2) + " s"
  );

  const numCols = sheet.getLastColumn();
  const output = [];
  let totalImported = 0;

  const tLoopStart = Date.now();
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    let orderDate = new Date(r[col[AMAZON_CONFIG.COLUMNS.ORDER_DATE]]);
    orderDate.setHours(0, 0, 0, 0);
    if (cutoff && orderDate < cutoff) continue;

    const orderID = r[col[AMAZON_CONFIG.COLUMNS.ORDER_ID]];
    const productName = r[col[AMAZON_CONFIG.COLUMNS.PRODUCT_NAME]];
    const asin = r[col[AMAZON_CONFIG.COLUMNS.ASIN]];

    const expectedFullDesc =
      "Amazon Order ID " + orderID + ": " +
      productName + " (" + asin + ")";

    if (existingFullDescSet.has(expectedFullDesc)) continue;

    const amount = parseFloat(r[col[AMAZON_CONFIG.COLUMNS.TOTAL_AMOUNT]]) * -1;
    totalImported += amount;

    const now = new Date();
    const month = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), "yyyy-MM");
    const week = getWeekStartDate(orderDate);

    const descriptionText = "[AMZ] " + productName;
    const fullDesc = expectedFullDesc;

    const row = new Array(numCols).fill("");

    row[tillerCols[TILLER_CONFIG.COLUMNS.DATE] - 1] = orderDate;
    row[tillerCols[TILLER_CONFIG.COLUMNS.DESCRIPTION] - 1] = descriptionText;
    row[tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION] - 1] = fullDesc;
    row[tillerCols[TILLER_CONFIG.COLUMNS.AMOUNT] - 1] = amount;
    row[tillerCols[TILLER_CONFIG.COLUMNS.TRANSACTION_ID] - 1] = generateGuid();
    row[tillerCols[TILLER_CONFIG.COLUMNS.DATE_ADDED] - 1] = now;
    row[tillerCols[TILLER_CONFIG.COLUMNS.MONTH] - 1] = month;
    row[tillerCols[TILLER_CONFIG.COLUMNS.WEEK] - 1] = week;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_NUMBER] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_NUMBER;
    row[tillerCols[TILLER_CONFIG.COLUMNS.INSTITUTION] - 1] = TILLER_CONFIG.STATIC_VALUES.INSTITUTION;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_ID] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_ID;
    row[tillerCols[TILLER_CONFIG.COLUMNS.METADATA] - 1] =
      "Imported by AmazonCSVImporter on " + now;

    output.push(row);
  }
  const tLoopEnd = Date.now();
  timing.push(
    "Server: main loop over CSV (" + (csv.length - 1) + " data rows, " +
    output.length + " new rows): " +
    ((tLoopEnd - tLoopStart) / 1000).toFixed(2) + " s"
  );

  if (!output.length) return "No new transactions found";

  const tWriteNewStart = Date.now();
  sheet.getRange(sheet.getLastRow() + 1, 1, output.length, numCols).setValues(output);
  const tWriteNewEnd = Date.now();
  timing.push(
    "Server: write new rows to sheet (" + output.length + " rows): " +
    ((tWriteNewEnd - tWriteNewStart) / 1000).toFixed(2) + " s"
  );

  if (totalImported !== 0) {
    const now = new Date();
    const month = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
    const week = getWeekStartDate(now);

    const offset = new Array(numCols).fill("");

    const desc = "Amazon purchase offset for " +
      Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    const offsetDate = new Date(now);
    offsetDate.setHours(0, 0, 0, 0);

    offset[tillerCols[TILLER_CONFIG.COLUMNS.DATE] - 1] = offsetDate;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.DESCRIPTION] - 1] = desc;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION] - 1] = desc;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.AMOUNT] - 1] = Math.abs(totalImported);
    offset[tillerCols[TILLER_CONFIG.COLUMNS.TRANSACTION_ID] - 1] = generateGuid();
    offset[tillerCols[TILLER_CONFIG.COLUMNS.DATE_ADDED] - 1] = now;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.MONTH] - 1] = month;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.WEEK] - 1] = week;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_NUMBER] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_NUMBER;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.INSTITUTION] - 1] = TILLER_CONFIG.STATIC_VALUES.INSTITUTION;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_ID] - 1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_ID;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.METADATA] - 1] =
      "Imported by AmazonCSVImporter on " + now;

    const tWriteOffsetStart = Date.now();
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, numCols).setValues([offset]);
    const tWriteOffsetEnd = Date.now();
    timing.push(
      "Server: write offset row to sheet: " +
      ((tWriteOffsetEnd - tWriteOffsetStart) / 1000).toFixed(2) + " s"
    );
  }

  const tSortStart = Date.now();
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .sort({ column: tillerCols[TILLER_CONFIG.COLUMNS.DATE], ascending: false });
  const tSortEnd = Date.now();
  timing.push(
    "Server: sort sheet by Date: " +
    ((tSortEnd - tSortStart) / 1000).toFixed(2) + " s"
  );

  const tEnd = Date.now();
  timing.push(
    "Server: TOTAL importAmazonRecent time: " +
    ((tEnd - t0) / 1000).toFixed(2) + " s"
  );

  const summary = output.length + " transactions imported";
  timing.unshift(summary);

  return timing.join("\n");
}
