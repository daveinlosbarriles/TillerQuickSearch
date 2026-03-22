// Amazon Orders Import: CSV upload and append to Tiller Transactions sheet.
// Independent of Code.gs; menu in Code.gs calls importAmazonCSV_LocalUpload() defined here.
// Runtime config is read from the "AMZ Import" sheet; defaults below are used only when creating that sheet.

const AMZ_IMPORT_SHEET_NAME = "AMZ Import";

/** Header that identifies a Digital Content Orders CSV (takes precedence if both markers exist). */
const DIGITAL_MARKER_HEADER = "Digital Order Item ID";
/** Header that identifies a standard Order History CSV. */
const STANDARD_MARKER_HEADER = "Carrier Name & Tracking Number";

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
  }
};

// Defaults for seeding the "AMZ Import" sheet when it does not exist.
const AMZ_IMPORT_DEFAULTS = {
  INTRO_ROW: "The below settings are for managing Amazon Order CSV Import.",
  TABLE1_INTRO: "Add one row for each credit card you use with Amazon, and the appropriate values from your Tiller Accounts tab.",
  TABLE1_HEADERS: ["Payment Type", "Account", "Account #", "Institution", "Account ID", "User for Digital orders?"],
  TABLE1_ROWS: [
    ["Visa - 8534", "Chase Amazon Visa", "xxxx8534", "Chase", "636838acde7b2a0033ff46d5", "No"]
  ],
  TABLE2_INTRO: "Edit the left column only if Amazon changes the column names again in the CSV.",
  TABLE2_HEADERS: ["Amazon CSV column name", "Digital Orders CSV column name", "Name in Code"],
  TABLE2_ROWS: [
    ["Order Date", "Order Date", "Order Date"],
    ["Order ID", "Order ID", "Order ID"],
    ["Product Name", "Product Name", "Product Name"],
    ["Total Amount", "Transaction Amount", "Total Amount"],
    ["ASIN", "ASIN", "ASIN"],
    ["Payment Method Type", "Payment Method", "Payment Method Type"],
    ["Carrier Name & Tracking Number", "", "Carrier Name & Tracking Number"],
    ["Original Quantity", "Original Quantity", "Original Quantity"],
    ["Purchase Order Number", "", "Purchase Order Number"],
    ["Ship Date", "", "Ship Date"],
    ["Shipping Charge", "", "Shipping Charge"],
    ["Total Discounts", "", "Total Discounts"],
    ["Unit Price", "Price", "Unit Price"],
    ["Unit Price Tax", "Price Tax", "Unit Price Tax"],
    ["Website", "", "Website"]
  ],
  TABLE3_INTRO: "Only edit left column if Amazon changes column names again in the CSV. The fields below map amazon data to the metadata field.",
  TABLE3_HEADERS: ["Amazon CSV column name", "Digital Orders CSV column name", "Metadata field name"],
  TABLE3_ROWS: [
    ["Order ID", "Order ID", "id"],
    ["Original Quantity", "Original Quantity", "quantity"],
    ["Unit Price", "Price", "item-price"],
    ["Unit Price Tax", "Price Tax", "unit-price-tax"],
    ["Shipping Charge", "", "shipping-charge"],
    ["Total Discounts", "", "total-discounts"],
    ["Total Amount", "", "total"],
    ["Ship Date", "", "ship-date"],
    ["Carrier Name & Tracking Number", "", "tracking"],
    ["Payment Method Type", "Payment Information", "payment-type"],
    ["Website", "", "site"],
    ["Purchase Order Number", "", "purchase-order"],
    ["purchase", "", "type"]
  ],
  TABLE4_TITLE: "Sheet and Column labels used from Tiller",
  TABLE4_HEADERS: ["Name in Code", "Tiller label"],
  TABLE4_ROWS: [
    ["SHEET_NAME", "Transactions"],
    ["DATE", "Date"],
    ["DESCRIPTION", "Description"],
    ["AMOUNT", "Amount"],
    ["TRANSACTION_ID", "Transaction ID"],
    ["FULL_DESCRIPTION", "Full Description"],
    ["DATE_ADDED", "Date Added"],
    ["MONTH", "Month"],
    ["WEEK", "Week"],
    ["ACCOUNT", "Account"],
    ["ACCOUNT_NUMBER", "Account #"],
    ["INSTITUTION", "Institution"],
    ["ACCOUNT_ID", "Account ID"],
    ["METADATA", "Metadata"]
  ]
};

// Logical field names required in core mapping (Table 2) for standard Order History imports (legacy 2-column sheet).
const REQUIRED_CORE_FIELDS = [
  "Order Date",
  "Order ID",
  "Product Name",
  "Total Amount",
  "ASIN",
  "Payment Method Type"
];

// Required keys in Tiller labels (Table 4).
const REQUIRED_TILLER_LABEL_KEYS = [
  "SHEET_NAME", "DATE", "DESCRIPTION", "AMOUNT", "TRANSACTION_ID",
  "FULL_DESCRIPTION", "DATE_ADDED", "MONTH", "WEEK", "ACCOUNT",
  "ACCOUNT_NUMBER", "INSTITUTION", "ACCOUNT_ID", "METADATA"
];

const AMZ_IMPORT_INVALID_MSG = "AMZ Import configuration settings are missing or invalid. Suggest deleting that tab to load default values.";

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
 * Gets or creates the "AMZ Import" sheet. If it does not exist, creates it and fills with defaults.
 * @returns {{ sheet: GoogleAppsScript.Spreadsheet.Sheet, wasCreated: boolean }}
 */
function getOrCreateAmzImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(AMZ_IMPORT_SHEET_NAME);
  if (existing) return { sheet: existing, wasCreated: false };

  const sheet = ss.insertSheet(AMZ_IMPORT_SHEET_NAME);
  let row = 1;
  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.INTRO_ROW);
  row += 2;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE1_INTRO);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length).setFontWeight("bold");
  row += 1;
  if (AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length) {
    sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length)
      .setValues(AMZ_IMPORT_DEFAULTS.TABLE1_ROWS);
    row += AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length;
  }
  row += 1;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE2_INTRO);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE2_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE2_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE2_HEADERS.length).setFontWeight("bold");
  row += 1;
  sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE2_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE2_HEADERS.length)
    .setValues(AMZ_IMPORT_DEFAULTS.TABLE2_ROWS);
  row += AMZ_IMPORT_DEFAULTS.TABLE2_ROWS.length;
  row += 1;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE3_INTRO);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE3_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE3_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE3_HEADERS.length).setFontWeight("bold");
  row += 1;
  sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE3_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE3_HEADERS.length)
    .setValues(AMZ_IMPORT_DEFAULTS.TABLE3_ROWS);
  row += AMZ_IMPORT_DEFAULTS.TABLE3_ROWS.length;
  row += 1;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE4_TITLE);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length).setFontWeight("bold");
  row += 1;
  sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE4_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length)
    .setValues(AMZ_IMPORT_DEFAULTS.TABLE4_ROWS);

  return { sheet: sheet, wasCreated: true };
}

/**
 * Reads config from the "AMZ Import" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} amzSheet
 * @returns {{
 *   paymentAccounts: Object,
 *   coreMappingStandard: Object,
 *   coreMappingDigital: Object,
 *   metadataMapping: Array<{ key: string, standardCol: string, digitalCol: string }>,
 *   digitalUserAccount: Object|null,
 *   digitalUserYesCount: number,
 *   legacyTwoColumnCore: boolean,
 *   tillerLabels: Object|null
 * }}
 */
function readAmzImportConfig(amzSheet) {
  const data = amzSheet.getDataRange().getValues();
  const paymentAccounts = {};
  const coreMappingStandard = {};
  const coreMappingDigital = {};
  const metadataMapping = [];
  let digitalUserAccount = null;
  let digitalUserYesCount = 0;
  let legacyTwoColumnCore = false;
  let tillerLabels = null;

  let i = 0;
  while (i < data.length) {
    const row = data[i];
    const first = row[0] ? String(row[0]).trim() : "";

    if (first === "Payment Type") {
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const paymentType = String(r[0]).trim();
        const userForDigital = r[5] != null && String(r[5]).trim().toLowerCase() === "yes";
        if (userForDigital) {
          digitalUserYesCount += 1;
          digitalUserAccount = {
            ACCOUNT: r[1] != null ? String(r[1]) : "",
            ACCOUNT_NUMBER: r[2] != null ? String(r[2]) : "",
            INSTITUTION: r[3] != null ? String(r[3]) : "",
            ACCOUNT_ID: r[4] != null ? String(r[4]) : ""
          };
        }
        paymentAccounts[paymentType] = {
          ACCOUNT: r[1] != null ? String(r[1]) : "",
          ACCOUNT_NUMBER: r[2] != null ? String(r[2]) : "",
          INSTITUTION: r[3] != null ? String(r[3]) : "",
          ACCOUNT_ID: r[4] != null ? String(r[4]) : ""
        };
        i += 1;
      }
      continue;
    }
    const second = row[1] ? String(row[1]).trim() : "";
    const third = row[2] != null ? String(row[2]).trim() : "";

    if (first === "Amazon CSV column name" && second === "Digital Orders CSV column name" && third === "Name in Code") {
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const amazonCol = r[0] != null ? String(r[0]).trim() : "";
        const digitalCol = r[1] != null ? String(r[1]).trim() : "";
        const fieldName = r[2] != null ? String(r[2]).trim() : "";
        if (fieldName) {
          if (amazonCol) coreMappingStandard[fieldName] = amazonCol;
          if (digitalCol) coreMappingDigital[fieldName] = digitalCol;
        }
        i += 1;
      }
      continue;
    }
    if (first === "Amazon CSV column name" && second === "Name in Code") {
      legacyTwoColumnCore = true;
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const csvCol = r[0] != null ? String(r[0]).trim() : "";
        const fieldName = r[1] != null ? String(r[1]).trim() : "";
        if (fieldName && csvCol) {
          coreMappingStandard[fieldName] = csvCol;
          coreMappingDigital[fieldName] = csvCol;
        }
        i += 1;
      }
      continue;
    }
    if (first === "Amazon CSV column name" && second === "Digital Orders CSV column name" && third === "Metadata field name") {
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const amazonCol = r[0] != null ? String(r[0]).trim() : "";
        const digitalCol = r[1] != null ? String(r[1]).trim() : "";
        const jsonKey = r[2] != null ? String(r[2]).trim() : "";
        if (jsonKey) {
          metadataMapping.push({ key: jsonKey, standardCol: amazonCol, digitalCol: digitalCol });
        }
        i += 1;
      }
      continue;
    }
    if (first === "Amazon CSV column name" && second === "Metadata field name") {
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const csvColOrLiteral = r[0] != null ? String(r[0]).trim() : "";
        const jsonKey = r[1] != null ? String(r[1]).trim() : "";
        if (jsonKey) {
          metadataMapping.push({ key: jsonKey, standardCol: csvColOrLiteral, digitalCol: csvColOrLiteral });
        }
        i += 1;
      }
      continue;
    }
    if (first === "Name in Code" && second === "Tiller label") {
      tillerLabels = {};
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const nameInCode = r[0] != null ? String(r[0]).trim() : "";
        const tillerLabel = r[1] != null ? String(r[1]).trim() : "";
        if (nameInCode) tillerLabels[nameInCode] = tillerLabel;
        i += 1;
      }
      continue;
    }
    i += 1;
  }

  return {
    paymentAccounts,
    coreMappingStandard,
    coreMappingDigital,
    metadataMapping,
    digitalUserAccount,
    digitalUserYesCount,
    legacyTwoColumnCore,
    tillerLabels
  };
}

/**
 * @param {Object} col - CSV header name -> column index
 * @returns {"digital"|"standard"|null}
 */
function detectAmazonCsvFileType(col) {
  if (col[DIGITAL_MARKER_HEADER] !== undefined && col[DIGITAL_MARKER_HEADER] !== null) {
    return "digital";
  }
  if (col[STANDARD_MARKER_HEADER] !== undefined && col[STANDARD_MARKER_HEADER] !== null) {
    return "standard";
  }
  return null;
}

/**
 * @param {Object} col
 * @param {{ coreMappingStandard: Object, coreMappingDigital: Object, metadataMapping: Array }} config
 * @param {boolean} isDigital
 * @returns {string|null} Error message or null if OK.
 */
function validateMappedCsvHeadersPresent(col, config, isDigital) {
  const map = isDigital ? config.coreMappingDigital : config.coreMappingStandard;
  const other = isDigital ? config.coreMappingStandard : config.coreMappingDigital;
  const fieldNames = {};
  Object.keys(map).forEach(function (k) { fieldNames[k] = true; });
  Object.keys(other).forEach(function (k) { fieldNames[k] = true; });

  for (const fieldName in fieldNames) {
    const resolved = isDigital
      ? (config.coreMappingDigital[fieldName] || "")
      : (config.coreMappingStandard[fieldName] || "");
    if (!resolved || String(resolved).trim() === "") continue;
    if (col[resolved] === undefined || col[resolved] === null) {
      return "Missing required CSV column for this file type: \"" + resolved + "\" (logical field: " + fieldName + "). Check AMZ Import Table 2.";
    }
  }

  for (let m = 0; m < config.metadataMapping.length; m++) {
    const row = config.metadataMapping[m];
    const resolved = isDigital ? row.digitalCol : row.standardCol;
    if (!resolved || String(resolved).trim() === "") continue;
    if (col[resolved] === undefined || col[resolved] === null) {
      return "Missing required CSV column for this file type: \"" + resolved + "\" (metadata key: " + row.key + "). Check AMZ Import Table 3.";
    }
  }
  return null;
}

/**
 * Resolves core mapping for one logical field for the current file type.
 */
function getCoreCsvColumn(config, fieldName, isDigital) {
  if (isDigital) {
    const d = config.coreMappingDigital[fieldName];
    if (d != null && String(d).trim() !== "") return String(d).trim();
    return config.coreMappingStandard[fieldName] || "";
  }
  const s = config.coreMappingStandard[fieldName];
  if (s != null && String(s).trim() !== "") return String(s).trim();
  return config.coreMappingDigital[fieldName] || "";
}

/**
 * Validates that all required AMZ Import config is present. No fallbacks.
 * @param {*} config
 * @returns {string|null} Null if valid; otherwise the error message to show.
 */
function validateAmzImportConfig(config) {
  if (!config) return AMZ_IMPORT_INVALID_MSG;
  if (!config.paymentAccounts || Object.keys(config.paymentAccounts).length === 0) {
    return AMZ_IMPORT_INVALID_MSG;
  }
  if (!config.tillerLabels || typeof config.tillerLabels !== "object") {
    return AMZ_IMPORT_INVALID_MSG;
  }
  for (let k = 0; k < REQUIRED_TILLER_LABEL_KEYS.length; k++) {
    const key = REQUIRED_TILLER_LABEL_KEYS[k];
    const val = config.tillerLabels[key];
    if (val === undefined || val === null || String(val).trim() === "") {
      return AMZ_IMPORT_INVALID_MSG;
    }
  }
  if (config.legacyTwoColumnCore) {
    for (let f = 0; f < REQUIRED_CORE_FIELDS.length; f++) {
      const fieldName = REQUIRED_CORE_FIELDS[f];
      const csvCol = config.coreMappingStandard && config.coreMappingStandard[fieldName];
      if (!csvCol || String(csvCol).trim() === "") return AMZ_IMPORT_INVALID_MSG;
    }
  } else {
    const need = ["Order Date", "Order ID", "Product Name", "Total Amount", "ASIN"];
    for (let n = 0; n < need.length; n++) {
      const fieldName = need[n];
      const hasStd = config.coreMappingStandard[fieldName] && String(config.coreMappingStandard[fieldName]).trim() !== "";
      const hasDig = config.coreMappingDigital[fieldName] && String(config.coreMappingDigital[fieldName]).trim() !== "";
      if (!hasStd && !hasDig) return AMZ_IMPORT_INVALID_MSG;
    }
  }
  return null;
}

/**
 * Builds the metadata amazon object for one CSV row using the metadata mapping.
 * @param {Array} csvRow - CSV row (array of values)
 * @param {Object} col - Map of CSV column name -> index
 * @param {Array} metadataMapping - Array of { key, standardCol, digitalCol }
 * @param {boolean} isDigital
 * @returns {Object}
 */
function buildAmazonMetadataObject(csvRow, col, metadataMapping, isDigital) {
  const numericKeys = ["quantity", "item-price", "unit-price-tax", "shipping-charge", "total-discounts", "total"];
  const obj = {};
  metadataMapping.forEach(function (m) {
    const k = m.key;
    let src = isDigital ? m.digitalCol : m.standardCol;
    if (src == null || String(src).trim() === "") {
      src = isDigital ? m.standardCol : m.digitalCol;
    }
    src = src != null ? String(src).trim() : "";
    if (!src) return;

    const colIndex = col[src];
    let val;
    if (colIndex !== undefined && colIndex !== null) {
      const raw = csvRow[colIndex];
      if (raw === "" || raw === undefined || raw === null) {
        val = numericKeys.indexOf(k) >= 0 ? 0 : "";
      } else if (numericKeys.indexOf(k) >= 0) {
        val = parseFloat(raw);
        if (isNaN(val)) val = 0;
      } else {
        val = String(raw).trim();
      }
    } else {
      if (numericKeys.indexOf(k) >= 0) val = 0;
      else val = src;
    }
    obj[k] = val;
  });
  return obj;
}

/**
 * Opens the Amazon Orders import dialog (HTML from AmazonOrdersDialog.html).
 * Ensures the AMZ Import config sheet exists and is valid; shows message if not.
 */
function importAmazonCSV_LocalUpload() {
  const result = getOrCreateAmzImportSheet();
  if (result.wasCreated) {
    SpreadsheetApp.getUi().alert("Please update your payment types on the AMZ Import tab before proceeding with your import.");
  }
  const config = readAmzImportConfig(result.sheet);
  const err = validateAmzImportConfig(config);
  if (err) {
    SpreadsheetApp.getUi().alert(err);
  } else if (config.tillerLabels && config.tillerLabels.SHEET_NAME) {
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.tillerLabels.SHEET_NAME);
    if (targetSheet) targetSheet.activate();
  }
  const html = HtmlService.createHtmlOutputFromFile("AmazonOrdersDialog")
    .setWidth(600)
    .setHeight(520);
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const paymentAccounts = config.paymentAccounts;
  const tillerLabels = config.tillerLabels;

  const configErr = validateAmzImportConfig(config);
  if (configErr) return configErr;

  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet '" + sheetName + "' not found.";

  const tillerCols = getTillerColumnMap(sheet);

  const tParseStart = Date.now();
  const csv = Utilities.parseCsv(csvText);
  const tParseEnd = Date.now();
  timing.push("Server: parse CSV: " + ((tParseEnd - tParseStart) / 1000).toFixed(2) + " s");

  if (!csv.length || !csv[0].length) {
    return "Error: CSV is empty or has no header row.";
  }

  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });

  const fileKind = detectAmazonCsvFileType(col);
  if (!fileKind) {
    return "Could not detect file type. The CSV must include column \"" + DIGITAL_MARKER_HEADER + "\" (Digital Content Orders) or \"" + STANDARD_MARKER_HEADER + "\" (Order History).";
  }
  const isDigital = fileKind === "digital";
  const detectedLabel = isDigital
    ? "Detected file type: Digital Content Orders"
    : "Detected file type: Order History (standard)";

  if (isDigital) {
    if (config.digitalUserYesCount === 0) {
      return detectedLabel + "\n" + "Digital Content Orders import requires exactly one row on AMZ Import Table 1 with \"User for Digital orders?\" set to Yes.";
    }
    if (config.digitalUserYesCount > 1) {
      return detectedLabel + "\n" + "Digital Content Orders import: multiple rows have \"User for Digital orders?\" set to Yes. Only one row should be Yes.";
    }
    if (!config.digitalUserAccount) {
      return detectedLabel + "\n" + "Digital Content Orders import: could not read account fields from the row with User for Digital orders? = Yes.";
    }
  }

  const headerErr = validateMappedCsvHeadersPresent(col, config, isDigital);
  if (headerErr) return detectedLabel + "\n" + headerErr;

  let cutoff = null;
  if (months) {
    cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
  }

  const lastRow = sheet.getLastRow();
  const existingFullDescSet = new Set();
  const tDupStart = Date.now();
  if (lastRow > 1) {
    const fullDescCol = tillerCols[tillerLabels.FULL_DESCRIPTION];
    if (fullDescCol) {
      const fullDescs = sheet.getRange(2, fullDescCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < fullDescs.length; i++) {
        const val = fullDescs[i][0];
        if (val) existingFullDescSet.add(String(val));
      }
    }
  }
  const tDupEnd = Date.now();
  timing.push(
    "Server: read Full Description column + build duplicate set (" +
    existingFullDescSet.size + " entries): " +
    ((tDupEnd - tDupStart) / 1000).toFixed(2) + " s"
  );

  const orderDateCol = getCoreCsvColumn(config, "Order Date", isDigital);
  const orderIdCol = getCoreCsvColumn(config, "Order ID", isDigital);
  const productNameCol = getCoreCsvColumn(config, "Product Name", isDigital);
  const asinCol = getCoreCsvColumn(config, "ASIN", isDigital);
  const totalAmountCol = getCoreCsvColumn(config, "Total Amount", isDigital);
  const paymentMethodColName = getCoreCsvColumn(config, "Payment Method Type", isDigital);

  if (!orderDateCol || !orderIdCol || !productNameCol || !asinCol || !totalAmountCol) {
    return "AMZ Import Table 2 is missing required core column mappings for this file type.";
  }
  if (!isDigital && (!paymentMethodColName || col[paymentMethodColName] === undefined)) {
    return "Your CSV must include the payment column mapped for Payment Method Type. Please request a new Order History from Amazon.";
  }

  const numCols = sheet.getLastColumn();
  const output = [];
  const totalByPaymentMethod = {};
  let duplicateCount = 0;
  const runTimestamp = new Date();
  const importTimestampStr = Utilities.formatDate(runTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  const tLoopStart = Date.now();
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    let orderDate = new Date(r[col[orderDateCol]]);
    orderDate.setHours(0, 0, 0, 0);
    if (cutoff && orderDate < cutoff) continue;

    const orderID = r[col[orderIdCol]];
    const productName = r[col[productNameCol]];
    const asin = r[col[asinCol]];

    const expectedFullDesc =
      "Amazon Order ID " + orderID + ": " +
      productName + " (" + asin + ")";

    if (existingFullDescSet.has(expectedFullDesc)) {
      duplicateCount += 1;
      continue;
    }

    let accountRow;
    if (isDigital) {
      accountRow = config.digitalUserAccount;
    } else {
      const paymentMethodType = String(r[col[paymentMethodColName]] || "").trim();
      accountRow = paymentAccounts[paymentMethodType];
      if (!accountRow) {
        return "Payment type \"" + paymentMethodType + "\" not found. Import was stopped. Add new payment type to AMZ Import tab.";
      }
    }

    const amount = parseFloat(r[col[totalAmountCol]]) * -1;
    const payKey = isDigital ? "Digital" : String(r[col[paymentMethodColName]] || "").trim();
    totalByPaymentMethod[payKey] = (totalByPaymentMethod[payKey] || 0) + amount;

    const month = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), "yyyy-MM");
    const week = getWeekStartDate(orderDate);

    const descriptionText = (isDigital ? "[AMZD] " : "[AMZ] ") + productName;
    const fullDesc = expectedFullDesc;

    const amazonMeta = buildAmazonMetadataObject(r, col, config.metadataMapping, isDigital);
    const metadataValue = "Imported by AmazonCSVImporter on " + importTimestampStr + " " + JSON.stringify({ amazon: amazonMeta });

    const rowOut = new Array(numCols).fill("");

    rowOut[tillerCols[tillerLabels.DATE] - 1] = orderDate;
    rowOut[tillerCols[tillerLabels.DESCRIPTION] - 1] = descriptionText;
    rowOut[tillerCols[tillerLabels.FULL_DESCRIPTION] - 1] = fullDesc;
    rowOut[tillerCols[tillerLabels.AMOUNT] - 1] = amount;
    rowOut[tillerCols[tillerLabels.TRANSACTION_ID] - 1] = generateGuid();
    rowOut[tillerCols[tillerLabels.DATE_ADDED] - 1] = runTimestamp;
    rowOut[tillerCols[tillerLabels.MONTH] - 1] = month;
    rowOut[tillerCols[tillerLabels.WEEK] - 1] = week;
    rowOut[tillerCols[tillerLabels.ACCOUNT] - 1] = accountRow.ACCOUNT;
    rowOut[tillerCols[tillerLabels.ACCOUNT_NUMBER] - 1] = accountRow.ACCOUNT_NUMBER;
    rowOut[tillerCols[tillerLabels.INSTITUTION] - 1] = accountRow.INSTITUTION;
    rowOut[tillerCols[tillerLabels.ACCOUNT_ID] - 1] = accountRow.ACCOUNT_ID;
    rowOut[tillerCols[tillerLabels.METADATA] - 1] = metadataValue;

    output.push(rowOut);
  }
  const tLoopEnd = Date.now();
  timing.push(
    "Server: main loop over CSV (" + (csv.length - 1) + " data rows, " +
    output.length + " new rows): " +
    ((tLoopEnd - tLoopStart) / 1000).toFixed(2) + " s"
  );

  if (!output.length) {
    let msg = "No new transactions found";
    if (duplicateCount > 0) {
      msg += "\n" + duplicateCount + " duplicate transactions were found and were not imported.";
    }
    return detectedLabel + "\n" + msg;
  }

  const tWriteNewStart = Date.now();
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, output.length, numCols).setValues(output);
  const tWriteNewEnd = Date.now();
  timing.push(
    "Server: write new rows to sheet (" + output.length + " rows): " +
    ((tWriteNewEnd - tWriteNewStart) / 1000).toFixed(2) + " s"
  );

  const tWriteOffsetStart = Date.now();
  const offsetRows = [];
  const offsetNow = new Date();
  const offsetMonth = Utilities.formatDate(offsetNow, Session.getScriptTimeZone(), "yyyy-MM");
  const offsetWeek = getWeekStartDate(offsetNow);
  const offsetDate = new Date(offsetNow);
  offsetDate.setHours(0, 0, 0, 0);
  const offsetTimestampStr = Utilities.formatDate(offsetNow, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for (const paymentMethodType in totalByPaymentMethod) {
    const total = totalByPaymentMethod[paymentMethodType];
    if (total === 0) continue;
    const accountRow = isDigital
      ? config.digitalUserAccount
      : paymentAccounts[paymentMethodType];
    if (!accountRow) continue;

    const desc = isDigital
      ? ("Amazon purchase offset (Digital) for " + offsetTimestampStr)
      : ("Amazon purchase offset (" + paymentMethodType + ") for " + offsetTimestampStr);
    const offset = new Array(numCols).fill("");
    offset[tillerCols[tillerLabels.DATE] - 1] = offsetDate;
    offset[tillerCols[tillerLabels.DESCRIPTION] - 1] = desc;
    offset[tillerCols[tillerLabels.FULL_DESCRIPTION] - 1] = desc;
    offset[tillerCols[tillerLabels.AMOUNT] - 1] = Math.abs(total);
    offset[tillerCols[tillerLabels.TRANSACTION_ID] - 1] = generateGuid();
    offset[tillerCols[tillerLabels.DATE_ADDED] - 1] = offsetNow;
    offset[tillerCols[tillerLabels.MONTH] - 1] = offsetMonth;
    offset[tillerCols[tillerLabels.WEEK] - 1] = offsetWeek;
    offset[tillerCols[tillerLabels.ACCOUNT] - 1] = accountRow.ACCOUNT;
    offset[tillerCols[tillerLabels.ACCOUNT_NUMBER] - 1] = accountRow.ACCOUNT_NUMBER;
    offset[tillerCols[tillerLabels.INSTITUTION] - 1] = accountRow.INSTITUTION;
    offset[tillerCols[tillerLabels.ACCOUNT_ID] - 1] = accountRow.ACCOUNT_ID;
    offset[tillerCols[tillerLabels.METADATA] - 1] = "Imported by AmazonCSVImporter on " + importTimestampStr;
    offsetRows.push(offset);
  }

  if (offsetRows.length > 0) {
    const offsetStartRow = sheet.getLastRow() + 1;
    sheet.getRange(offsetStartRow, 1, offsetRows.length, numCols).setValues(offsetRows);
  }
  const tWriteOffsetEnd = Date.now();
  timing.push(
    "Server: write offset row(s) to sheet (" + offsetRows.length + " row(s)): " +
    ((tWriteOffsetEnd - tWriteOffsetStart) / 1000).toFixed(2) + " s"
  );

  const tSortStart = Date.now();
  const sortLastRow = sheet.getLastRow();
  if (sortLastRow >= 2) {
    const sortNumRows = sortLastRow - 1;
    sheet.getRange(2, 1, sortNumRows, sheet.getLastColumn())
      .sort({ column: tillerCols[tillerLabels.DATE], ascending: false });
  }
  const tSortEnd = Date.now();
  timing.push(
    "Server: sort sheet by Date: " +
    ((tSortEnd - tSortStart) / 1000).toFixed(2) + " s"
  );

  const metaCol = tillerCols[tillerLabels.METADATA];
  if (metaCol) {
    try {
      let filter = sheet.getFilter();
      const lastRowNow = sheet.getLastRow();
      const numColsNow = sheet.getLastColumn();
      const dataRange = sheet.getRange(1, 1, lastRowNow, numColsNow);
      if (!filter) {
        filter = dataRange.createFilter();
      } else {
        const fr = filter.getRange();
        if (fr.getNumRows() !== lastRowNow || fr.getNumColumns() !== numColsNow) {
          filter.remove();
          filter = dataRange.createFilter();
        }
      }
      const criteria = SpreadsheetApp.newFilterCriteria()
        .whenTextContains(importTimestampStr)
        .build();
      filter.setColumnFilterCriteria(metaCol, criteria);
    } catch (e) {
      // If filter or Metadata column fails, don't fail the import
    }
  }

  const tEnd = Date.now();
  timing.push(
    "Server: TOTAL importAmazonRecent time: " +
    ((tEnd - t0) / 1000).toFixed(2) + " s"
  );

  let summary = output.length + " transactions imported";
  if (duplicateCount > 0) {
    summary += "\n" + duplicateCount + " duplicate transactions were found and were not imported.";
  }
  timing.unshift(summary);
  timing.unshift(detectedLabel);

  return timing.join("\n");
}
