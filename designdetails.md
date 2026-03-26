# Tiller Amazon Import — design details

This document describes how the **ZIP sidebar** (`AmazonOrdersSidebar.html`) and server (`amazonorders.gs`) work together: payloads, config on **AMZ Import**, deduplication, offsets, payments, metadata, sheet post-processing, and error handling.

---

## 1. Loading and parsing one file at a time (payload size)

**Goal:** Keep each `google.script.run` RPC under Apps Script limits by not sending the full ZIP as one giant argument.

**Client (`AmazonOrdersSidebar.html`):**

- After ZIP extract, full CSV texts live in `state.files`.
- Import builds an ordered **step list** (`orderHistory`, `digitalOrders`, `digitalReturns`, `refundDetails`) from checkboxes.
- **`pump()`** sends **one** `importAmazonBundleChunk(JSON.stringify(payload))` per step. Each payload includes only the CSV string(s) needed for that step (e.g. `refundDetails` also sends trimmed `orderHistoryCsv` for payment join).
- A final **`step: "finalize"`** call runs **sort + Metadata filter** once after all chunks (`importAmazonBundleChunk` in `amazonorders.gs`, `finalize` branch).

**Server (`amazonorders.gs`):**

- `importAmazonBundleChunk` parses JSON, assigns a stable **`bundleImportTimestampIso`** for the run, and dispatches to `importAmazonRecent`, `importDigitalReturnsCsv`, or `importRefundDetailsCsv` with **deferred** sheet post-process so sort/filter happens in `finalize`.

**Legacy path:** `importAmazonBundle` can run all sections in one call; the sidebar uses chunks.

---

## 2. Trimming CSV by date before upload (payload size)

**Where:** Client only — `buildTransmitFiles` in `AmazonOrdersSidebar.html` uses **Papa Parse** to drop data rows **before** a minimum calendar date while keeping the header row.

**Anchor date:** `computeTransmitAnchorDate` — max of “latest” relevant dates across selected pipelines (OH/DO `Order Date`, refunds `Refund Date`, digital returns `Return Date`), or today — so trim windows are tied to the ZIP, not an arbitrary clock.

**Order History / Digital Orders:**

- **With cutoff:** Minimum transmit date = start of cutoff day, optionally minus **`AMZ_OH_DO_ANCHOR_TRIM_DAYS` (45)** when **refunds** or **digital returns** are selected — so **older OH/DO rows** can be included **only for server-side payment lookup** without increasing refund/return CSV size the same way.
- **Without cutoff:** Minimum = anchor − 45 days (still caps huge history files).

**Refund Details / Digital Returns (with cutoff):** Trimmed on **`Refund Date`** / **`Return Date`** to cutoff start (client uses those header names for trim; server can use renamed columns from **AMZ Import** for parsing).

**Server:** Still applies the same **cutoff** when inserting rows (`amzCutoffStartOfDay_`); trimming is an extra optimization for **RPC payload** only.

**Edge case:** Rows with **missing/unparseable** dates in the trimmed column are **kept** in `trimCsvRowsByMinDate` so they aren’t silently dropped before the server can log or handle them.

---

## 3. Constants from **AMZ Import** (sheet-driven config)

Configuration is read by `readAmzImportConfig` and validated by `validateAmzImportConfig`. Hard-coded **defaults** exist only to **seed** a new tab (`AMZ_IMPORT_DEFAULTS`, `getOrCreateAmzImportSheet`).

| Area | Storage on sheet | Purpose | Pipelines / consumers |
|------|------------------|---------|------------------------|
| **Payment → account** | Table: Payment Type, Account, …, Use for Digital | Map Amazon payment **string** to Tiller account fields; exactly one **Yes** for digital | Orders (standard), offsets, `analyzePaymentMethodsForOrderHistory`; digital user row for digital **orders** |
| **CSV column map** | `Source file` \| `Header` \| `Name in code` \| `Metadata field name` | Maps Amazon file/column → logical field + metadata JSON keys; `_file_detection` rows set **standard vs digital** marker headers | All imports; `amzDetectAmazonCsvFileType`; `amzGetCoreCsvColumn`; `amzGetSourceMapHeader` for refunds/returns |
| **Tiller column labels** | Name in Code → Tiller label | Resolves **Transactions** column headers (sheet name, Date, Metadata, …) | All writers + dedup scan + sort/filter |
| **In-code only** | — | e.g. `AMZ_WHOLE_FOODS_WEBSITE` (`panda01`), description prefixes | Website filter on Order History; labels |

**Note:** The unified CSV map is required (`csvMapPresent`); legacy Table 2/3 blocks are not read.

---

## 4. Duplicate detection (existing sheet + new import)

**Principle:** Dedup keys are derived from **Metadata** JSON’s `amazon` object, not from **Full Description** (users may edit descriptions).

**Before import:**

- `amzGetLastTransactionDataRow` — last row with a **Date** (not merely “last row”).
- `amzAppendDuplicateKeysFromTransactions_` — reads Metadata column down to that row; `amzAddDuplicateKeysFromImportMetadataCell_` parses JSON and calls `amzAddDedupKeysForAmazonMeta_`.

**Key shapes (by `amazon.type`):** e.g. `physical-purchase-line|orderId|ASIN`, `digital-purchase|orderId`, `refund-detail|orderId|amount`, `digital-return|orderId|asin`, plus `*-offset` types — see `amzAddDedupKeysForAmazonMeta_`.

**During import:**

- New rows: check `Set` before append; increment duplicate counter if key already present.
- New keys are added to the set as rows are queued so **within-file** duplicates are also caught.

---

## 5. Offset calculation (purchase / digital purchase)

**Where:** `importAmazonRecent` — **`perOrderOffset`** accumulates **per Order ID** the **sum of line amounts** (already sign-adjusted: purchases flow uses negative line amounts; offset uses **`Math.abs(total)`** for the balancing positive row).

**Rules:**

- **One offset row per Order ID** when net ≠ 0 and an account row exists for **`payKey`** (standard: payment method string; digital: uses digital user account).
- Skip offset if **net $0** or **no matching payment account** for the offset (counts reported in summary).
- Offset row gets **`purchase-offset`** / **`digital-purchase-offset`** metadata type for dedup.

**Refunds / digital returns** use separate offset logic in `importRefundDetailsCsv` / `importDigitalReturnsCsv` (group by order, sum amounts, `physical-refund-offset` / `digital-return-offset`).

---

## 6. Refund matching (payment and data joins)

**Orders returns (`importRefundDetailsCsv`):**

- Optional **Order History** text in the same bundle builds **`orderId → Payment Method Type`** via `amzOrderIdToPaymentStringMapFromCsv` (standard file type only).
- Refund transaction row: account from `amzResolvePhysicalRefundAccountRow`.
- **Offset** row per order: same payment lookup; metadata `physical-refund-offset`.

**Digital returns (`importDigitalReturnsCsv`):**

- Optional **Digital Content Orders** CSV → `amzOrderIdToPaymentStringMapFromCsv` in digital mode; else **Use for Digital** account from **AMZ Import**.

**Column names** for refund/return files come from **AMZ Import** unified map (`amzGetSourceMapHeader`, `refund details.csv` / `digital returns.csv`) with fallbacks to Amazon defaults.

---

## 7. Payment types: extraction vs **AMZ Import** mappings

**Extraction (`analyzePaymentMethodsForOrderHistory`):**

- Parse standard Order History, apply **cutoff** and **Website** toggles (`AMZ_WHOLE_FOODS_WEBSITE`, skip panda01 / skip non-panda01).
- Collect **unique** trimmed strings from the mapped **Payment Method Type** column.

**Comparison:**

- Each string looked up in **`config.paymentAccounts`** (`amzLookupPaymentAccountRow`); UI gets configured vs missing + **Accounts** sheet suggestions (4-digit heuristic).

**Import (`importAmazonRecent`):**

- **Strict:** unknown payment type **stops** the import with an error asking to add the row on **AMZ Import** (no silent blank account for orders).

---

## 8. Metadata fields: how CSV values become `amazon` JSON

**Driven by sheet:** Rows in the CSV map with a non-empty **Metadata field name** build `metadataMapping` (per-key standard vs digital column or literal).

**Resolution:** `amzResolveMetadataColumnName` picks header for current file type, with digital fallback logic.

**Build:** `amzBuildAmazonMetadataObject` — if resolved “column” exists in CSV, read cell; numeric keys get **0** for empty; if header missing, value may be the **literal** string from mapping (e.g. `type` ← `purchase`).

**Multi-line digital orders:** `amzBuildAmazonMetadataObjectFromRows` sums columns that match total amount column for numeric keys; sets **`lineItemCount`**.

**Envelope:** Prefix `Imported by AmazonCSVImporter on <timestamp> ` plus `JSON.stringify({ amazon: … })` on write.

**Dedup** uses types/ids inside that `amazon` object (see §4).

---

## 9. Sort order and Metadata filter (why this approach)

**Implemented in** `amzApplyTransactionsSortAndFilterCore_`:

1. **Remove** any existing **basic filter** first — sorting while a filter is active can **block rows from moving** (“failure-prone” behavior called out in code comments).
2. **Sort** data rows (row 2..last) by **Date descending** — prefer **`Range.sort`** on the data range for performance vs full-sheet sort on large grids; fallback to **`sheet.sort`** with optional temporary freeze row 1.
3. **Create** a new basic filter on a range sized from **used columns + Metadata column**, with **whenTextContains** on the **import timestamp** substring so the sheet focuses on **this run’s** rows.

**Why not rely on other methods:** Full-sheet sort / sort with active filter / fragile dimension reads caused real failures; code uses explicit **dimension clamping**, `amzEnsureSheetGridCovers`, and defensive **`getMaxRows`/`getMaxColumns`** handling to avoid **“coordinates outside dimensions”** when applying the filter.

**Deferred mode:** Bundle chunks use `deferTransactionsSheetPostProcess`; `amzImportBundleTransactionsPostLog_` runs sort/filter once at `finalize`.

---

## 10. Error handling and missing values (especially refund dates)

**General:**

- Missing required **mappings / columns** → validation errors (`validateAmzImportConfig`, `amzValidateMappedCsvHeadersPresent`).
- **Order Date** / **Return Date** empty or unparseable → row skipped; may appear in capped “Skipped row detail” log (`amzLogSkippedCsvDataIfUnderCap_`).
- **Payment** missing on standard order → **hard stop** (see §7).

**Refund Details transaction date** — `amzResolveRefundDetailsOrderDate_`:

1. Try mapped **Refund Date** → `amzParseAmazonCsvDateLoose_` (empty / invalid → `null`; **no “today” substitution** in parser).
2. If not usable, try **Creation Date** the same way.
3. If still `null`, row is skipped as invalid refund date (`importRefundDetailsCsv`); counted in **`skippedInvalidRefundDate`**.

Strings like **“Not Applicable”** fail `Date` parse → `null` → falls through to **Creation Date** when present; if both fail, row is skipped (not dated to “today”).

---

## Section index (original request)

1. One file at a time / payload — §1  
2. Client date trim before server — §2  
3. AMZ Import constants — §3  
4. Duplicate detection — §4  
5. Offset calculation — §5  
6. Refund matching — §6  
7. Payment extraction vs AMZ Import — §7  
8. Metadata field selection — §8  
9. Filter and sort — §9  
10. Missing values / refund dates — §10  

---

*Document reflects `amazonorders.gs` and `AmazonOrdersSidebar.html` as implemented. Update this file when behavior changes.*
