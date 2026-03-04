// Filter refresh speed test (onEdit). Optimized: only touches Quick Search Match column.
// Remove this file or delete the onEdit function when done testing.
// If you added onEdit elsewhere, remove the duplicate so only this one runs.

const FILTER_TEST_SHEET_NAME = "Transactions";
const FILTER_TEST_DEBOUNCE_SEC = 2;  // Set to 0 to disable debounce and run on every edit.

/**
 * Simple onEdit trigger: re-applies filter on the Match column only (faster than looping all columns).
 * Logs timing to Executions (View > Executions). Debounces so rapid edits only trigger once per FILTER_TEST_DEBOUNCE_SEC.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getSheetName() !== FILTER_TEST_SHEET_NAME) return;

  if (FILTER_TEST_DEBOUNCE_SEC > 0) {
    const cache = CacheService.getScriptCache();
    const key = "filterTestLastRun";
    if (cache.get(key)) return;
    cache.put(key, "1", FILTER_TEST_DEBOUNCE_SEC);
  }

  const t0 = Date.now();
  const criteriaColIndex = getQuickSearchCriteriaColCached();
  const t1 = Date.now();
  if (criteriaColIndex == null || criteriaColIndex < 2) return;

  const filter = sheet.getFilter();
  const t2 = Date.now();
  if (!filter) return;

  const matchCol1Based = criteriaColIndex - 1;
  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE", ""]).build();
  const t3 = Date.now();
  try {
    filter.setColumnFilterCriteria(matchCol1Based, criteria);
  } catch (err) {
    return;
  }
  const t4 = Date.now();

  Logger.log(
    "onEdit filter: getCriteriaCol " + (t1 - t0) + " ms, getFilter " + (t2 - t1) + " ms, " +
    "buildCriteria " + (t3 - t2) + " ms, setColumnFilterCriteria " + (t4 - t3) + " ms, total " + (t4 - t0) + " ms"
  );
}
