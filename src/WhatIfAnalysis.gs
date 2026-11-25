/**
 * @OnlyCurrentDoc
 */
/**
 * What-If Analysis 2026
 * This add-on stores state in a singleton document property, named `dt2_`.
 * It is an array of JSON metadata for all the DataTables, with the structure:
 * [
 *   {
 *     "name": [user input string, arbitrary],
 *     "range": [selected range, either Named Range or A1 notation fully qualified with sheet],
 *     "rowinput": [selected cell, either Named Range or A1 notation fully qualified with sheet],
 *     "colinput": [selected cell, either Named Range or A1 notation fully qualified with sheet]
 *   },
 *   ...
 * ]
 *
 * It automatically migrates v1 configuration to the v2 format.
 * For reference, v1 format was:
 * {
 *   "[named range]": {
 *     "sheet": [selected range].getSheet().getName()
 *     "range": [selected range].getA1Notation()
 *     "rowinput": [user input string: target cell, A1 notation]
 *     "colinput": [user input string: target cell, A1 notation]
 *   },
 *   ...
 * }
 */
const DATATABLE_KEY = 'dt_';
const DATATABLE_KEY_V2 = 'dt2_'

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('What-If Analysis')
    .addItem('Open Sidebar', 'sidebar_')
  .addToUi();

  // DEPRECATED: retain for migration
  if (e && e.authMode != ScriptApp.AuthMode.NONE) {
    const hasV1 = !!PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY);
    if (hasV1) {
      ui.createMenu('What-If Analysis')
      .addItem('Open Sidebar', 'sidebar_')
      .addSubMenu(
        ui.createMenu('Legacy Version')
        .addItem('Refresh Data Tables', 'refresh_')
        .addItem(`Delete Legacy Version`, 'deleteLegacy_'))
      .addToUi();
    }
  }
}

/**
 * Menu Item: Sidebar
 */
function sidebar_() {
  SpreadsheetApp.getUi()
  .showSidebar(
    HtmlService
    .createHtmlOutputFromFile('sidepanel.html') // TODO: rename
    .setTitle('What-If Analysis'));
}

/**
 * (Legacy) Menu Item: Delete Legacy property
 */
function deleteLegacy_() {
  PropertiesService.getDocumentProperties().deleteProperty(DATATABLE_KEY);
}

/**
 * (Legacy) Menu Item: Refresh All Datatables
 */
function refresh_() {
  recalcDataTable(Infinity);
}

/**
 * Sidebar Button: Recalculate
 *
 * idx: int, can be Infinity to recalculate all
 */
function recalcDataTable(idx) {
  const dt = fetchDataTables();
  if (idx === Infinity) {
    dt.forEach(config => datatables_(config));
  } else if (Number.isInteger(idx) && dt[idx]) {
    datatables_(dt[idx]);
  }
}

/**
 * Sidebar Button: Load (and also on sidebar init)
 *
 * returns the full config array
 */
function fetchDataTables() {
  const jsonStrV2 = PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY_V2);
  if (!!jsonStrV2) {
    // jsonV2 already exists, prefer it: v1 data tables may no longer be usable.
    return JSON.parse(jsonStrV2);
  } else {
    // auto-migrate v1 to v2
    const rawV1 = PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY);
    const jsonV1 = rawV1 ? JSON.parse(rawV1) : {};
    const jsonV2 = convertToV2_(jsonV1);

    // store jsonV2
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(jsonV2));
    return jsonV2;
  }
}

/**
 * Sidebar Button: Save, Delete
 *
 * idx: int, can be -1 to indicate a new config
 * config: object, can be null to indicate deletion
 *
 * Can throw an exception due to validateConfig_
 */
function upsertDataTable(idx, config) {
  const dt = fetchDataTables();
  if (idx !== -1 && (!Number.isInteger(idx) || !dt[idx])) {
    throw new Error("DataTable no longer exists. Reload the sidebar.");
  }
  if (config === null) {
    if (idx === -1) throw new Error("Select a saved DataTable to delete.");
    dt.splice(idx, 1);
  } else {
    validateConfig_(config);
    if (idx === -1) {
      dt.push(config);
    } else {
      dt[idx] = config;
    }
  }
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(dt));
  return dt;
}

/**
 * Sidebar Button: Import All
 *
 * json: a full array of dt configs
 * reorderOnly: keep every saved config and skip range validation for existing configs and copies
 *
 * Can throw an exception due to validateConfig_
 */
function importDataTables(json, reorderOnly = true) {
  const dt = Try(() => JSON.parse(json));
  if (!Array.isArray(dt)) {
    throw new Error("Import must be a valid JSON array of DataTables.");
  }
  // Encode the config fields in a fixed order so JSON property order is irrelevant.
  const orderKeyFn = (config) => JSON.stringify(config == null ? null : [
    config.name, config.range, config.colinput, config.rowinput,
  ]);
  const savedKeys = new Set(fetchDataTables().map(orderKeyFn));
  if (reorderOnly === true) {
    const incomingKeys = new Set(dt.map(orderKeyFn));
    if ([...savedKeys].some(key => !incomingKeys.has(key))) {
      throw new Error("The saved state changed. Reload and try again.");
    }
  } else {
    for (const [idx, config] of dt.entries()) {
      if (savedKeys.has(orderKeyFn(config))) continue;
      try {
        validateConfig_(config);
      } catch (err) {
        throw new Error(`DataTable (index ${idx+1} of ${dt.length}): ${err.message}`);
      }
    }
  }
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(dt));
  return dt;
}

/**
 * Sidebar Button: Focus on Data Table Range, Summary on click
 */
function focusDataTable(idx) {
  const config = fetchDataTables()[idx];
  if (config) focusDataTableRange(config, "range");
}

/** Focus one editor reference without requiring a complete, saved config. */
function focusDataTableRange(config, field) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const tableRange = resolveRange_(config.range, sheet, ss);
  const inputSheet = tableRange?.getSheet() || sheet;
  const target = field === "range" ? tableRange : resolveRange_(config[field], inputSheet, ss);
  if (!target) throw new Error("Enter a valid range to focus on.");
  target.activate();
  SpreadsheetApp.flush();
}

/**
 * Sidebar Button: Capture Range (must already be selected)
 */
function captureRange() {
  const activeRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  return safeA1Notation_(activeRange.getSheet().getName(), activeRange.getA1Notation());
}

/**
 * Returns `null` if range is not found
 */
function resolveRange_(ref, s, ss) {
  if (!ref) return null;
  const str = String(ref).trim();
  if (!str) return null;
  // Resolve A1 notation before trying a workbook named range.
  return Try(() => s.getRange(str)) || Try(() => ss.getRangeByName(str));
}

function validateConfig_(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getActiveSheet();

  if (!config || typeof config !== "object" || Array.isArray(config)) {
    throw new Error("A DataTable must be a JSON object.");
  }
  for (const field of ["range", "colinput", "rowinput", "name"]) {
    if (config[field] != null && typeof config[field] !== "string") {
      throw new Error(`${field} must be a text reference.`);
    }
  }
  const range = resolveRange_(config.range, s, ss);
  if (!range) {
    throw new Error(`Range (${config.range}) does not refer to a valid range.`);
  }
  if (range.getNumRows() < 2 || range.getNumColumns() < 2) {
    throw new Error("Range must span at least two rows and two columns.");
  }

  const inputSheet = range.getSheet();
  const colinput = resolveRange_(config.colinput, inputSheet, ss);
  const rowinput = resolveRange_(config.rowinput, inputSheet, ss);
  for (const [label, ref, input] of [
    ["Column Input", config.colinput, colinput],
    ["Row Input", config.rowinput, rowinput],
  ]) {
    if (!ref || !ref.trim()) continue;
    if (!input || input.getNumRows() !== 1 || input.getNumColumns() !== 1) {
      throw new Error(`${label} (${ref}) must refer to one valid cell.`);
    }
    if (input.getSheet().getSheetId() === range.getSheet().getSheetId() &&
        input.getRow() >= range.getRow() && input.getRow() <= range.getLastRow() &&
        input.getColumn() >= range.getColumn() && input.getColumn() <= range.getLastColumn()) {
      throw new Error(`${label} (${ref}) must be outside the DataTable range.`);
    }
  }
  if (!colinput && !rowinput) {
    throw new Error(`Neither rowinput (${config.rowinput}) nor colinput (${config.colinput}) refer to valid ranges.`);
  }
  return {config, dtRange: range, colInputCell: colinput, rowInputCell: rowinput};
}

/** Helper functions for the API */
function safeA1Notation_(sheetName, a1part) {
  if (!sheetName) return a1part;
  return "'" + String(sheetName).replace(/'/g, "''") + "'!" + a1part;
}

function Try(fn, defaultValue = null) {
  try { return fn(); } catch { return defaultValue; }
}

function convertToV2_(jsonV1) {
  if (!jsonV1 || typeof jsonV1 !== 'object') return [];
  return Object.keys(jsonV1)
    .map(function (key) {
      return Try(() => {
        const item = jsonV1[key];
        if (!item) return null;
        const range = safeA1Notation_(item.sheet, item.range);
        return {
          "name": key.startsWith("DataTable_") ? key.slice(10) : key,
          "range": range,
          "colinput": item.colinput,
          "rowinput": item.rowinput
        };
      }, null);
    }).filter(_ => !!_);
}

function datatables_(config) {
  const {dtRange, colInputCell, rowInputCell} = validateConfig_(config);
  datatablesInternal_(dtRange, colInputCell, rowInputCell);
}

/** Working implementation */
function datatablesInternal_(dtRange, colInputCell, rowInputCell) {
  const s = dtRange.getSheet();
  const range_row = dtRange.getRow();
  const range_col = dtRange.getColumn();
  const num_rows = dtRange.getNumRows();
  const num_columns = dtRange.getNumColumns();
  const output_values = [];

  // A scenario can fail after changing an input. Backup its formula or value first.
  const originals = [colInputCell, rowInputCell].filter(Boolean).map(input => ({
    input,
    formula: input.getFormula(),
    value: input.getValue(),
  }));
  try {
    if (rowInputCell && colInputCell) {
      /* bivariate (single formula)
         | formula    | row_value_1 | row_value_2 |
         |------------|-------------|-------------|
         |col_value_1 |
         |col_value_2 |
         |col_value_3 |

         output.getValues() are single cells for each CxR combination:
         [ [ formula_c1_r1 ] ]
         [ [ formula_c1_r2 ] ]
         [ [ formula_c2_r1 ] ]
         [ [ formula_c2_r2 ] ]
         [ [ formula_c3_r1 ] ]
         [ [ formula_c3_r2 ] ]
      */
      var col_test_values = s.getRange(range_row + 1, range_col, num_rows - 1, 1).getValues().flat();
      var row_test_values = s.getRange(range_row, range_col + 1, 1, num_columns - 1).getValues().flat();
      var output = s.getRange(range_row, range_col, 1, 1);

      for (var i = 0; i < col_test_values.length; i++) {
        output_values[i] = [];
        for (var j = 0; j < row_test_values.length; j++) {
          colInputCell.setValue(col_test_values[i]);
          rowInputCell.setValue(row_test_values[j]);
          SpreadsheetApp.flush();
          output_values[i][j] = output.getValue();
        }
      }
    } else if (colInputCell) {
      /* univariate along column, potentially multi-output
         |             | formula_1 | formula_2 |
         |-------------|-----------|------ ----|
         |test_value_1 |
         |test_value_2 |
         |test_value_3 |

         output.getValues() are columns for each test value:
         [ [ formula_1_test_value_1 , formula_2_test_value_1 ] ]
         [ [ formula_1_test_value_2 , formula_2_test_value_2 ] ]
         [ [ formula_1_test_value_3 , formula_2_test_value_3 ] ]
      */
      var col_test_values = s.getRange(range_row + 1, range_col, num_rows - 1, 1).getValues().flat();
      var output = s.getRange(range_row, range_col + 1, 1, num_columns - 1);
      for (var i = 0; i < col_test_values.length; i++) {
        colInputCell.setValue(col_test_values[i]);
        SpreadsheetApp.flush();
        output_values[i] = output.getValues().flat();
      }
    } else if (rowInputCell) {
      /* univariate along row, potentially multi-output
         |          | test_value_1 | test_value_2 | ... | test_value_N |
         |----------|--------------|--------------|-----|--------------|
         |formula_1 |
         |formula_2 |

         output.getValues() are columns for each test value:
         [ [ formula_1_test_value_1 ], [ formula_2_test_value_1 ] ]
         [ [ formula_1_test_value_2 ], [ formula_2_test_value_2 ] ]
         [ [ formula_1_test_value_3 ], [ formula_2_test_value_3 ] ]
      */
      var row_test_values = s.getRange(range_row, range_col + 1, 1, num_columns - 1).getValues().flat();
      var output = s.getRange(range_row + 1, range_col, num_rows - 1, 1);
      for (var j = 0; j < row_test_values.length; j++) {
        rowInputCell.setValue(row_test_values[j]);
        SpreadsheetApp.flush();
        var colOut = output.getValues().flat();
        for (var i = 0; i < colOut.length; i++) {
          if (!output_values[i]) output_values[i] = [];
          output_values[i][j] = colOut[i];
        }
      }
    }
  } finally {
    originals.forEach(({input, formula, value}) => {
      if (formula) {
        input.setFormula(formula);
      } else {
        input.setValue(value);
      }
    });
    SpreadsheetApp.flush();
  }

  // Gets all the test values in a batch operation, so that expensive getCell and setValue calls are not made in the for-loop.
  const resultCells = s.getRange(range_row + 1, range_col + 1, num_rows - 1, num_columns - 1);
  resultCells.setValues(output_values);
  SpreadsheetApp.flush();
}
