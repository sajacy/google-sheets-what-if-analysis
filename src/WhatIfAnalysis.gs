/** 
 * What-If Analysis 2022
 * This add-on stores state in a singleton document property, named `dt_`. 
 * It contains a map of JSON metadata for all the DataTables, with the structure: 
 * {
 *   "[named range]": {
 *     "sheet": [selected range].getSheet().getName()
 *     "range": [selected range].getA1Notation()
 *     "rowinput": [user input string: target cell, A1 notation]
 *     "colinput": [user input string: target cell, A1 notation]
 *   },
 *   ...
 * }
 * 
 * where [named range] and it's corresponding value are created by `create_` function.
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
    if (true) {
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
function importDataTables(json, reorderOnly = false) {
  const dt = Try(() => JSON.parse(json));
  if (!Array.isArray(dt)) {
    throw new Error("Import must be a valid JSON array of DataTables.");
  }
  let savedKeys;
  // Encode the config fields in a fixed order so JSON property order is irrelevant.
  const orderKey = config => JSON.stringify(config == null ? null : [
    config.name, config.sheet, config.range, config.colinput, config.rowinput,
  ]);
  if (reorderOnly === true) {
    savedKeys = new Set(fetchDataTables().map(orderKey));
    const incomingKeys = new Set(dt.map(orderKey));
    if ([...savedKeys].some(key => !incomingKeys.has(key))) {
      throw new Error("The saved tables changed. Revert to the last save and try again.");
    }
  }
  dt.forEach((config, idx) => {
    if (savedKeys?.has(orderKey(config))) return;
    try {
      validateConfig_(config);
    } catch (err) {
      throw new Error(`DataTable ${idx + 1}: ${err.message}`);
    }
  });
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(dt));
  return dt;
}

/**
 * Sidebar Button: Focus on Data Table Range, Summary on click 
 */
function focusDataTable(idx) {
  const dt = fetchDataTables();
  const config = dt[idx];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = config.sheet ? ss.getSheetByName(config.sheet) : ss.getActiveSheet();
  const range = resolveRange_(config.range, sheet, ss);
  if (range) range.activate();
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
  if (str.includes("!")) {
    return Try(() => ss.getRange(str));
  }
  // Relative address must resolve to data table's sheet `s` first
  return Try(() => s.getRange(str)) || Try(() => ss.getRange(str));
}

function validateConfig_(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!config || typeof config !== "object" || Array.isArray(config)) {
    throw new Error("A DataTable must be a JSON object.");
  }
  if (typeof config.name !== "string" || !config.name.trim()) {
    throw new DtConfigInputError(config, "name");
  }
  for (const field of ["sheet", "range", "colinput", "rowinput"]) {
    if (config[field] != null && typeof config[field] !== "string") {
      throw new Error(`${field} must be a text reference.`);
    }
  }

  // check sheet: backwards compatible with legacy: check for "config.sheet"
  const s = Try(() => {
    if (config.sheet) return ss.getSheetByName(config.sheet);
    if (config.range && String(config.range).includes("!")) {
      return ss.getRange(config.range.trim()).getSheet();
    }
    return ss.getActiveSheet();
  });
  if (!s) { throw new DtConfigInputError(config, "sheet"); }

  // check data table range: relative addresses belong to sheet `s`
  const range = resolveRange_(config.range, s, ss);
  if (!range) { throw new DtConfigInputError(config, "range"); }

  // check inputs: relative cell addresses resolve to sheet `s` first
  const colinput = resolveRange_(config.colinput, s, ss);
  const rowinput = resolveRange_(config.rowinput, s, ss);
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
  if (!colinput && !rowinput) { throw new DtConfigInputError(config, "inputs"); }
  if (range.getNumRows() < 2 || range.getNumColumns() < 2) {
    throw new Error("Range must span at least two rows and two columns.");
  }

  Logger.log(`Validated table: ${config.name}, range: ${range.getA1Notation()}`);
  return {config, dtRange: range, colInputCell: colinput, rowInputCell: rowinput};
}

class DtConfigInputError extends Error {
  constructor(config, reason) {
    switch(reason) {
      case "sheet": super(`Sheet (${config.sheet}) does not exist.`); break;
      case "name":  super(`Name (${config.name}) must be non-empty.`); break;
      case "range":  super(`Range (${config.range}) does not refer to a valid range.`); break;
      case "inputs": super(`Neither rowinput (${config.rowinput}) nor colinput (${config.colinput}) refer to valid ranges.`); break;
      default: super(`Config (${JSON.stringify(config)}) did not parse correctly.`)
    }
  }
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
        const range = item.sheet ? safeA1Notation_(item.sheet, item.range) : item.range;
        return {
          "name": key.startsWith("DataTable_") ? key.slice(10) : key,
          "sheet": item.sheet,
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
