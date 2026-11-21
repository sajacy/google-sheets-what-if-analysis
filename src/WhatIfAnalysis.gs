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
  ui.createAddonMenu()
    .addSubMenu(
      ui.createMenu('Legacy Version')
      .addItem('Create Data Table', 'create_')
      .addItem('Refresh Data Tables', 'refresh_')
      .addItem('Delete All Data Tables', 'delete_'))
    .addItem('Open Sidebar', 'sidepanel_')
    .addToUi();
    
  // Avoid accessing properties if the user hasn't yet run the add-on (via menu items)
  // https://developers.google.com/apps-script/reference/script/auth-mode
  if (e && e.authMode != ScriptApp.AuthMode.NONE) {
    // initialize document state for datatables
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY) || '{}');
  }
}

function sidepanel_() {
  SpreadsheetApp.getUi()
  .showSidebar(
    HtmlService
    .createHtmlOutputFromFile('sidepanel.html')
    .setTitle('What-If Analysis'));
}

/** Datatables V2 API
 * fetchDataTables
 * recalcDataTable
 * upsertDataTable
 * focusDataTable
 * captureRange
 * 
 * validateConfig
 * 
 * ScriptError: TypeError: Cannot read property 'replace' of undefined
 */
function fetchDataTables() {
  const jsonStrV2 = PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY_V2);
  if (!!jsonStrV2) {
    // jsonV2 already exists, prefer it: v1 data tables may no longer be usable.
    return JSON.parse(jsonStrV2);
  } else {
    const rawV1 = PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY);
    const jsonV1 = rawV1 ? JSON.parse(rawV1) : {};

    // convert to jsonV2
    const jsonV2 = convertToV2_(jsonV1);

    // store jsonV2
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(jsonV2));

    return jsonV2;
  }
}

function recalcDataTable(idx) {
  const dt = fetchDataTables();
  if (idx === undefined || idx === null || idx === 'all') {
    dt.forEach(config => datatables_(config));
    return dt;
  } else if (dt[idx]) {
    datatables_(dt[idx]);
    return dt[idx];
  }
  return null;
}

function refresh_() {
  recalcDataTable('all');
}

function test_recalcDataTable_idx0() {
  recalcDataTable(0);
}

/** 
 * idx: int, can be -1 to indicate a new config
 * config: object, can be null to indicate deletion
 * 
 * Function can throw an exception due to validateConfig
 */
function upsertDataTable(idx, config) {
  const dt = fetchDataTables();
  if(!!config && validateConfig(config)) {
    // config is upsertable, do the upsert:
    if (idx == -1) {
      dt.push(config);
    } else {
      dt[idx] = config;
    }
  } else {
    // config was null, delete it at idx:
    dt.splice(idx,1);
  }
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY_V2, JSON.stringify(dt));
  return dt;
}

function focusDataTable(idx) {
  const dt = fetchDataTables();
  const config = dt[idx];
  SpreadsheetApp.getActiveSpreadsheet().getRange(config.range).activate();
  SpreadsheetApp.flush();
}

function captureRange() {
  const activeRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  return safeA1Notation_(activeRange.getSheet().getName(), activeRange.getA1Notation());
}

function resolveRange_(ref, s, ss) {
  if (!ref) return null;
  const str = String(ref).trim();
  if (str.includes("!")) {
    return Try(() => ss.getRange(str));
  }
  // Relative address must resolve to data table's sheet `s` first
  return Try(() => s.getRange(str)) || Try(() => ss.getRange(str));
}

function validateConfig(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // check config has a name (required for V2)
  if (!config.name) { throw new DtConfigInputError(config, "name"); }

  // check sheet: backwards compatible with legacy: check for "config.sheet"
  const s = Try(() => {
    if (config.sheet) return ss.getSheetByName(config.sheet);
    if (config.range && String(config.range).includes("!")) {
      return ss.getRange(config.range).getSheet();
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
  if (!colinput && !rowinput) { throw new DtConfigInputError(config, "inputs"); }

  Logger.log(`Validated table: ${config.name}, range: ${range.getA1Notation()}`);
  return { "config": config, "range": range, "colinput": colinput, "rowinput": rowinput };
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
    }).filter(o => o);
}

/**
 * New features to support:
 * 
 *  - Side Panel to manage DataTables
 *  - Nicer UI for specifying ranges
 *  - Change focus to DataTable
 *  - Recalculate individually
 *  - Remove dependency on NamedRanges (for existence/deletion)
 *  - Copying DataTables to new Sheet (DocumentProperties don't copy with File > Make a Copy)
 *  - allow user-specified names (Migrate to keyless dt_ object: array only)
 * 
 * Auto-Refresh 
 *  - Watch for new DataTable inputs and recalculate the outputs. 
 */

function test_datatables() {
  [
    {
    "name": "Bivariate",
    "range": "H4:K12",
    "sheet": "Original 2018 Tax Calculator",
    "colinput": "A2",
    "rowinput": "A4"
  }, 
  {
    "name": "Univariate",
    "range": "H16:I24",
    "sheet": "Original 2018 Tax Calculator",
    "colinput": "A2",
    "rowinput": null
  }, 
  {
    "name": "BadSheet",
    "range": "H16:I24",
    "sheet": "Non-existent Sheet",
    "colinput": "A2",
    "rowinput": null
  },
    {
    "name": "BadRange",
    "range": "Bad-Range-Ref",
    "sheet": "Original 2018 Tax Calculator",
    "colinput": "A2",
    "rowinput": null
  },
    {
    "name": "BadInputs",
    "range": "H16:I24",
    "sheet": "Original 2018 Tax Calculator",
    "colinput": "Another-bad-range-ref",
    "rowinput": "Yet-another-bad-range-ref"
  },
    {
    // "name": "Missing Name",
    "range": "H16:I24",
    "sheet": "Original 2018 Tax Calculator",
    "colinput": "A2",
    "rowinput": null
  }].forEach(function (example, i) {
    try {
      Logger.log(`Test ${i}:`);
      datatables_(example);
      Logger.log("OK");
    } catch (e) {
      Logger.log(e);
    }
  });
}

function datatables_(config) {
  const { range, colinput, rowinput } = validateConfig(config);
  datatablesInternal_(range, colinput, rowinput);
}

/** Working implementation */
function datatablesInternal_(datatable, colinput, rowinput) {
  var s = datatable.getSheet();
  var range_row = datatable.getRow();
  var range_col = datatable.getColumn();
  var num_rows = datatable.getNumRows();
  var num_columns = datatable.getNumColumns();
  var output_values = [];

  if (rowinput && colinput) {
    // bivariate
    var colOriginal = colinput.getValue();
    var rowOriginal = rowinput.getValue();
    var col_test_values = s.getRange(range_row + 1, range_col, num_rows - 1, 1).getValues().flat();
    var output = s.getRange(range_row, range_col, 1, 1);
    var row_test_values = s.getRange(range_row, range_col + 1, 1, num_columns - 1).getValues().flat();

    for (var i = 0; i < col_test_values.length; i++) {
      output_values[i] = [];
      for (var j = 0; j < row_test_values.length; j++) {
        colinput.setValue(col_test_values[i]);
        rowinput.setValue(row_test_values[j]);
        SpreadsheetApp.flush();
        output_values[i][j] = output.getValue();
      }
    }
    colinput.setValue(colOriginal);
    rowinput.setValue(rowOriginal);
    SpreadsheetApp.flush();
  } else if (colinput) {
    // univariate along column, potentially multi-output
    var colOriginal = colinput.getValue();
    var col_test_values = s.getRange(range_row + 1, range_col, num_rows - 1, 1).getValues().flat();
    var output = s.getRange(range_row, range_col + 1, 1, num_columns - 1);
    for (var i = 0; i < col_test_values.length; i++) { 
      colinput.setValue(col_test_values[i]);
      SpreadsheetApp.flush();
      output_values[i] = output.getValues().flat();
    }
    colinput.setValue(colOriginal);
    SpreadsheetApp.flush();
  } else if (rowinput) {
    // univariate along row, potentially multi-output
    var rowOriginal = rowinput.getValue();
    var row_test_values = s.getRange(range_row, range_col + 1, 1, num_columns - 1).getValues().flat();
    var output = s.getRange(range_row + 1, range_col, num_rows - 1, 1);
    for (var j = 0; j < row_test_values.length; j++) {
      rowinput.setValue(row_test_values[j]);
      SpreadsheetApp.flush();
      var colOut = output.getValues().flat();
      for (var i = 0; i < colOut.length; i++) {
        if (!output_values[i]) output_values[i] = [];
        output_values[i][j] = colOut[i];
      }
    }
    rowinput.setValue(rowOriginal);
    SpreadsheetApp.flush();
  }

  // Gets all the test values in a batch operation, so that expensive getCell and setValue calls are not made in the for-loop.
  var dt_range_to_be_set = s.getRange(range_row + 1, range_col + 1, num_rows - 1, num_columns - 1);
  dt_range_to_be_set.setValues(output_values);
  SpreadsheetApp.flush();
}

/**
 * Create an UpdateCellsRequest for a single cell.
 * 
 * @param {Range} range The range to update
 * @param {object} value The value to write to the cell
 * @return {UpdateCellsRequest} A request object
 */
function updateCell_(range, value) {
  const extendedValue = (typeof value === 'number') ? 
    { numberValue: value } : 
    { stringValue: value };
  const o = SheetsApi.newUpdateCellsRequest();
  o.fields = 'userEnteredValue';
  o.range = rangeToGrid_(range);
  o.rows = [{ values: [{ userEnteredValue: extendedValue }] }]
  return { 'updateCells': o };
}

/** 
 * Creates a 1x1 GridRange for the specified range and offsets 
 * 
 * @param {Range} range The range representing the full data table
 * @param {number} offsetCols The column index offset
 * @param {number} offsetRows The row index offset
 */
function rangeToGrid_(range, offsetCols = 0, offsetRows = 0) {
  // find range reference indices
  const colIndex = range.getColumn()-1;
  const rowIndex = range.getRow()-1;

  const o = SheetsApi.newGridRange();
  o.sheetId = range.getSheet().getSheetId();
  o.startColumnIndex = colIndex + offsetCols;
  o.startRowIndex = rowIndex + offsetRows;
  // end indices are exclusive, so add 1
  o.endColumnIndex = colIndex + offsetCols + 1;
  o.endRowIndex = rowIndex + offsetRows + 1;
  return o;  
}

function copyCell_(range, toCol, toRow, fromCol = 0, fromRow = 0) {
  const o = SheetsApi.newCopyPasteRequest();
  o.source = rangeToGrid_(range, fromCol, fromRow);
  o.destination = rangeToGrid_(range, toCol, toRow);
  o.pasteType = 'PASTE_VALUES';
  o.pasteOrientation = 'NORMAL';
  return { 'copyPaste': o };
}

/** 
 * An internal implementation using batchUpdate.
 * 
 * @param {Range} datatable The range representing the full data table
 * @param {Range} colinput The cell to write column test inputs (nullable)
 * @param {Range} rowinput The cell to write row test inputs (nullable)
 */
function datatablesInternalV2_(datatable, colinput, rowinput) {
  const requests = [];
  const values = datatable.getValues();
  const DUMMY = true;
  
  if (colinput && rowinput) {
    const colOriginal = colinput.getValue();
    const rowOriginal = rowinput.getValue();

    /*
     * Because the inputs are actually declared along the transposed axes,
     * ri := col index
     * ci := row index
     *   out | ri=0 | ri=1 | ri=2
     *   ----+------+------+------
     *   ci=0| ...  | ...  |  ...
     *   ci=1| ...  | ...  | datatable[row=2][col=3]
     *   ci=2| ...  | ...  |  ...
     * 
     * also, note the +1 offset due to the test value column and row.
     */
    const biColTestValues = values.map(row => row[0]).slice(1);
    const biRowTestValues = values[0].slice(1);
    biColTestValues.forEach((cv, ci) => {
      requests.push(updateCell_(colinput, cv));
      biRowTestValues.forEach((rv, ri) => {
        requests.push(updateCell_(rowinput, rv));
        requests.push(copyCell_(datatable, ri + 1, ci + 1));
      });
    });
    
    // restore originals
    requests.push(updateCell_(colinput, colOriginal))
    requests.push(updateCell_(rowinput, rowOriginal))

  } else if (colinput) {
    /*
     * Because column inputs are actually declared along rows,
     * oi := col index
     * ci := row index
     *       | oi=0 | oi=1 | oi=2
     *   ----+------+------+------
     *   ci=0| ...  | ...  |  ...
     *   ci=1| ...  | ...  | datatable[row=2][col=3]
     *   ci=2| ...  | ...  |  ...
     * 
     * also, note the +1 offset due to the test value column and output row.
     */
    const colOriginal = colinput.getValue();
    const colTestValues = values.map(row => row[0]).slice(1);
    const outputs = values[0].slice(1).map(v => DUMMY);
    colTestValues.forEach((cv, ci) => {
      requests.push(updateCell_(colinput, cv));
      outputs.forEach((dummy, oi) => 
        requests.push(copyCell_(datatable, oi + 1, ci + 1, 1, 0)));
    });
    
    // restore original
    requests.push(updateCell_(colinput, colOriginal))
    
  } else if (rowinput) {
    /*
     * Because row inputs are actually declared along columns,
     * ri := col index
     * oi := row index
     *       | ri=0 | ri=1 | ri=2
     *   ----+------+------+------
     *   oi=0| ...  | ...  |  ...
     *   oi=1| ...  | ...  | datatable[row=2][col=3]
     *   oi=2| ...  | ...  |  ...
     * 
     * also, note the +1 offset due to the test value row and output column.
     */
    const rowOriginal = rowinput.getValue();
    const rowTestValues = values[0].slice(1);
    const outputs = values.map(row => row[0]).slice(1).map(v => DUMMY);
    rowTestValues.forEach((rv, ri) => {
      requests.push(updateCell_(rowinput, rv));
      outputs.forEach((dummy, oi) => 
        requests.push(copyCell_(datatable, ri + 1, oi + 1, 0, 1)));
    });
    
    // restore original
    requests.push(updateCell_(rowinput, rowOriginal))

  }
  batchUpdate_(requests);
}

function batchUpdate_(requests) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const batchRequest = SheetsApi.newBatchUpdateSpreadsheetRequest();
  batchRequest.requests = requests;
  //batchRequest.includeSpreadsheetInResponse = false;
  //batchRequest.responseIncludeGridData = true;
  Logger.log(JSON.stringify(batchRequest,null,2));
  const response = {}; // SheetsApi.Spreadsheets.batchUpdate(batchRequest, ssId);
  Logger.log(`BatchUpdate response: ${JSON.stringify(response,null,2)}`);
  return response;
}

function batchGet_(requests) {
  
}
