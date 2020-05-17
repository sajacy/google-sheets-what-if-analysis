var DATATABLE_KEY = 'dt_';

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('What-If Analysis')
    .addItem('Create Data Table', 'create_')
    .addItem('Refresh Data Tables', 'refresh_')
    .addItem('Delete All Data Tables', 'delete_')
    .addItem('Help', 'help_')
    .addToUi();
    
  // Avoid accessing properties if the user hasn't yet run the add-on (via menu items)
  // https://developers.google.com/apps-script/reference/script/auth-mode
  if (e && e.authMode != ScriptApp.AuthMode.NONE) {
    // initialize document state for datatables
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY) || "{}");
  }
}

function help_() {
  SpreadsheetApp.getUi().alert(`
    Selected range must be at least 2 columns.

    For a univariate Data Table (1 input variable):
    Top left cell can contain any text column header for the test values which need to follow downwards on the same leftmost column. The header for the second column needs to be a formula giving the model output (based on the input variable), or reference a cell with the model output.

    For a 2D (bivariate) Data Table (2 input variables):
    Top left cell needs to be a formula giving the model output, or reference a cell with the model output. Furthermore, the leftmost column needs to have the various test values listed downwards. Also, the topmost row needs to have more than 1 column in addition to the leftmost column, and they need to contain test values.

    WhatIfAnalysis will in both cases fill in the remaining part of the matrix. It will insert test values into the cells which you specify as input variable(s), and fill in the matrix with the corresponding model output results.

    NB:

    For simple models then WhatIfAnalysis is not needed, and results can be obtained much faster with repeating a formula normally throughout the matrix (in-spreadsheet formulas are generally faster than scripts like this WhatIfAnalysis script). For a bivariat Data Table (starting at top left cell A1) based on for instance a model like =(ROW_VARIABLE_CELL + COL_VARIABLE_CELL * 2), then the formula =$A2+B$1*2 could be input to B2 (cell 2,2 in the matrix) and dragged downwards to fill the entire first column, and then dragged to the right, to fill the entire matrix.

    But in more complex models the output can be the result of a chain of formulas throughout various cells, and then WhatIfAnalysis is particularly useful. In some cases though, a conjoined formula can be reconstructed by going to the end output cell and into the formula there successively inserting the formulas contained in the cells it referenced. Then the approach mentioned in the simple model scenario above will work. This can give a significant speed increase (and allow for instantaneous and automatically refreshing values when changes occur), compared to using WhatIfAnalysis (which must be manually run to recalculate values). Although this approach is always theoretically possible, it is not always practically feasible (i.e. worth the effort). Typically in cases where several formulas in the chain of formulas specify ranges of cells which again include other formulas. Then WhatIfAnalysis comes handy.

    For the ultimately fastest approach with a complex model, consider representing it entirely in code, and have it perform the calculations there (without relying on the spreadsheet). To do that, go to Tools -> Script editor, and write your own script using JavaScript. Your custom function in your script should take in all your input data (no formulas) and do all the calculations in code (instead of using spreadsheet formulas). This way the script can perform all the calculations in memory, without interacting with the spreadsheet at every step like WhatIfAnalysis does (which takes some time every time input values are set in the spreadsheet and output values are recorded from it). When your script has calculated all the outputs it can merely output them all to the spreadsheet at once (thus reducing the interactions with the spreadsheet to the minimum).
  `);
}

function create_() {
  
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY)) || {};
  var ui = SpreadsheetApp.getUi();
  var dt_range = SpreadsheetApp.getActiveRange();
  var selected_cols = dt_range.getNumColumns();
  var config = false;
  
  if (selected_cols < 2) {
    help_();
  } else {
    // 2D data-table: row and column inputs
    var result_colinput = ui.prompt("Cell to Set with Values from Left Column",
                                    `Specify the column input cell
                                     For example, enter "A2" to set cell A2 with the values in the left column.`, ui.ButtonSet.OK_CANCEL);
    
    if (result_colinput.getSelectedButton() == ui.Button.OK) {
      SpreadsheetApp.getActiveSpreadsheet().getRange(result_colinput.getResponseText());
      
      if (selected_cols > 2) {
        var result_rowinput = ui.prompt("Cell to Set with Values from Top Row", 
                                        `Specify the row input cell.
                                         For example, enter "Sheet1!A4" to set cell A4 on Sheet1 with the values in the top row.
                                         Leave blank if the top row contains formulas for outputs.`, ui.ButtonSet.OK_CANCEL);
        if (result_rowinput.getSelectedButton() == ui.Button.OK) {
          if (result_rowinput.getResponseText()) {
            // allow blank, but verify if non-blank
            SpreadsheetApp.getActiveSpreadsheet().getRange(result_rowinput.getResponseText());
          }
          config = { "sheet": dt_range.getSheet().getName(), "range": dt_range.getA1Notation(), "rowinput": result_rowinput.getResponseText(), "colinput": result_colinput.getResponseText() };
        }
      } else {
        config = { "sheet": dt_range.getSheet().getName(), "range": dt_range.getA1Notation(), "rowinput": null, "colinput": result_colinput.getResponseText() };
      }
    }
  }

  if (config) {
    // actually do the work now:
    datatables_(config);
    
    // save named range and property to be able to refresh data
    var topleft = dt_range.getA1Notation().split(":")[0];
    var name = "DataTable_" + dt_range.getSheet().getName().replace(/[^A-Za-z0-9]/g,"") + "_" + topleft;

/*
   // cleanup old (pre-version 10) names
    SpreadsheetApp.getActiveSpreadsheet().getNamedRanges().forEach(function(nr) {
      if (nr.match(new RegExp("DataTable_" + topleft + "[^_]+$")) && dt_[nr]) {
          SpreadsheetApp.getActiveSpreadsheet().removeNamedRange(nr);
          delete dt_[nr];
      }
    });
*/
    
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange(name, dt_range);
    dt_[name] = config;
    
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
  }
}

function refresh_() {
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY)) || {};
  var ranges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  for (var i = 0; i < ranges.length; i++) {
    var name = ranges[i].getName();
    if (dt_ && dt_[name]) {
      // book-keeping for cleanup
      dt_[name].exists = true;
      // update the config if the range has been moved:
      var currentRange = ranges[i].getRange();
      dt_[name].sheet = currentRange.getSheet().getName();
      dt_[name].range = currentRange.getA1Notation();
      // re-evaluate the configured datatable
      datatables_(dt_[name]);
    }
  }
  
  // cleanup data tables if named range was deleted
  var keys = Object.keys(dt_);
  for (var i = 0 ; i < keys.length; i++) {
    if (!dt_[keys[i]].exists) {
      delete dt_[keys[i]];
    }
  }
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
}

function delete_() {
  var ui = SpreadsheetApp.getUi();
  var responseButton = ui.alert('Are you sure?', 'Are you sure you want to delete all Data Tables?', ui.ButtonSet.YES_NO);
  if (responseButton == ui.Button.YES) {
    var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY)) || {};
    var keys = Object.keys(dt_);
    for (var i = 0 ; i < keys.length; i++) {
      if (SpreadsheetApp.getActiveSpreadsheet().getNamedRanges().indexOf(keys[i]) > -1) {
        SpreadsheetApp.getActiveSpreadsheet().removeNamedRange(keys[i]);
        delete dt_[keys[i]];
      }
    }
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
  }
}

function datatables_(config) {
  // Note: We use getActiveSheet() instead of getActiveSpreadsheet() here,
  // because of its expanded getRange() methods.
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheet);
  var dt_range = s.getRange(config.range);
  var range_row = dt_range.getRow();
  var range_col = dt_range.getColumn();
  var num_rows = dt_range.getNumRows();
  var num_columns = dt_range.getNumColumns();
  var output_values = [];

  var colinput = s.getRange(config.colinput);
  var colOriginal = colinput.getValue();
  var col_test_values = s.getRange(range_row + 1, range_col, num_rows - 1, 1).getValues().flat();

  if (config.rowinput) {
    // bivariate
    var output = s.getRange(range_row, range_col, 1, 1);
    var row_test_values = s.getRange(range_row, range_col + 1, 1, num_columns - 1).getValues().flat();
    var rowinput = s.getRange(config.rowinput);
    var rowOriginal = rowinput.getValue();

    for (var i = 0; i < col_test_values.length; i++) {
      output_values[i] = []
      for (var j = 0; j < row_test_values.length; j++) {
        colinput.setValue(col_test_values[i]);
        rowinput.setValue(row_test_values[j]);
        output_values[i][j] = output.getValue();
      }
    }
    rowinput.setValue(rowOriginal); // reset row input value
  } else {
    // univariate, potentially multi-output
    var output = s.getRange(range_row, range_col + 1, 1, num_columns - 1);
    for (var i = 0; i < col_test_values.length; i++) { 
      colinput.setValue(col_test_values[i]);
      output_values[i] = output.getValues().flat();
    }
  }

  // Gets all the test values in a batch operation, so that expensive getCell and setValue calls are not made in the for-loop.
  var dt_range_to_be_set = s.getRange(range_row + 1, range_col + 1, num_rows - 1, num_columns - 1);
  dt_range_to_be_set.setValues(output_values);

  // Reset the input value
  colinput.setValue(colOriginal);
}
