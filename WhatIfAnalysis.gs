var DATATABLE_KEY = 'dt_';

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('What-If Analysis')
    .addItem('Create Data Table', 'create_')
    .addItem('Refresh Data Tables', 'refresh_')
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
  
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY));
  var ui = SpreadsheetApp.getUi();
  var dt_range = SpreadsheetApp.getActiveRange();
  var config = false;
  
  if (dt_range.getNumColumns() < 2) {
    help_();
  } else if (dt_range.getNumColumns() > 2) {
    // 2D data-table: row and column inputs
    var result_colinput = ui.prompt("Specify Cell to Set with Values from Left Column", 'Specify the first input cell\nFor example, enter "A2" to set cell A2 with the values in the left column.', ui.ButtonSet.OK_CANCEL);
    if (result_colinput.getSelectedButton() == ui.Button.OK) {
      var result_rowinput = ui.prompt("Specify Cell to Set with Values from Top Row", 'Specify the second input cell\nFor example, enter "A4" to set cell A4 with the values in the top row.', ui.ButtonSet.OK_CANCEL);
      if (result_rowinput.getSelectedButton() == ui.Button.OK) {
        var output2d = dt_range.getCell(1,1);
        var rowinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_rowinput.getResponseText());
        var colinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_colinput.getResponseText());
        config = { "range": dt_range.getA1Notation(), "output": output2d.getA1Notation(), "rowinput": result_rowinput.getResponseText(), "colinput": result_colinput.getResponseText() };
      }
    }
  } else {
    // column inputs only
    var result_input = ui.prompt('Specify Model Input', 'Specify the (column) input cell.\nFor example, enter "A2" to set cell A2 with the values in the left column.', ui.ButtonSet.OK_CANCEL);
    if (result_input.getSelectedButton() == ui.Button.OK) {
      var input = SpreadsheetApp.getActiveSpreadsheet().getRange(result_input.getResponseText());
      var output = dt_range.getCell(1,2);
      config = { "range": dt_range.getA1Notation(), "output": output.getA1Notation(), "rowinput": null, "colinput": result_input.getResponseText() };
    }
  }

  if (config) {
    // actually do the work now:
    datatables_(config);
    
    // save named range and property to be able to refresh data
    var name = "DataTable_" + dt_range.getA1Notation().replace(/[^A-Z0-9]/g,"");
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange(name, dt_range);
    dt_[name] = config;
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
  }
}

function refresh_() {
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY));
  var ranges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  for (var i = 0; i < ranges.length; i++) {
    var name = ranges[i].getName();
    if (dt_[name]) {
      // book-keeping for cleanup
      dt_[name].exists = true;
      // update the config if the range has been moved:
      var currentRange = ranges[i].getRange();
      dt_[name].range = currentRange.getA1Notation();
      dt_[name].output = !!(dt_[name].rowinput) ? currentRange.getCell(1,1).getA1Notation() : currentRange.getCell(1,2).getA1Notation();
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

function datatables_(config) {
  // Note: We use getActiveSheet() instead of getActiveSpreadsheet() here,
  // because of its expanded getRange() methods.
  var s = SpreadsheetApp.getActiveSheet();
  var dt_range = s.getRange(config.range);
  
  if (!config.rowinput) {
    // For a Univariate Data Table (1 test value).
    // Such a Data Table is here used when there is only 1 column with output results.
    // The selected range matrix should have exactly 2 columns: 1 for test input values, and 1 for output results.
    var input = s.getRange(config.colinput);
    var original = input.getValue();
    var output = s.getRange(config.output);
    
    var topleft_cell_row = dt_range.getRow();
    var topleft_cell_col = dt_range.getColumn();
    var dt_range_to_be_set = s.getRange(
      topleft_cell_row+1,
      topleft_cell_col+1,
      dt_range.getNumRows()-1,
      dt_range.getNumColumns()-1
    );
    
    // Gets all the test values in a batch operation, so that expensive getCell and setValue calls are not made in the for-loop.
    var test_values_length = dt_range.getNumRows()-1;
    var test_values_array = s.getRange(
      topleft_cell_row+1,
      topleft_cell_col,
      test_values_length,
      1
    ).getValues().flat();
    var output_values = new Array(test_values_length);
    // Iterating over both test_values_array and output_values at the same time,
    // using the fact that they were created above with the same length.
    for (var i = 0; i < test_values_length; i++) {
      input.setValue(test_values_array[i]);
      // ...
      // Here the spreadsheet formulas will do their thing in the spreadsheet behind the scenes, calculating the output value.
      // ...
      // Must use [] around the value here, to make the output_values array into a matrix,
      // since the dt_range_to_be_set later requires a matrix as input.
      output_values[i] = [output.getValue()];
    }
    
    // Set all the output values at once in the spreadsheet, now when all the tests are complete.
    dt_range_to_be_set.setValues(output_values);
    // Reset the input value
    input.setValue(original);
    
  } else {
    // For a 2D (bivariate) Data Table (2 test values).
    // Such a Data Table is here used when there are 2 or more more columns with output results.
    // The selected range matrix should have 3 or more columns: 1 column for row test input values, plus 2 or more columns for column test input values.
    // Outputs will be filled into the remainding cells in the matrix, using the corresponding row and column test input values.
    var colinput = s.getRange(config.colinput);
    var rowinput = s.getRange(config.rowinput);
    var colOriginal = colinput.getValue();
    var rowOriginal = rowinput.getValue();
    var output = s.getRange(config.output);
    
    var topleft_cell_row = dt_range.getRow();
    var topleft_cell_col = dt_range.getColumn();
    var num_rows = dt_range.getNumRows();
    var num_columns = dt_range.getNumColumns();
    
    var col_test_values_array = s.getRange(
      topleft_cell_row+1,
      topleft_cell_col,
      num_rows-1,
      1
    ).getValues().flat();
    
    var row_test_values_array = s.getRange(
      topleft_cell_row,
      topleft_cell_col+1,
      1,
      num_columns-1
    ).getValues()[0];
    
    // Gets all the test values in a batch operation, so that expensive getCell and setValue calls are not made in the for-loop.
    var dt_range_to_be_set = s.getRange(
      topleft_cell_row+1,
      topleft_cell_col+1,
      num_rows-1,
      num_columns-1
    );
    var output_values = dt_range_to_be_set.getValues();
    for (var i = 0; i < output_values.length; i++) { 
      for (var j = 0; j < output_values[i].length; j++) {
        colinput.setValue(col_test_values_array[i]);
        rowinput.setValue(row_test_values_array[j]);
        output_values[i][j] = output.getValue();
      }
    }
    
    dt_range_to_be_set.setValues(output_values);
    
    // Reset the input values
    colinput.setValue(colOriginal);
    rowinput.setValue(rowOriginal);
  }
}
