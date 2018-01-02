function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('What-If Analysis')
  .addSubMenu(ui.createMenu("Data Tables")
              .addItem('Create', 'create_')
              .addItem('Refresh', 'refresh_'))
      .addToUi();
  
  // initialize document state for datatables
  PropertiesService.getDocumentProperties().setProperty("dt_", PropertiesService.getDocumentProperties().getProperty("dt_") || "{}");
}

function create_() {
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty("dt_"));
  var ui = SpreadsheetApp.getUi();
  var dt_range = SpreadsheetApp.getActiveRange();
  var config = false;
  
  if (dt_range.getNumColumns() < 2) {
    ui.alert("Range must be at least 2x2: input values on left column (and top row, for 2D datatable) and output values to the right / bottom");
  } else if (dt_range.getNumColumns() > 2) {
    // 2D data-table: row and column inputs
    // TODO: validate OK and CANCEL user flows
    var result_rowinput = ui.prompt("Select Model Row Input", "Specify the the row input cell", ui.ButtonSet.OK_CANCEL);
    var result_colinput = ui.prompt("Select Model Column Input", "Specify the column input cell", ui.ButtonSet.OK_CANCEL);
    var output2d = dt_range.getCell(1,1);
    var rowinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_rowinput.getResponseText());
    var colinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_colinput.getResponseText());
    config = { "range": dt_range.getA1Notation(), "output": output2d.getA1Notation(), "rowinput": rowinput.getA1Notation(), "colinput": colinput.getA1Notation() };
  } else {
    // column inputs only
    var result_input = ui.prompt("Select Model Input", "Specify the (column) input cell", ui.ButtonSet.OK_CANCEL);
    var input = SpreadsheetApp.getActiveSpreadsheet().getRange(result_input.getResponseText());
    var output = dt_range.getCell(1,2);
    config = { "range": dt_range.getA1Notation(), "output": output.getA1Notation(), "rowinput": null, "colinput": input.getA1Notation() };
  }
  
  // actually do the work now:
  datatables_(config);

  // save named range and property to be able to refresh data
  var name = "DataTable_" + dt_range.getA1Notation().replace(/[^A-Z0-9]/g,"");
  SpreadsheetApp.getActive().setNamedRange(name, dt_range);
  dt_[name] = config;
  PropertiesService.getDocumentProperties().setProperty("dt_", JSON.stringify(dt_));  
}

function refresh_() {
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty("dt_"));
  var ranges = SpreadsheetApp.getActive().getNamedRanges();
  for (var i = 0; i < ranges.length; i++) {
    var name = ranges[i].getName();
    SpreadsheetApp.getUi().alert("Refreshing " + name);
    if (dt_[name]) {
      // book-keeping for cleanup
      dt_[name].exists = true;
      // re-evaluate the configured datatable
      datatables_(dt_[name]);
    }
  }
  
  // cleanup data tables if named range was deleted
  var keys = Object.keys(dt_);
  for (var i = 0 ; i < keys.length; i++) {
    if (!dt_[keys[i]].exists) {
      delete dt_[name];
    }
  }
  PropertiesService.getDocumentProperties().setProperty("dt_", JSON.stringify(dt_));
}

function datatables_(config) {
  var s = SpreadsheetApp.getActive();
  var dt_range = s.getRange(config.range);
  if (!config.rowinput) {
    var input = s.getRange(config.colinput);
    var original = input.getValue();
    var output = s.getRange(config.output);
    for (var i = 2; i <= dt_range.getNumRows(); i++) { 
      input.setValue(dt_range.getCell(i, 1).getValue());
      dt_range.getCell(i, 2).setValue(output.getValue());
    }
    input.setValue(original);
  } else {
    // 2D
    var colinput = s.getRange(config.colinput);
    var rowinput = s.getRange(config.rowinput);
    var colOriginal = colinput.getValue();
    var rowOriginal = rowinput.getValue();
    var output = s.getRange(config.output);
    for (var i = 2; i <= dt_range.getNumRows(); i++) { 
      for (var j = 2; j <= dt_range.getNumColumns(); j++) {
        colinput.setValue(dt_range.getCell(i, 1).getValue());
        rowinput.setValue(dt_range.getCell(1, j).getValue());
        dt_range.getCell(i, j).setValue(output.getValue());
      }
    }
    colinput.setValue(colOriginal);
    rowinput.setValue(rowOriginal);
  }
}
