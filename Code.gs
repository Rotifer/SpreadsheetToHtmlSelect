// Display the GUI
// Execute this function from the script editor ro run the application.
// Note the call to "setRngName()". This ensures that all newly added
// values are included in the dropdown list when the GUI is displayed
function displayGUI() {
  var ss,
      html;
  setRngName();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  html = HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ss.show(html);
}
// Called by Client-side JavaScript in the index.html.
// Uses the range name argument to extract the values.
// The values are then sorted and returned as an array
// of arrays.
function getValuesForRngName(rngName) {
  var rngValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rngName).getValues();
  return rngValues.sort();
}

//Expand the range defined by the name as rows are added
function setRngName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName('DropdownValues'),
      firstCellAddr = 'A2',
      dataRngRowCount = sh.getDataRange().getLastRow(),
      listRngAddr = (firstCellAddr + ':A' + dataRngRowCount),
      listRng = sh.getRange(listRngAddr);
  ss.setNamedRange('Cities', listRng);
  //Logger.log(ss.getRangeByName('Cities').getValues());
}

// For debugging
function test_setRngName() {
  setRngName();
}
