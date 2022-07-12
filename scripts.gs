function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Utils")
    .addItem('Paste time', 'time')
    .addItem("Fill form", "main")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

var sheet = SpreadsheetApp.getActive().getSheetByName("<your_data_table>");

var PRINT_OPTIONS = {
  'size': 7,               // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  'fzr': false,            // repeat row headers
  'portrait': true,        // false=landscape
  'fitw': true,            // fit window or actual size
  'gridlines': false,      // show gridlines
  'printtitle': false,
  'sheetnames': false,
  'pagenum': 'UNDEFINED',  // CENTER = show page numbers / UNDEFINED = do not show
  'attachment': false,
  'horizontal_alignment': 'CENTER',
  'vertical_alignment': 'CENTER',
}

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function printForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var gid = sheet.getSheetId();
  var url = ss.getUrl().replace(/edit/, '') + 'export?format=pdf' + PDF_OPTS + "&gid=" + gid;

  var htmlTemplate = HtmlService.createTemplateFromFile('js');
  htmlTemplate.url = url;
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setHeight(10).setWidth(100), 'Done');
}

function objectToQueryString(obj) { // Adding arguments to the URL
  return Object.keys(obj).map(function(key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
}

function getInput() { 
  var orderRow = Browser.inputBox('Fill', 'Please, enter the row to use:', Browser.Buttons.OK_CANCEL);
  if (orderRow == "cancel") { return; }
  var rowNumber = Number(orderRow)
  if (isNaN(rowNumber) || rowNumber < 3 || rowNumber > sheet.getLastRow()){
    Browser.msgBox("Error", Utilities.formatString('Row "%s" is not valid.', orderRow), Browser.Buttons.OK);
    return;
  }
  return rowNumber;
}

function getData(rowNumber) { // Getting data from row
  var columns = sheet.getRange(2, 1, 1, sheet.getMaxColumns()).getValues()[0]; // Put your own table configuration
  var data = sheet.getDataRange().getValues()[rowNumber-1];
  return data;
}

function fillInOrder(dataArray){ // Function to paste data to form
  var formSheet = SpreadsheetApp.getActive().getSheetByName("<your_form>");
  var formCells = // Cells that needed to be fill
    ["D4", "C36:D36", "B7:F7", "C9:F9", "B10:F10", "C12:F12", 
    "B14:C14", "E14:F14", "B16:C16", "E16:F16", "C18:F18", "B22:F22", 
    "B26:F26", "B29:F29"];
  for (var i = 0; i < formCells.length; i++){ formSheet.getRange(formCells[i]).setValue(dataArray[i]) }
}

function main(){ // Main function
  var rowNumber = getInput();
  var dataArray = getData(rowNumber);
  fillInOrder(dataArray);
  printForm();
}

function onEdit() { // Function that paste time automatically
  var currentSheet = SpreadsheetApp.getActiveSheet();
  if( currentSheet.getName() == "<your_data_table>" ) {
    var actCell = currentSheet.getActiveCell();
    if( actCell.getColumn() == 1) {
      var nextCell = actCell.offset(0, 1); // offset 0 by vertical and 1 by horizontal
      if( nextCell.getValue() === '') nextCell.setValue(new Date());
    }
  }
}
