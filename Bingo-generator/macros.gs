function SaveBingo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  let name = spreadsheet.getRange('Generator!C2').getValue();

  let values = spreadsheet.getRange('PreviousBingoWords!C2:2').getValues()[0];
  values = values.filter(n => n);
  let freeColumn = 3 + values.length;
  for (let index in values) {
    if (values[index] == name) {
      freeColumn = parseInt(index) + 3;
    }
  }


  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  let sheet = spreadsheet.getActiveSheet();
  sheet.getRange(6,freeColumn).activate();
  spreadsheet.getRange('Generator!A2:A50').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet.getRange(3,freeColumn).activate();
  spreadsheet.getRange('Generator!B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet.getRange(2,freeColumn).activate();
  spreadsheet.getRange('Generator!C2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet.getRange(5,freeColumn).activate();
  spreadsheet.getRange('Generator!D4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet.getRange(4,freeColumn).activate();
  spreadsheet.getRange('Generator!B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
};

function UsePreviousBingo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
  spreadsheet.getRange('A2').activate();
  spreadsheet.getRange('PreviousBingoWords!A6:A100').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('C30').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
  spreadsheet.getRange('D4').activate();
  spreadsheet.getRange('PreviousBingoWords!A5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
  spreadsheet.getRange('B4').activate();
  spreadsheet.getRange('PreviousBingoWords!A4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getRange('PreviousBingoWords!A3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PreviousBingoWords'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Generator'), true);
  spreadsheet.getRange('C2').activate();
  spreadsheet.getRange('PreviousBingoWords!A2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};
