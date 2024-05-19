function einblenden() {
  var sheetName = 'Import Personaltabelle';

  // Holen des aktiven Spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Holen des Blatts nach Namen
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Blatt einblenden
  if (sheet) {
    sheet.showSheet();
  } else {
    Logger.log('Blatt nicht gefunden: ' + sheetName);
  }
}
