function onOpen() 
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Funktionen')
    .addItem('Besprechungsprotokoll Erstellen', 'Besprechung_Start')
    .addToUi();

  LSPD.onOpen();
}