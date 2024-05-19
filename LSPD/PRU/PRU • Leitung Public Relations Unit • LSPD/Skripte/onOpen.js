function onOpen() 
{
  SpreadsheetApp.getUi().createMenu('Funktionen')
    .addItem('Sortieren', 'Sortieren')
    .addToUi();

  LSPD.onOpen();
}