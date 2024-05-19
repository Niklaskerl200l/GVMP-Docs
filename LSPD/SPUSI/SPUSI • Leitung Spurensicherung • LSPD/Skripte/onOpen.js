function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Funktionen')
    .addItem('Personal Sortieren', 'Personal_Sortieren')
    .addToUi();

  LSPD.onOpen();
}