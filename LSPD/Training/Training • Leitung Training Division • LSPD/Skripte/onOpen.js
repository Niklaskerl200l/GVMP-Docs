function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  var user = Session.getTemporaryActiveUserKey();
  
  Logger.log("Benutzer: " + user);

  ui.createMenu('Funktionen')
    .addItem('Personal Sortieren', 'Personal_Sortieren')
    .addToUi();

  LSPD.onOpen();
}