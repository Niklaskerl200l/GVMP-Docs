function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  var user = Session.getTemporaryActiveUserKey();
  
  Logger.log("Benutzer: " + user);

  ui.createMenu('Funktionen')
    .addItem('Besprechungsprotokoll Erstellen', 'Besprechung_Start')
    .addToUi();

  LSPD.onOpen();
  Bewerber_onOpen();
  Bewerber_onOpen_Abteilung();
  
}