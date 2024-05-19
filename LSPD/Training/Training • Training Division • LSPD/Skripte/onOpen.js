function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

  ui.createMenu('Funktionen')
    .addItem('Besprechungsprotokoll Erstellen', 'Besprechung_Start')
    .addItem('Ausbildungsblatt Sortieren', 'Ausbildungsblatt_Sortieren')
    .addToUi();

  LSPD.onOpen();

  Bewerber_onOpen();


  if(Sheet_Startseite.getRange("C27").getValue() != "Keine neue Prüfungstermine")
  {
    SpreadsheetApp.getUi().alert("Es sind noch offene Prüfungstermine vorhanden.");
  }
}