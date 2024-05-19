function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  var user = Session.getTemporaryActiveUserKey();
  
  Logger.log("Benutzer: " + user);

  ui.createMenu('Funktionen')
    //.addItem('Dokumentation Namensänderung', 'Namen_Ersetzen')
    .addItem('Besprechungsprotokoll Erstellen', 'Besprechung_Start')
    //.addItem('Dokumentation Archivieren', 'Archivieren')
    .addToUi();

    LSPD.onOpen();

    Bewerber_onOpen();

  var Sheet_Rueckrufe = SpreadsheetApp.getActive().getSheetByName("Rückrufanfragen");
  var Array_Rueckrufe = Sheet_Rueckrufe.getRange("C4:G").getValues();

  var Count = 0;

  for(var i = 0; i < Array_Rueckrufe.length; i++)
  {
    if(Array_Rueckrufe[i][0] != "")
    {
      if(Array_Rueckrufe[i][3] == false && Array_Rueckrufe[i][4] == false)
      {
        Count++;
      }
    }
  }

  if(Count > 0)
  {
    SpreadsheetApp.getUi().alert("Rückrufe", "Es sind noch " + Count + " Rückrufe offen!", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}