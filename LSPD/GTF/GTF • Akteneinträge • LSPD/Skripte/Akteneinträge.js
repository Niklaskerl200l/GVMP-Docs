function Akteneintrag(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("K") && Zeile >= 7 && Value == "TRUE")
  {
    var Lock = LockService.getScriptLock();
    
    try
    {
      Lock.waitLock(28000);
    }
    catch(e)
    {
      Logger.log('Timeout wegen Lock bei Einsatz Eintragung');
      SpreadsheetApp.getUi().alert("Ein Fehler ist aufgetreten versuche es noch einmal");
      Fehler = true;
      return 0;
    }

    var Sheet_Akteneintrag = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Logs");

    var Name = Sheet_Akteneintrag.getRange("B" + Zeile).getValue();
    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue()

    Sheet_Log.getRange("B" + Letzte_Zeile + ":D" + Letzte_Zeile).setValues([[Name,new Date(), LSPD.Umwandeln()]]);

    SpreadsheetApp.flush();
    Lock.releaseLock();
  }
}
