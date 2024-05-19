function Rankup(e)
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

  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Master = SpreadsheetApp.getActive().getSheetByName("Personal Master");

  var Rang = Sheet_Master.getRange("D" + Zeile).getValue();
  var Name = Sheet_Master.getRange("C" + Zeile).getValue();

  Sheet_Master.getRange("Q" + Zeile).setValue(false);
  
  if(Rang >= 0 && Rang <= 11 && Value == "TRUE")
  {
    Sheet_Master.getRange("D" + Zeile).setValue(Rang + 1);
    Sheet_Master.getRange("J" + Zeile).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"));

    Sheet_Master.getRange("B6:O" + Sheet_Master.getLastRow()).sort([{column: 4, ascending: false}, {column: 3, ascending: true}]);

    Selection_Master_Name(Name);
  }
  else
  {
    SpreadsheetApp.getUi().alert("Noch ein paar Wünsche?\n\nEs funktioniert nur ein Rang 13 und der bin ich!");
  }

  SpreadsheetApp.flush();
  Lock.releaseLock();
}
