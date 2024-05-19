function Stunden_Trigger()
{
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Log Stempeluhr");
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Stempeluhr).getSheetByName("Archiv Stempeluhr");

  var Letzt_Zeile = Sheet_Archiv.getLastRow();

  if(Letzt_Zeile < 3)
  {
    Letzt_Zeile = 3;
  }

  var Array_Archiv = Sheet_Archiv.getRange("B3:D" + Letzt_Zeile).getValues();

  Sheet_Export.getRange("B" + (Sheet_Export.getLastRow() + 1) + ":D" + (Sheet_Export.getLastRow() + Array_Archiv.length)).setValues(Array_Archiv);

  Sheet_Archiv.getRange("B3:D" + Letzt_Zeile).setValue("");
}

function Archivieren_Leitstelle()
{
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Log Leitstelle");
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Leitstelle).getSheetByName("Archiv Leitstelle");

  var Letzt_Zeile = Sheet_Archiv.getLastRow();

  if(Letzt_Zeile < 3)
  {
    Letzt_Zeile = 3;
  }

  var Array_Archiv = Sheet_Archiv.getRange("B3:D" + Letzt_Zeile).getValues();

  Sheet_Export.getRange("B" + (Sheet_Export.getLastRow() + 1) + ":D" + (Sheet_Export.getLastRow() + Array_Archiv.length)).setValues(Array_Archiv);

  Sheet_Archiv.getRange("B3:D" + Letzt_Zeile).setValue("");
}
