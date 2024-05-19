function Rueckrufe(e)
{
  var Sheet_Rueckrufe = SpreadsheetApp.getActive().getSheetByName("Rückrufanfragen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 3 && Value == "TRUE")
  {
    Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
    Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln());
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 3 && Value == "TRUE")
  {
    Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
    Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM HH:mm"));
  }
}