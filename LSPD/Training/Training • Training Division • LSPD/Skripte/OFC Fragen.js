function ofcfragen(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if((Spalte == Spalte_in_Index("C") || Spalte == Spalte_in_Index("H")  || Spalte == Spalte_in_Index("M")) && Value != undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("OFC Theoriefragen").getRange(Zeile,Spalte + 2).setValue(LSPD.Umwandeln());
  }
}
