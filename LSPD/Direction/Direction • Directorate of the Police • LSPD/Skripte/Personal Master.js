function Personal_Master(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("P") && Zeile >= 6 && Zeile <= 150 && Value == "TRUE")
  {
    Entlassung(e);
  }
  else if(Spalte == Spalte_in_Index("Q") && Zeile >= 6 && Zeile <= 150 && Value == "TRUE")
  {
    Rankup(e);
  }
  else if(Spalte == Spalte_in_Index("R") && Zeile >= 6 && Zeile <= 150 && Value == "TRUE")
  {
    Rankdown(e);
  }
  else if(Spalte == Spalte_in_Index("O") && Zeile == 2  && Value != "" && Value != undefined && Value != null)
  {
    Selection_Master_Name(SpreadsheetApp.getActive().getSheetByName("Personal Master").getRange("O2").getValue());
  }
}