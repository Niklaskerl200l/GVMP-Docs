function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 6 && Zeile <= 21 && Spalte == Spalte_in_Index("C"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  if(Zeile >= 6 && Zeile <= 21 && Spalte == Spalte_in_Index("G"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
}

function Sortieren()
{
  var Sortier_Bereich = "A4:G18"

  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");

  Sheet_Personal.getRange(Sortier_Bereich).sort([{column: 1, ascending: false}]);
}