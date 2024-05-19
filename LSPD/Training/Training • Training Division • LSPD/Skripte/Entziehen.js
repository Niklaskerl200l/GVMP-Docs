function Entziehen(e)
{
  var Sheet_Entziehen = SpreadsheetApp.getActive().getSheetByName("Entziehen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Value != undefined && Value != "")
  {
    Sheet_Entziehen.getRange("G" + Zeile + ":H" + Zeile).setValues([[LSPD.Umwandeln(),new Date()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && (Value == undefined && Value == ""))
  {
    Sheet_Entziehen.getRange("G" + Zeile + ":H" + Zeile).setValue("");
  }
}
