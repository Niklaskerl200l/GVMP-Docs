function Gespraeche(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("J") && Zeile >= 3 && Value == "TRUE")
  {
    var Sheet_Gespraeche = SpreadsheetApp.getActive().getSheetByName("Gespraeche");

    Sheet_Gespraeche.getRange("K" + Zeile + ":L" + Zeile).setValues([[new Date(),LSPD.Umwandeln()]]);
  }
}
