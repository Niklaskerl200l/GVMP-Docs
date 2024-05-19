function FG_Uebersicht(e)
{
  var Sheet_Uebersicht = SpreadsheetApp.getActive().getSheetByName("Fraktionsgespräche");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 4 && Zeile <= 18 && Value == "TRUE")
  {
    Sheet_Uebersicht.getRange(Zeile, Spalte).setValue(false);

    var Array_Sheets = SpreadsheetApp.getActive().getSheets();
    var Seite = Sheet_Uebersicht.getRange("B" + Zeile).getValue();

    var Gefunden = false;
    for(var i = 0; i < Array_Sheets.length; i++)
    {
      if(Array_Sheets[i].getName().toString().toUpperCase() == ("FG " + Seite.toString().toUpperCase()))
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == true)
    {
      Array_Sheets[i].setActiveSelection("D2");
    }
  }
}