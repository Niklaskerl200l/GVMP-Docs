function Offene_Kontrollen(e)
{
  var Sheet_Offene_Kontrollen = SpreadsheetApp.getActive().getSheetByName("Offene Kontrollen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("E") && Zeile >= 5 && Zeile <= 252 && Value == "TRUE")
  {
    Sheet_Offene_Kontrollen.getRange(Zeile, Spalte).uncheck();

    var Fahrzeug = Sheet_Offene_Kontrollen.getRange("B" + Zeile).getValue();

    var Sheet_Kontrolle = SpreadsheetApp.getActive().getSheetByName("Kontrolle");
    var Zeile_Kontrolle = Sheet_Kontrolle.getRange("B1").getValue();

    var Sheet_Fuhrpark = SpreadsheetApp.getActive().getSheetByName("Fuhrpark");
    var Array_Fuhrpark = Sheet_Fuhrpark.getRange("B4:C253").getValues();

    var Gefunden = false;
    var Modell;

    for(var i = 0; i < Array_Fuhrpark.length; i++)
    {
      if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0].toString() == Fahrzeug)
      {
        Gefunden = true;
        Modell = Array_Fuhrpark[i][1];

        break;
      }
    }

    if(Gefunden == true)
    {
      Array_Fuhrpark = Sheet_Fuhrpark.getRange("I4:M253").getValues();
      Gefunden = false;

      for(var i = 0; i < Array_Fuhrpark.length; i++)
      {
        if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0] == Modell)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == true)
      {
        Sheet_Kontrolle.getRange("B" + Zeile_Kontrolle).setValue(Fahrzeug);
        Sheet_Kontrolle.getRange("E" + Zeile_Kontrolle).setNote(`mind. ${Array_Fuhrpark[i][2].toString()} Verbandskästen`);
        Sheet_Kontrolle.getRange("F" + Zeile_Kontrolle).setNote(`mind. ${Array_Fuhrpark[i][3].toString()} Reparaturkästen`);
        Sheet_Kontrolle.getRange("G" + Zeile_Kontrolle).setNote(`mind. ${Math.floor(Number(Array_Fuhrpark[i][4]) * 0.75).toString()}l`);

        Sheet_Kontrolle.setActiveSelection("B" + Zeile_Kontrolle);
      }
      else
      {
        return SpreadsheetApp.getUi().alert("Fehler!", "Dieses Fahrzeug ist nicht in den Kontrolledaten hinterlegt!", SpreadsheetApp.getUi().ButtonSet.OK);
      }
    }
  }
}