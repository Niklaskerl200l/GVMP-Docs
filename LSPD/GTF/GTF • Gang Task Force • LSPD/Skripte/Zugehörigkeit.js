function Zugehoerigkeit(e)
{
  var Sheet_Zugehoerigkeit = SpreadsheetApp.getActive().getSheetByName("Zugehörigkeit");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("J") && Zeile >= 4 && Value != "" && Value != undefined)
  {
    Sheet_Zugehoerigkeit.getRange("K" + Zeile + ":L" + Zeile).setValues([[new Date(), LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 4 && (Value == "" || Value == undefined))
  {
    Sheet_Zugehoerigkeit.getRange("K" + Zeile + ":L" + Zeile).setValue("");
  }

  if(Spalte == Spalte_in_Index("J") && Zeile >= 4 && Value == "Angenommen")
  {
    var Sheet_Dokumentation = SpreadsheetApp.getActive().getSheetByName("Dokumentation");

    var Array_Zugehoerigkeit = Sheet_Zugehoerigkeit.getRange("C" + Zeile + ":F" + Zeile).getValues();

    Sheet_Dokumentation.getRange("B8:C8").setValues([[Array_Zugehoerigkeit[0][0],Array_Zugehoerigkeit[0][3]]]);
    Sheet_Dokumentation.getRange("F8").setValue(Array_Zugehoerigkeit[0][1]);
    Sheet_Dokumentation.getRange("G8:H8").setValue("");
    Sheet_Dokumentation.getRange("I8").setValue(new Date());
    Sheet_Dokumentation.getRange("J8").setValue(LSPD.Umwandeln());

    Sheet_Dokumentation.setActiveSelection("K8");
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 3 && Value == "Watchlist")
  {
    var Sheet_Dokumentation = SpreadsheetApp.getActive().getSheetByName("Dokumentation");

    var Array_Zugehoerigkeit = Sheet_Zugehoerigkeit.getRange("C" + Zeile + ":F" + Zeile).getValues();

    Sheet_Dokumentation.getRange("B8:C8").setValues([[Array_Zugehoerigkeit[0][0],"Überprüfung"]]);
    Sheet_Dokumentation.getRange("F8").setValue(Array_Zugehoerigkeit[0][1]);
    Sheet_Dokumentation.getRange("G8:H8").setValue("Watchlist");
    Sheet_Dokumentation.getRange("I8").setValue(new Date());
    Sheet_Dokumentation.getRange("J8").setValue(LSPD.Umwandeln());

    Sheet_Dokumentation.setActiveSelection("K8");
  }
}
