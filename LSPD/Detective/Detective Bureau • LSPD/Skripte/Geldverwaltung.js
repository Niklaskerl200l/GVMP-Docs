function Geldverwaltung(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Geldverwaltung = SpreadsheetApp.getActive().getSheetByName("Geldverwaltung");

  if(Spalte == Spalte_in_Index("E") && Zeile >= 3 && Value != undefined)
  {
    Sheet_Geldverwaltung.getRange("F" + Zeile).setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 3 && Value == "TRUE")
  {
    Sheet_Geldverwaltung.getRange("H" + Zeile).setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 3 && Value == "TRUE")
  {
    Sheet_Geldverwaltung.getRange("J" + Zeile).setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("E") && Zeile >= 3 && Value == undefined)
  {
    Sheet_Geldverwaltung.getRange("F" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 3 && Value == "FALSE")
  {
    Sheet_Geldverwaltung.getRange("H" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 3 && Value == "FALSE")
  {
    Sheet_Geldverwaltung.getRange("J" + Zeile).setValue("");
  }

}
