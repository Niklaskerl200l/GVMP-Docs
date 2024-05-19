function Meldungsblatt(e)  //LG hier war ich auch
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 6 && Value != undefined)    // Datum, Beamter eintragen
  {
    SpreadsheetApp.getActive().getSheetByName("Meldungsblatt").getRange("D" + Zeile + ":E" + Zeile).setValues([[new Date(),LSPD.Umwandeln()]]);
  }

  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 6 && Value == undefined)   // Datume, Beamter leeren
  {
    SpreadsheetApp.getActive().getSheetByName("Meldungsblatt").getRange("D" + Zeile + ":E" + Zeile).setValue("");
  }

  else if(Spalte == Spalte_in_Index("F") && Zeile >= 5 && Zeile <= 6 && Value == "TRUE")    // Zeile Übertragen
  {
    var Sheet_Meldungsblatt = SpreadsheetApp.getActive().getSheetByName("Meldungsblatt");

    var Array = Sheet_Meldungsblatt.getRange("B"+ Zeile + ":E" + Zeile).getValues();

    Sheet_Meldungsblatt.getRange("B" + Zeile + ":F" + Zeile).setValue("");

    Sheet_Meldungsblatt.insertRowAfter(9);

    Sheet_Meldungsblatt.getRange("C10:D10").merge();

    Sheet_Meldungsblatt.getRange("B10:F10").setValues([[Array[0][0],Array[0][1],"",Array[0][2],Array[0][3]]]);
  }
}
