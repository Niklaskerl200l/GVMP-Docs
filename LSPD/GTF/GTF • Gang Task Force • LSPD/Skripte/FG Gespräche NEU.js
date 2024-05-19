function FG_Neu(e)
{
  var Sheet_Gespraech = SpreadsheetApp.getActive().getSheetByName(e.source.getActiveSheet().getName());
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("W") && Zeile == 3 && Value == "TRUE") // Gespräch ausblenden...
  {
    Sheet_Gespraech.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.flush();

    Sheet_Gespraech.hideSheet();
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 9 && Zeile <= 18) // Anliegen ein- & austragen...
  {
    if(Value != undefined)
    {
      Sheet_Gespraech.getRange("S" + Zeile).setValue(new Date());
      Sheet_Gespraech.getRange("U" + Zeile).setValue(LSPD.Umwandeln());
    }
    else if(Value == undefined)
    {
      Sheet_Gespraech.getRange("S" + Zeile + ":W" + Zeile).clearContent();
    }
  }
  else if(Spalte == Spalte_in_Index("Y") && Zeile >= 9 && Zeile <= 18 && Value == "TRUE") // Anliegen archivieren...
  {
    Sheet_Gespraech.getRange(Zeile, Spalte).setValue(false);

    var Array_Anliegen = Sheet_Gespraech.getRange("B" + Zeile + ":W" + Zeile).getValues();
    Array_Anliegen = Array_Anliegen[0];

    Sheet_Gespraech.getRange("B" + Zeile + ":W" + Zeile).clearContent();

    Sheet_Gespraech.insertRowAfter(36);

    Sheet_Gespraech.getRange("B37:C37").merge();
    Sheet_Gespraech.getRange("D37:K37").merge();
    Sheet_Gespraech.getRange("L37:S37").merge();
    Sheet_Gespraech.getRange("T37:V37").merge();
    Sheet_Gespraech.getRange("W37:Y37").merge();

    Sheet_Gespraech.getRange("B37:W37").setValues([[new Date(), "", Array_Anliegen[0], "", "", "", "", "", "", "", Array_Anliegen[8], "", "", "", "", "", "", "", "Anliegen", "", "", Array_Anliegen[19]]]);
  }
  else if(Spalte == Spalte_in_Index("X") && Zeile == 21 && Value == "TRUE") // Gesprächsverlauf archivieren...
  {
    Sheet_Gespraech.getRange(Zeile, Spalte).setValue(false);

    var Gespraech = Sheet_Gespraech.getRange("B" + Zeile).getValue();

    Sheet_Gespraech.getRange("B" + Zeile).clearContent();

    Sheet_Gespraech.insertRowAfter(36);

    Sheet_Gespraech.getRange("B37:C37").merge();
    Sheet_Gespraech.getRange("D37:K37").merge();
    Sheet_Gespraech.getRange("L37:S37").merge();
    Sheet_Gespraech.getRange("T37:V37").merge();
    Sheet_Gespraech.getRange("W37:Y37").merge();

    Sheet_Gespraech.getRange("B37:W37").setValues([[new Date(), "", Gespraech, "", "", "", "", "", "", "", "-", "", "", "", "", "", "", "", "Gespräch", "", "", LSPD.Umwandeln()]]);
  }
}