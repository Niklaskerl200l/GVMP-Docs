function EST_Verlaengerung(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Verlaengerung = SpreadsheetApp.getActive().getSheetByName("EST Verlängerung")

  if(Spalte == Spalte_in_Index("B") && Zeile == 3 && Value != "" && Value != undefined)
  {
    Sheet_Verlaengerung.getRange("E3:F3").setValues([[LSPD.Umwandeln(),new Date()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile == 3 && (Value == "" || Value == undefined))
  {
    Sheet_Verlaengerung.getRange("E3:F3").setValue("");
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile == 3 && Value == "TRUE")
  {
    var Array_Eingabe = Sheet_Verlaengerung.getRange("B3:G3").getValues();

    Sheet_Verlaengerung.getRange("B3:G3").setValue("");
    
    Sheet_Verlaengerung.insertRowAfter(6);

    Sheet_Verlaengerung.getRange("B7:G7").setValues(Array_Eingabe);
  }
}
