//in arbeit Fabio
function Einsatz_Dokumentation(e) 
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Einsatz");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;
  var Datum =new Date();

  
  if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value == undefined)
  {
    Sheet.getRange("D" + Zeile + ":H" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("H") && Zeile == 5 && Value == true)
  {

      var Array_Einsatz = Sheet_Einsatz.getRange("B5:F5").getValues();
    var Array_Letzter_Einsatz = Sheet_Einsatz.getRange("B9:H9").getValues();

    var Cooldown = new Date();

    Cooldown.setMinutes(Cooldown.getMinutes() - 30);

    if(Array_Letzter_Einsatz[0][0] == Array_Einsatz[0][0] && Array_Letzter_Einsatz[0][1] == Array_Einsatz[0][1] && Array_Letzter_Einsatz[0][5] >= Cooldown)
    {
      Sheet_Einsatz.getRange("B5:H5").setValue("");
      SpreadsheetApp.getUi().alert("Dieser Einsatz wurde schon Eingetragen");
    }
    else
    {
      Sheet_Einsatz.insertRowAfter(8);

      Sheet_Einsatz.getRange("B9:I9").setValues([[Array_Einsatz[0][0], Array_Einsatz[0][1],Array_Einsatz[0][2],Array_Einsatz[0][3],Array_Einsatz[0][4], new Date(), LSPD.Umwandeln(), Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MM")]]);

      Sheet_Einsatz.getRange("B5:H5").setValue("");
    }
  }








}
