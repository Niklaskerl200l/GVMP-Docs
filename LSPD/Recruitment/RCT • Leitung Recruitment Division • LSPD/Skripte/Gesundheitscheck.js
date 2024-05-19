function Gesundheitscheck(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile == 5 && Spalte == Spalte_in_Index("B") && Value != undefined)
  {
    var Sheet_Gesundheitscheck = SpreadsheetApp.getActive().getSheetByName("Gesundheitscheck");

    Sheet_Gesundheitscheck.getRange("C5:D5").setValues([[new Date(),LSPD.Umwandeln()]]);
  }

  if(Zeile == 5 && Spalte == Spalte_in_Index("E") && Value == "TRUE")
  {
    var Sheet_Gesundheitscheck = SpreadsheetApp.getActive().getSheetByName("Gesundheitscheck");

    var Array = Sheet_Gesundheitscheck.getRange("B5:E5").getValues();

    Sheet_Gesundheitscheck.getRange("B5:E5").setValue("");
    Sheet_Gesundheitscheck.insertRowAfter(8);
    Sheet_Gesundheitscheck.getRange("B9:E9").setValues(Array);
    //Sheet_Gesundheitscheck.getRange("B" + (Sheet_Gesundheitscheck.getLastRow()+1) + ":E" + (Sheet_Gesundheitscheck.getLastRow()+1)).setValues(Array);

  }
}
