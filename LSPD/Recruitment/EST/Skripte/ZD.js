function ZD(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile == 6 && Spalte == Spalte_in_Index("J") && Value == "TRUE")
  {
    var Sheet_ZD = SpreadsheetApp.getActive().getSheetByName("ZD");

    var Zeile_Letzte = Sheet_ZD.getRange("L7").getValue();
    var Array_ZD = Sheet_ZD.getRange("B6:I6").getValues();
    
    Sheet_ZD.getRange("B6:J6").setValue("");
    Sheet_ZD.getRange("L" + Zeile_Letzte + ":P" + Zeile_Letzte).setValues([[Array_ZD[0][0],Array_ZD[0][2],Array_ZD[0][4],Array_ZD[0][5],Array_ZD[0][7]]]);
  }
  else if(Zeile == 6 && Spalte == Spalte_in_Index("B") && Value != undefined)
  {
    var Sheet_ZD = SpreadsheetApp.getActive().getSheetByName("ZD");

    Sheet_ZD.getRange("D6:F6").setValues([[Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"),,LSPD.Umwandeln()]]);
  }
  else if(Zeile == 6 && Spalte == Spalte_in_Index("B") && Value == undefined)
  {
    var Sheet_ZD = SpreadsheetApp.getActive().getSheetByName("ZD");

    Sheet_ZD.getRange("D6:F6").setValue("");
  }
}
