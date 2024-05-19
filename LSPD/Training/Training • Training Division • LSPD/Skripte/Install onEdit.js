function Install_onEdit(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var SheetName = e.source.getActiveSheet().getName();

  if(SheetName == "Entziehen")
  {
    if(Spalte == Spalte_in_Index("I") && Zeile >= 5 && Value == "TRUE")
    {
      var Array_Eintag = SpreadsheetApp.getActive().getSheetByName("Entziehen").getRange("B" + Zeile + ":I" + Zeile).getValues();
      
      Array_Eintag[0][7] = "Offen";

      SpreadsheetApp.openById(LSPD.ID_Leitung_Training).getSheetByName("Entziehen").insertRowAfter(5).getRange("B6:I6").setValues(Array_Eintag);
      SpreadsheetApp.getActive().getSheetByName("Entziehen").getRange("B" + Zeile + ":I" + Zeile).setValue("");
    }
  }
}