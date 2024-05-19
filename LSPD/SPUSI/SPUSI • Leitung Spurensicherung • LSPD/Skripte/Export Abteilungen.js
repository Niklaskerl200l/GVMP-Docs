function Abteilungen(e)
{
  var Zeile = e.range.getRow();
  
  if(Zeile >= 12)   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
}
