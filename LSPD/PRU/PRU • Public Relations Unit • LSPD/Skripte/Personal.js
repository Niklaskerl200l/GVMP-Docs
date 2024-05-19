function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte >= 2 && Spalte <= 3 && Zeile >= 5 && Zeile <= 20 && Value != undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  else if(Spalte >= 3 && Zeile >= 5 && Zeile <= 20 && Value == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");

  }
}
