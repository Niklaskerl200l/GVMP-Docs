function Bewerber(e) 
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte >= 7 && Zeile >= 7 && Zeile <= 19 && Value != undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("Bewerber").getRange("K" + Zeile).setValue(LSPD.Umwandeln());
    SpreadsheetApp.getActive().getSheetByName("Bewerber").getRange("L" + Zeile).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"));
  }
  else if(Spalte >= 7 && Zeile >= 7 && Zeile <= 19 && Value == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("Bewerber").getRange("K" + Zeile).setValue("");
    SpreadsheetApp.getActive().getSheetByName("Bewerber").getRange("L" + Zeile).setValue("");
  }
}