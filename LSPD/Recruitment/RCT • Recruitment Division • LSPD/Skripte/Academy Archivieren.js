function Academy_Archivieren()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Academy = SpreadsheetApp.getActiveSheet();
  var Sheet_Export = SpreadsheetApp.openById("1v7WSWCY0qAa6PMvYtH85XdIatfAgHH9pp3_Uo-m4xJk");

  var Datum = Sheet_Academy.getRange("G3").getValue();
  var Anzahl_Personen = Sheet_Academy.getRange("F3").getValue() - 5;
  var Zeile_Export = Sheet_Export.getSheetByName("Übersicht").getRange("B1").getValue();
  
  Sheet_Academy.copyTo(Sheet_Export).setName(Sheet_Academy.getName());
  
  var URL = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_Academy.getName()).getSheetId();

  Sheet_Export = Sheet_Export.getSheetByName("Übersicht");

  Sheet_Export.getRange("B" + Zeile_Export + ":C" + Zeile_Export).setValues([[Datum,Anzahl_Personen]]);
  Sheet_Export.getRange("D" + Zeile_Export).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  SS.deleteSheet(Sheet_Academy);
}
