function Besprechung_Start()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Besprechung_Vorlage = SpreadsheetApp.getActive().getSheetByName("Besprechungsprotokoll Vorlage");

  var Datum = SpreadsheetApp.getUi().prompt("Datum der Besprechung? (dd.MM.yyyy)").getResponseText();
  try
  {
    Logger.log("Test");
    Sheet_Besprechung_Vorlage.copyTo(SS).setName("Besprechungsprotokoll " + Datum);
  }
  catch(err)
  {
    Datum = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm");
    Sheet_Besprechung_Vorlage.copyTo(SS).setName("Besprechungsprotokoll " + Datum);
  }

  var Sheet_Besprechung = SpreadsheetApp.getActive().getSheetByName("Besprechungsprotokoll " + Datum);

  SS.setActiveSheet(Sheet_Besprechung);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(5);

  Sheet_Besprechung.getRange("B1").setValue(Datum);
  Sheet_Besprechung.getRange("B2").setValue("Besprechungsprotokoll vom " + Datum);
}

function Besprechung_Archivieren()
{
  if(SpreadsheetApp.getActive().getSheetName() != "Besprechungsprotokoll Vorlage")
  {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var Sheet_Besprechung = SpreadsheetApp.getActiveSheet();
    var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_PRU_Besprechnungen);

    var Datum = Sheet_Besprechung.getRange("B1").getValue();
    var Anzahl = Sheet_Besprechung.getRange("C1").getValue();

    Sheet_Besprechung.copyTo(Sheet_Export).setName(Sheet_Besprechung.getName());

    var Link = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_Besprechung.getName()).getSheetId();

    Sheet_Export.getSheetByName("Übersicht");

    Sheet_Export.getRange("B" + (Sheet_Export.getLastRow() + 1) + ":C" + (Sheet_Export.getLastRow() + 1)).setValues([[Datum,Anzahl]])
    Sheet_Export.getRange("D" + Sheet_Export.getLastRow()).setFormula("=HYPERLINK(\""+Link+"\";\"Link\")");

    SS.deleteSheet(Sheet_Besprechung);
  }
}