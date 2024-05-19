function Archivieren()
{
    var SS_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Officer_Prüfung);
    var SS_Officer = SpreadsheetApp.getActiveSpreadsheet();
    var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Officer_Prüfung).getSheetByName("Übersicht");
    var Sheet_Officer = SpreadsheetApp.getActiveSheet();

    if(Sheet_Officer.getName() == "Prüfung Vorlage")
    {
      SpreadsheetApp.getUi().alert("HALT STOP!\nHier bleibt alles so wie's ist!")
      return 0;
    }

    var Anzhal = Sheet_Export.getRange("B1").getValue();
    var Name = Sheet_Officer.getRange("F4").getValue();
    var Datum = Sheet_Officer.getRange("F5").getValue();
    var Versuche = Sheet_Officer.getRange("I5").getValue();
    var Pruefer1T = Sheet_Officer.getRange("P4").getValue();
    var Pruefer2T = Sheet_Officer.getRange("P5").getValue();
    var Pruefer1P = Sheet_Officer.getRange("P44").getValue();
    var Pruefer2P = Sheet_Officer.getRange("P45").getValue();
    var Ergebniss = Sheet_Officer.getRange("Q114").getValue();

    var Sheet_Import_Termine = SpreadsheetApp.getActive().getSheetByName("Import Termine");
    var Array_Termine = Sheet_Import_Termine.getRange("C2:C").getValues();

    try
    {
      for(var i = Array_Termine.length - 1; i => 0; i--)
      {
        if(Array_Termine[i][0] == Name)
        {
          Sheet_Import_Termine.getRange("N" + (i + 2)).setValue(true);
        }
      }
    }
    catch(err)
    {
      Logger.log("Person hat keinen Termin gemacht");
      Logger.log(err.stack);
    }

    Sheet_Officer.copyTo(SS_Export).setName(Sheet_Officer.getName() + " Archiv");

    var Sheet_Export_Archiv = SS_Export.getSheetByName(Sheet_Officer.getName() + " Archiv");

    Sheet_Export_Archiv.getRange("A1:R112").setValues(Sheet_Officer.getRange("A1:R112").getValues());

    SS_Officer.deleteSheet(Sheet_Officer);


    var URL = SS_Export.getUrl() + "#gid=" + Sheet_Export_Archiv.getSheetId();

    Sheet_Export.getRange("B" + Anzhal + ":I" + Anzhal).setValues([[Name,Datum,Versuche,Pruefer1T,Pruefer2T,Pruefer1P,Pruefer2P,Ergebniss]]);

    Sheet_Export.getRange("J" + Anzhal).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");
  
}
