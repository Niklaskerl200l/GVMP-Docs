function Minuten_Trigger()
{
  var Sheet_Bewerber = SpreadsheetApp.getActive().getSheetByName("Import Bewerber");
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Dienstblatt).getSheetByName("Startseite");

  var Array_Dienstblatt = Sheet_Export.getRange("S61:S105").getValues();
  var Array_Bewerber = Sheet_Bewerber.getRange("C2:G47").getValues();

  for(var i = 0; i < Array_Bewerber.length && i < 45; i++)
  {
    if(Array_Bewerber[i][0] != "TRUE")
    {
      if(Array_Bewerber[i][1] != Array_Dienstblatt[i][0])
      {
        Sheet_Export.getRange("S" + (i+61) + ":W" + (i+61)).setValues([[Array_Bewerber[i][1],"","",Array_Bewerber[i][4],Array_Bewerber[i][2]]]);
      }
    }
  }

  //------------------------------------------------------------------------------------------------------//


  var Sheet_Gesundheitscheck = SpreadsheetApp.openById(LSPD.ID_Gesundheitscheck).getSheetByName("Kontaktliste");
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");

  var Array_Personal = Sheet_Personal.getRange("C6:G19").getValues();
  var Array_GSC = Sheet_Gesundheitscheck.getRange("B6:D19").getValues();
  var Array_Neu = new Array();

  for(var i = 0; i < Array_Personal.length; i++)
  {
    Array_Neu.push([Array_Personal[i][0],Array_Personal[i][1],Array_Personal[i][4]]);
  }

  if(Array_Neu.toString() != Array_GSC.toString())
  {
    Logger.log("Update GSC");
    Sheet_Gesundheitscheck.getRange("B6:D19").setValues(Array_Neu);
  }
}
