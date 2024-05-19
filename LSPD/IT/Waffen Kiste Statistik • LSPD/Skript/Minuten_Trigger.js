function Minuten_Trigger()
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);

  if(Minute == 15 || Minute == 30 || Minute == 45 || Minute == 0)
  {
    var Sheet_Statistik = SpreadsheetApp.getActive().getSheetByName("Statistik");
    var Sheet_Dienstblatt = SpreadsheetApp.openById(LSPD.ID_Dienstblatt).getSheetByName("Startseite");

    var Array_WK_DB = Sheet_Dienstblatt.getRange("U10:V12").getValues();
    
    var Letzte_Zeile = Sheet_Statistik.getLastRow() + 1;

    Sheet_Statistik.getRange("B" + Letzte_Zeile + ":L" + Letzte_Zeile).setValues([[Array_WK_DB[0][0],Array_WK_DB[0][1], new Date(),"", Array_WK_DB[1][0],Array_WK_DB[1][1], new Date(),"", Array_WK_DB[2][0],Array_WK_DB[2][1], new Date()]])
  }
}
