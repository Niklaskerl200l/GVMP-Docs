function Minuten_Trigger()
{
  var SS_Import = SpreadsheetApp.openById(LSPD.ID_Dienstblatt);

  var Sheet_FIB = SpreadsheetApp.getActive().getSheetByName("Export FIB");
  var Sheet_Startseite = SS_Import.getSheetByName("Startseite");
  var Sheet_Auswertung = SS_Import.getSheetByName("Auswertungsgedöns");
  //var Sheet_Doku = SS_Import.getSheetByName("Beschlagnahmungen");
  var Sheet_Einsatz = SS_Import.getSheetByName("Einsatz");

  //var Array_Doku_LSPD = Sheet_Doku.getRange("B13:N28").getValues();
  var Array_Doku_FIB = Sheet_FIB.getRange("B9:N9").getValues();

  var Array_Akte_LSPD = Sheet_Einsatz.getRange("K5:K9").getValues();
  var Array_Akte_FIB = Sheet_FIB.getRange("I3:L7").getValues();

  if(Sheet_Startseite.getRange("N20").getValue() != Sheet_FIB.getRange("D3").getValue())
  {
    Logger.log("Aktuallisiere PD auf Streife");
    Sheet_FIB.getRange("D3").setValue(Sheet_Startseite.getRange("N20").getValue());
  }

  if(Sheet_Startseite.getRange("R20").getValue() != Sheet_FIB.getRange("E3").getValue())
  {
    Logger.log("Aktuallisiere PD im Büro");
    Sheet_FIB.getRange("E3").setValue(Sheet_Startseite.getRange("R20").getValue());
  }

  if(Sheet_Startseite.getRange("E6").getValue() != Sheet_FIB.getRange("G3").getValue())
  {
    Logger.log("Aktuallisiere PD Funk");
    Sheet_FIB.getRange("G3").setValue(Sheet_Startseite.getRange("E6").getValue());
  }

  if(Sheet_Auswertung.getRange("E5").getValue() != Sheet_FIB.getRange("B3").getValue())
  {
    Logger.log("Aktuallisiere PD Einsatz");
    Sheet_FIB.getRange("B3").setValue(Sheet_Auswertung.getRange("E5").getValue());
  }

  /*
  if(Array_Doku_LSPD[0].toString() != Array_Doku_FIB[0].toString())
  {
    Logger.log("Aktuallisiere PD Doku");

    var Gefunden = false;

    for(var y = 0; y < Array_Doku_LSPD.length; y++)
    {
      if(Array_Doku_LSPD[y].toString() == Array_Doku_FIB[0].toString())
      {
        Sheet_FIB.insertRowsBefore(9,y);
        var Array_Ausgabe = new Array();

        for(var y2 = 0; y2 < y; y2++)
        {
          Array_Ausgabe.push(Array_Doku_LSPD[y2]);
        }

        Sheet_FIB.getRange(9,2,y,13).setValues(Array_Ausgabe);

        Sheet_FIB.deleteRows(Sheet_FIB.getLastRow() - y + 1,y);

        Gefunden = true;
      }
    }

    if(Gefunden == false)
    {
      Logger.log("Aktuallisiere PD Doku Kommplett");

      var Array_Doku_LSPD = Sheet_Doku.getRange(13,2,Sheet_Doku.getLastRow() - 12,13).getValues();

      Sheet_FIB.getRange(9,2,Sheet_FIB.getLastRow() - 8, 13).setValue("");

      Sheet_FIB.getRange(9,2,Array_Doku_LSPD.length,Array_Doku_LSPD[0].length).setValues(Array_Doku_LSPD);
    }
  }
  */

  if(Array_Akte_LSPD[0].toString() != Array_Akte_FIB[0].toString())
  {
    Logger.log("Sammelakte");
    Sheet_FIB.getRange("I3:I7").setValues(Array_Akte_LSPD);
  }

  if(Sheet_Einsatz.getRange("B5").getValue() != Sheet_FIB.getRange("B6").getValue())
  {
    Logger.log("Aktuallisiere Einsatz");
    Sheet_FIB.getRange("B6").setValue(Sheet_Einsatz.getRange("B5").getValue());
  }

  if(Sheet_Einsatz.getRange("C5").getValue() != Sheet_FIB.getRange("C6").getValue())
  {
    Logger.log("Aktuallisiere Einsatz Funk");
    Sheet_FIB.getRange("C6").setValue(Sheet_Einsatz.getRange("C5").getValue());
  }

  if(Sheet_Einsatz.getRange("B8").getValue() != Sheet_FIB.getRange("D6").getValue())
  {
    Logger.log("Aktuallisiere Einsatz TV Krankenhaus");
    Sheet_FIB.getRange("D6").setValue(Sheet_Einsatz.getRange("B8").getValue());
  }

  if(Sheet_Einsatz.getRange("E8").getValue() != Sheet_FIB.getRange("E6").getValue())
  {
    Logger.log("Aktuallisiere Einsatz Abtransport durch");
    Sheet_FIB.getRange("E6").setValue(Sheet_Einsatz.getRange("E8").getValue());
  }

  if(Sheet_Einsatz.getRange("G8").getValue() != Sheet_FIB.getRange("F6").getValue())
  {
    Logger.log("Aktuallisiere Einsatz LSMC");
    Sheet_FIB.getRange("F6").setValue(Sheet_Einsatz.getRange("G8").getValue());
  }
}
