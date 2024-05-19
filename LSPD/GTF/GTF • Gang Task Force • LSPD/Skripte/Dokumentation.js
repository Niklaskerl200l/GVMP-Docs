function Dokumentation(e)
{
  var Sheet_Dokumentation = SpreadsheetApp.getActive().getSheetByName("Dokumentation");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 8 && Value != "" && Value != undefined && (OldValue == "" || OldValue == undefined))  // Eingabe von Namen
  {
    var Rang     = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
    var HN       = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
    var Tel      = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

    if(Sheet_Dokumentation.getRange("D" + Zeile).getFormula() == "")
    {
      Sheet_Dokumentation.getRange("D" + Zeile).setValue(Rang);
    }
    else if(Sheet_Dokumentation.getRange("E" + Zeile).getFormula() == "")
    {
      Sheet_Dokumentation.getRange("E" + Zeile).setValue(HN);
    }
    else if(Sheet_Dokumentation.getRange("F" + Zeile).getFormula() == "")
    {
      Sheet_Dokumentation.getRange("F" + Zeile).setValue(Tel);
    }

    Sheet_Dokumentation.getRange("I" + Zeile + ":J" + Zeile).setValues([[new Date(),LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 8 && Value != "" && Value != undefined)  // Eingabe von Namen
  {
    Sheet_Dokumentation.getRange("I" + Zeile + ":J" + Zeile).setValues([[new Date(),LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 8 && (Value == "" || Value == undefined))  // Löschen von Name
  {
    var Fraktion = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$C;2;false);))';
    var Rang     = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
    var HN       = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
    var Tel      = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

    Sheet_Dokumentation.getRange("C" + Zeile + ":J" + Zeile).setValues([[Fraktion,Rang,HN,Tel,"","","",""]]);
  }
  else if(Spalte == Spalte_in_Index("K") && Zeile >= 5 && Zeile <= 8 && Value == "TRUE")  // Prüfen der Eingaben
  {
    var Lock = LockService.getScriptLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(e)
    {
      Logger.log('Timeout wegen Lock bei Einsatz Eintragung');
      SpreadsheetApp.getUi().alert("Ein Fehler ist aufgetreten versuche es noch einmal");
      Fehler = true;
      return 0;
    }


    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName('Logs');

    var Array_Eingabe = Sheet_Dokumentation.getRange("B" + Zeile + ":J" + Zeile).getValues();
    var Array_Ausgabe = Eintrag_Check(Array_Eingabe,0,1,7,8,true);

    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue();
    
    if(Array_Ausgabe == 1) return 0;
    
    Sheet_Log.getRange("B" + Letzte_Zeile + ":F" + Letzte_Zeile).setValues([[ Array_Ausgabe[0][0],Array_Ausgabe[0][1], Array_Ausgabe[0][5],new Date(),Array_Ausgabe[0][8] ]]);

    SpreadsheetApp.flush();
    Lock.releaseLock();
  }
}