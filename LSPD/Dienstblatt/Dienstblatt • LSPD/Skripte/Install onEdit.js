function Install_onEdit(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(SheetName == "Startseite")
  {
    if((Spalte == Spalte_in_Index("C") || Spalte == Spalte_in_Index("E") || Spalte == Spalte_in_Index("I")) && Zeile == 12 && Value != undefined)
    {
      var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
      Sheet_Startseite.getRange(Zeile, Spalte).clearContent();

      var SS_GTF = SpreadsheetApp.openById(LSPD.ID_GTF);
      var Sheet_GTF_Datenbank = SS_GTF.getSheetByName("Import Aktuell");
      var Array_GTF_Datenbank = Sheet_GTF_Datenbank.getRange("B3:I").getValues();

      var Gefunden = false;
      for(var i = 0; i < Array_GTF_Datenbank.length; i++)
      {
        if(Array_GTF_Datenbank[i][0].toString().toUpperCase().replace(" ", "_").includes(Value.toString().toUpperCase().replace(" ", "_")) == true && Array_GTF_Datenbank[i][0] != "")
        {
          Gefunden = true;
          break;
        }
      }

      SpreadsheetApp.flush();
      var UI = SpreadsheetApp.getUi();

      if(Gefunden == true)
      {
        UI.alert("GTF Suche...", "Name: " + Array_GTF_Datenbank[i][0] + "\nZugehörigkeit: " + Array_GTF_Datenbank[i][1] + "\nLetzte Aktivität: " + Utilities.formatDate(Array_GTF_Datenbank[i][7], "CET", "dd.MM.yyyy HH:mm").toString(), UI.ButtonSet.OK);
      }
      else
      {
        UI.alert("GTF Suche...", "Es wurde nichts gefunden...", UI.ButtonSet.OK);
      }
    }
    else if(Spalte == Spalte_in_Index("C") && Zeile == 12 && Value == undefined)
    {
      var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
      
      Sheet_Startseite.getRange("E12:I12").setValues([["",,,,""]]);
    }
    else if(Spalte == Spalte_in_Index("I") && Zeile >= 83 && Zeile <= 87 && Value == "TRUE")
    {
      var Lock = LockService.getDocumentLock();
      try
      {
        Lock.waitLock(28000);
      }
      catch(err)
      {
        throw Error("Zeitüberschreitung!");
      }

      var SS_Detective = SpreadsheetApp.openById(LSPD.ID_Detective);
      var Sheet_Rueckrufe = SS_Detective.getSheetByName("Rückrufanfragen");

      var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

      var Array_Insert = 
      [[
        new Date(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("C")).getValue(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("E")).getValue(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("G")).getValue()
      ]];

      Sheet_Rueckrufe.insertRowAfter(4);
      Sheet_Rueckrufe.getRange("B5:E5").setValues(Array_Insert);

      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("C")).clearContent();
      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("E")).clearContent();
      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("G")).clearContent();

      Lock.releaseLock();
    }
    else if(Spalte == Spalte_in_Index("I") && Zeile >= 90 && Zeile <= 94 && Value == "TRUE")
    {
      var Lock = LockService.getDocumentLock();
      try
      {
        Lock.waitLock(28000);
      }
      catch(err)
      {
        throw Error("Zeitüberschreitung!");
      }

      var SS_GTF = SpreadsheetApp.openById(LSPD.ID_GTF);
      var Sheet_Rueckrufe = SS_GTF.getSheetByName("Rückrufanfragen");

      var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

      var Array_Insert = 
      [[
        new Date(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("C")).getValue(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("E")).getValue(),
        Sheet_Startseite.getRange(Zeile, Spalte_in_Index("G")).getValue()
      ]];

      Sheet_Rueckrufe.insertRowAfter(4);
      Sheet_Rueckrufe.getRange("B5:E5").setValues(Array_Insert);

      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("C")).clearContent();
      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("E")).clearContent();
      Sheet_Startseite.getRange(Zeile, Spalte_in_Index("G")).clearContent();

      Lock.releaseLock();
    }
    else if(Spalte == Spalte_in_Index("I") && Zeile >= 109 && Zeile <= 119 && Value == "TRUE")
    {
      var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);

      var Array_Eintrag = Sheet_Startseite.getRange("C" + Zeile + ":H" + Zeile).getValues();
      Array_Eintrag = Array_Eintrag[0];

      var SS_Wartung = SpreadsheetApp.openById(LSPD.ID_Fahrzeugwartung);
      var Sheet_Schaeden = SS_Wartung.getSheetByName("Schadensliste");

      Sheet_Schaeden.appendRow(["", Array_Eintrag[0], Array_Eintrag[2], new Date(), Array_Eintrag[5], "", ""]);

      Sheet_Startseite.getRange("C" + Zeile + ":H" + Zeile).clearContent();
    }
  }
  else if(SheetName == "Wartung")
  {
    if(Spalte == Spalte_in_Index("Q") && Zeile == 5 && Value == "TRUE")  // Sortieren Fahrzeug Liste
    {
      SpreadsheetApp.getActive().getSheetByName("Wartung").getRange(Zeile,Spalte).setValue("");
      Wartung_Sortieren();
    }
  }

  LSPD.onEdit(e);
}
