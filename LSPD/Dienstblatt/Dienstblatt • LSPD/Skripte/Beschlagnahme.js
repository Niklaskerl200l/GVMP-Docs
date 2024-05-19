// Rework by Cockstas
var Beschlagnahme_Lock = false;

function Beschlagnahme(e)
{
  var Sheet_Beschlagnahme = SpreadsheetApp.getActive().getSheetByName("Beschlagnahme");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 9 && Value == undefined && OldValue != undefined)
  {
    Sheet_Beschlagnahme.getRange("B" + Zeile + ":H" + Zeile).clearContent();
    Sheet_Beschlagnahme.getRange("I" + Zeile + ":J" + Zeile).setValue(false).insertCheckboxes();

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setBackground(null);
    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setFontColor(null);

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setHorizontalAlignment("center");
    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setVerticalAlignment("middle");
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 5 && Zeile <= 9 && Value == "TRUE")
  {
    Sheet_Beschlagnahme.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Beschlagnahme wird dokumentiert...", "Beschlagnahme...", 30);

    var Lock = LockService.getDocumentLock();
    try
    {
      Beschlagnahme_Lock = true;
      Lock.waitLock(28000);
    }
    catch(err)
    {
      Beschlagnahme_Lock = false;
      SpreadsheetApp.getActive().toast("Fehler!\nVersuchen Sie es bitte erneut!", "Beschlagnahme...");
      throw Error("Beschlagnahme: Locküberschreitung...");
    }

    var Array_Beschlagnahme = Sheet_Beschlagnahme.getRange("B" + Zeile + ":I" + Zeile).getValues();
    Array_Beschlagnahme = Array_Beschlagnahme[0];

    if(Array_Beschlagnahme[0] == "")
    {
      Lock.releaseLock();
      return SpreadsheetApp.getActive().toast("Fehler!\nBitte geben Sie den Namen des TV an!", "Beschlagnahme...", 30);
    }

    Log_Zaehler("Beschlagnahmung\nDokumentiert", Array_Beschlagnahme[0] + "\n" + Array_Beschlagnahme[3] + "\n" + Array_Beschlagnahme[4] + "\n" +  Array_Beschlagnahme[5]);

    Sheet_Beschlagnahme.insertRowAfter(12);
    Sheet_Beschlagnahme.getRange("B13:K13").setValues(
      [[
        Array_Beschlagnahme[0],
        "",
        Array_Beschlagnahme[2],
        Array_Beschlagnahme[3],
        Array_Beschlagnahme[4],
        Array_Beschlagnahme[5],
        Array_Beschlagnahme[6],
        Array_Beschlagnahme[7],
        new Date(),
        LSPD.Umwandeln()
      ]]
    );

    Sheet_Beschlagnahme.getRange("I13").insertCheckboxes();

    Lock.releaseLock();
    Beschlagnahme_Lock = false;
    SpreadsheetApp.getActive().toast("Beschlagnahmung abgeschlossen...", "Beschlagnahmne...");

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":H" + Zeile).clearContent();
    Sheet_Beschlagnahme.getRange("I" + Zeile + ":J" + Zeile).setValue(false).insertCheckboxes();

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setBackground(null);
    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setFontColor(null);

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setHorizontalAlignment("center");
    Sheet_Beschlagnahme.getRange("B" + Zeile + ":K" + Zeile).setVerticalAlignment("middle");

    Sheet_Beschlagnahme.getRange("B" + Zeile + ":C" + Zeile).merge();
  }
}

function Beschlagnahme_Zugehoerigkeit()
{
  Logger.log("Beschlagnahme: Starte Suche nach Zugehörigkeiten...");

  if(Beschlagnahme_Lock == true)
  {
    return Logger.log("Beschlagnahme: Suche abgebrochen... Lock aktiv!");
  }

  var Sheet_Beschlagnahme = SpreadsheetApp.getActive().getSheetByName("Beschlagnahme");
  var Array_Beschlagnahme = Sheet_Beschlagnahme.getRange("B13:C").getValues();

  if(Array_Beschlagnahme.filter(function(e){return e[0] != "" && e[1] == ""}).length > 0)
  {
    var SS_GTF = SpreadsheetApp.openById(LSPD.ID_GTF);
    var Sheet_GTF_Datenbank = SS_GTF.getSheetByName("Import Aktuell");
    var Array_GTF_Datenbank = Sheet_GTF_Datenbank.getRange("B3:C").getValues();

    var Gefunden = false;
    for(var i = 0; i < Array_Beschlagnahme.length; i++)
    {
      if(Array_Beschlagnahme[i][0] != "" && Array_Beschlagnahme[i][1] == "")
      {
        Gefunden = false;
        for(var o = 0; o < Array_GTF_Datenbank.length; o++)
        {
          if(Array_GTF_Datenbank[o][0].toString().toUpperCase() == Array_Beschlagnahme[i][0].toString().toUpperCase())
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == true)
        {
          Sheet_Beschlagnahme.getRange("C" + (i + 13)).setValue(Array_GTF_Datenbank[o][1]);
          Logger.log("\t" + Array_Beschlagnahme[i][0] + " ist ein Mitglied der/des " + Array_GTF_Datenbank[o][1]);
        }
        else
        {
          Sheet_Beschlagnahme.getRange("C" + (i + 13)).setValue("Zivilist");
        }
      }
    }
  }

  Logger.log("Beschlagnahme: Suche abgeschlossen...");
}