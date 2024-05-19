function Blitzermarathon3(e) 
{
  var Sheet_Blitzermarathon = SpreadsheetApp.getActive().getSheetByName("Blitzermarathon");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("L") && Zeile >= 23 && Zeile <= 25 && Value == "TRUE")
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

    var Array_Eintragen = Sheet_Blitzermarathon.getRange("B" + Zeile + ":L" + Zeile).getValues();

    Sheet_Blitzermarathon.getRange("B" + Zeile + ":L" + Zeile).setValue("");

    Sheet_Blitzermarathon.insertRowAfter(28);
    Sheet_Blitzermarathon.getRange("E29:F29").merge();
    Sheet_Blitzermarathon.getRange("I29:J29").merge();

    Sheet_Blitzermarathon.getRange("B29:L29").setValues(Array_Eintragen);

    Sheet_Blitzermarathon.setActiveSelection("B28")

    SpreadsheetApp.flush();
    Lock.releaseLock();
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 23 && Zeile <= 25 && Value != undefined)
  {
    var Array_Streifen = Sheet_Blitzermarathon.getRange("Y5:Z" + Sheet_Blitzermarathon.getLastRow()).getValues();
    var Array_Messstellen = Sheet_Blitzermarathon.getRange("AE4:AF37").getValues();

    var Name = Benutzername();

    var Streife;
    var KMH;

    for(var i = 0; i < Array_Streifen.length; i++)
    {
      if(Name == Array_Streifen[i][0])
      {
        Streife = Array_Streifen[i][1];
        break;
      }
    }

    for(var i = 0; i < Array_Messstellen.length; i++)
    {
      if(Array_Messstellen[i][0] == Streife)
      {
        KMH = Array_Messstellen[i][1];
        break;
      }
    }

    Sheet_Blitzermarathon.getRange("C" + Zeile + ":K" + Zeile).setValues([[Streife,KMH,"","","","",Name,"",Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm")]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 23 && Zeile <= 25 && Value == undefined)
  {
    Sheet_Blitzermarathon.getRange("B" + Zeile + ":K" + Zeile).setValue("");
  }
}