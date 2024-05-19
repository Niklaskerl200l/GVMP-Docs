function Fahrzeugsperren(e)
{
  var Sheet_Fahrzeugsperren = SpreadsheetApp.getActive().getSheetByName("Parkkrallen (NEU)");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("H") && Zeile == 2 && Value != undefined)
  {
    var Array_Sperren = Sheet_Fahrzeugsperren.getRange("B11:B").getValues();
    var Gefunden = false;

    for(var i = 0; i < Array_Sperren.length; i++)
    {
      if(Array_Sperren[i][0] != "" && Array_Sperren[i][0] == Value)
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == true)
    {
      Sheet_Fahrzeugsperren.setActiveSelection("B" + (i + 11));
    }

    SpreadsheetApp.flush();
    Utilities.sleep(15000);

    Sheet_Fahrzeugsperren.getRange(Zeile, Spalte).clearContent();
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 7)
  {
    if(Value != undefined && e.oldValue == undefined)
    {
      Sheet_Fahrzeugsperren.getRange("D" + Zeile).setValue("Unangemeldet");
      Sheet_Fahrzeugsperren.getRange("F" + Zeile).setValue("Kostenpflichtige Entfernung einer Parkkralle");
    }
    else if(Value == undefined && e.oldValue != undefined)
    {
      Sheet_Fahrzeugsperren.getRange("B" + Zeile + ":F" + Zeile).clearContent();
    }
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 5 && Zeile <= 7 && Value == "TRUE")
  {
    Sheet_Fahrzeugsperren.getRange(Zeile, Spalte).setValue(false);

    var Array_Eintrag = Sheet_Fahrzeugsperren.getRange("B" + Zeile + ":F" + Zeile).getValues();
    Array_Eintrag = Array_Eintrag[0];

    Sheet_Fahrzeugsperren.getRange("B" + Zeile + ":F" + Zeile).clearContent();

    Array_Eintrag[5] = new Date();
    Array_Eintrag[6] = LSPD.Umwandeln();

    var Lock = LockService.getDocumentLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(err)
    {
      throw Error("Fehler!\nLocküberschreitung: Fahrzeugsperren");
    }

    Sheet_Fahrzeugsperren.insertRowAfter(10);
    Sheet_Fahrzeugsperren.getRange("B11:H11").setValues([Array_Eintrag]);

    Log_Zaehler("Parkkralle\nAngebracht", Array_Eintrag[0]);

    SpreadsheetApp.flush();
    Lock.releaseLock();
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 11 && Value == "TRUE")
  {
    Sheet_Fahrzeugsperren.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.flush();

    var UI = SpreadsheetApp.getUi();
    var Confirmation = UI.alert("Parkkralle entfernen...", "Hallo " + LSPD.Umwandeln() + ",\nmöchten Sie diese Parkkralle lösen?", UI.ButtonSet.YES_NO);

    if(Confirmation == UI.Button.YES)
    {
      var Array_Eintrag = Sheet_Fahrzeugsperren.getRange("B" + Zeile + ":H" + Zeile).getValues();
      Array_Eintrag = Array_Eintrag[0];

      var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Log Kralle");
      Sheet_Archiv.insertRowAfter(3);
      Sheet_Archiv.getRange("B4:J4").setValues([[
        "",
        Array_Eintrag[0],
        Array_Eintrag[1],
        Array_Eintrag[3],
        Array_Eintrag[4],
        Array_Eintrag[6],
        Array_Eintrag[5],
        LSPD.Umwandeln(),
        new Date()
      ]]);

      Log_Zaehler("Parkkralle\nEntfernt", Array_Eintrag[0]);
      Sheet_Fahrzeugsperren.deleteRow(Zeile);
    }
  }
}