function Beschwerden_Neu(e)
{
  var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 36)
  {
    if(Value != undefined && OldValue == undefined) // Neuer Fall erstellt...
    {
      var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
      var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();

      Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":F" + Zeile).setValues([[Fallnummer, Value, "", "Offen", new Date()]]);
      Sheet_Beschwerden_Neu.getRange("Q" + Zeile).setValue(LSPD.Umwandeln());
    }
    else if(Value == undefined && OldValue != undefined) // Fall löschen...
    {
      var UI = SpreadsheetApp.getUi();
      var Confirmation = UI.alert("Detective Bureau", "Möchten Sie diesen Fall wirklich löschen?", UI.ButtonSet.YES_NO);

      if(Confirmation == UI.Button.YES)
      {
        Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":Q" + Zeile).clearContent();
      }
    }
  }
  else if(Spalte == Spalte_in_Index("R") && Zeile >= 3 && Zeile <= 36 && Value == "TRUE") // Beschwerde in Bearbeitung...
  {
    Sheet_Beschwerden_Neu.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();

    SpreadsheetApp.getActive().toast("Beschwerde wird übertragen...", "Bitte warten...");

    var Array_Beschwerde_Neu = Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":Q" + Zeile).getValues();
    Array_Beschwerde_Neu = Array_Beschwerde_Neu[0];

    var Array_Beschwerden_In_Bearbeitung =
    [[
      Array_Beschwerde_Neu[0],
      "Team",
      Array_Beschwerde_Neu[1],
      Array_Beschwerde_Neu[2],
      LSPD.Umwandeln(),
      "In Bearbeitung",
      "In Bearbeitung",
      "",
      "",
      "",
      "",
      Array_Beschwerde_Neu[4],
      Array_Beschwerde_Neu[5],
      Array_Beschwerde_Neu[6],
      Array_Beschwerde_Neu[7],
      Array_Beschwerde_Neu[8],
      Array_Beschwerde_Neu[9],
      "",
      Array_Beschwerde_Neu[10],
      Array_Beschwerde_Neu[11],
      Array_Beschwerde_Neu[12],
      Array_Beschwerde_Neu[13],
      Array_Beschwerde_Neu[14],
      "",
      "",
      "",
      Array_Beschwerde_Neu[15],
      false,
      new Date()
    ]];

    var Lock = LockService.getDocumentLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(err)
    {
      throw Error("Fehler!\nLocküberschreitung: Beschwerdenübertragung");
    }

    var Sheet_Beschwerden_In_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden in Bearbeitung");
    var Zeile_Beschwerden_In_Bearbeitung = Sheet_Beschwerden_In_Bearbeitung.getRange("B1").getValue();

    Sheet_Beschwerden_In_Bearbeitung.getRange("B" + Zeile_Beschwerden_In_Bearbeitung + ":AD" + Zeile_Beschwerden_In_Bearbeitung).setValues(Array_Beschwerden_In_Bearbeitung);
    Sheet_Beschwerden_In_Bearbeitung.getRange("AC" + Zeile_Beschwerden_In_Bearbeitung).setValue(false).insertCheckboxes();

    Sheet_Beschwerden_In_Bearbeitung.setActiveSelection("C" + Zeile_Beschwerden_In_Bearbeitung);

    Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":Q" + Zeile).clearContent();

    SpreadsheetApp.getActive().toast("Warte immernoch! Da kommt noch was...");

    Sort_Beschwerden();
    Lock.releaseLock();

    SpreadsheetApp.getActive().toast("Jetzt kannst du anfangen...");
  }
}