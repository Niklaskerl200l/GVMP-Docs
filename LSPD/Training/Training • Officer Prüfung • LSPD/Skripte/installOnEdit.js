function installOnEdit(e)
{ 
  var SheetName = e.source.getActiveSheet().getName();
  if(SheetName.includes("Officer Prüfung "))
  {
    var Sheet_Pruefung = SpreadsheetApp.getActive().getSheetByName(SheetName);
    var Spalte = e.range.getColumn();
    var Zeile = e.range.getRow();
    var Value = e.value;

    if(Spalte == Spalte_in_Index("B") && Zeile >= 113 && Zeile <= 114 && Value == "TRUE")
    {
      Sheet_Pruefung.getRange(Zeile, Spalte).setValue(false);

      var UI = SpreadsheetApp.getUi();

      var Confirmation = UI.alert(SheetName, "Möchten Sie diese Prüfung abbrechen? Sie wird anschließend sofort ins Archiv verschoben.", UI.ButtonSet.YES_NO);
      if(Confirmation == UI.Button.YES)
      {
        var Confirmation_Detective = UI.alert("Detective Bureau", "Möchten Sie dass dieser Prüfungabbruch an das Detective Bureau als Beschwerde gemeldet wird?", UI.ButtonSet.YES_NO);
        if(Confirmation_Detective == UI.Button.YES)
        {
          var SS_Detective = SpreadsheetApp.openById(LSPD.ID_Detective);
          var Sheet_Beschwerden_Neu = SS_Detective.getSheetByName("Beschwerden Neu");
          var Zeile_Beschwerden_Neu = Sheet_Beschwerden_Neu.getRange("B1").getValue();

          var Sheet_Auswertungsgedoens = SS_Detective.getSheetByName("Auswertungsgedöns");
          var Fallnummer = Sheet_Auswertungsgedoens.getRange("C3").getValue();

          var PBO = Sheet_Pruefung.getRange("F4").getValue();

          var Array_Insert = 
          [[
            Fallnummer,
            PBO,
            "",
            "Offen",
            new Date(),
            "Training Division",
            "LSPD",
            new Date(),
            "Prüfungsabbruch",
            "Officer Prüfung von " + PBO + " wurde aus folgendem Grund abgebrochen: " + Sheet_Pruefung.getRange(Zeile, Spalte + 1).getValue() 
          ]];

          Sheet_Beschwerden_Neu.getRange(Zeile_Beschwerden_Neu, Spalte_in_Index("B"), 1, Array_Insert[0].length).setValues(Array_Insert);
          Sheet_Beschwerden_Neu.getRange("Q" + Zeile_Beschwerden_Neu).setValue("LSPD IT");
        }

        Archivieren();
      }
    }
  }
}