function Install_onEdit(e) {
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  Logger.log("Sheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);

  try
  {
    if (SheetName == "Forenverwaltung NEU" && Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("D") && Value == "TRUE") {
      var Sheet_Forum = SpreadsheetApp.getActive().getSheetByName("Forenverwaltung NEU");
      var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Recruitment).getSheetByName("Bewerbungen Neu");
      var Sheet_Export_Bewerber = SpreadsheetApp.openById(LSPD.ID_Recruitment).getSheetByName("Bewerbungen");

      var Bewerber = Sheet_Forum.getRange("C" + Zeile).getValue();
      var Bearbeiter = Sheet_Forum.getRange("I" + Zeile).getValue();

      var Letzte_Zeile = Sheet_Export.getRange("B3").getValue();
      var Letzte_Zeile_Bewerber = Sheet_Export_Bewerber.getRange("B1").getValue() + 1;

      var Array_Neu = Sheet_Export.getRange("B5:B" + Letzte_Zeile).getValues();

      var Gefunden = false;

      for (var y = 0; y < Array_Neu.length; y++) {
        if (Array_Neu[y][0] == Bewerber) {
          Gefunden = true;
          break;
        }
      }

      if (Gefunden == false) {
        Logger.log(Bearbeiter);

        Sheet_Forum.getRange("I" + Zeile).setValue("");

        Sheet_Export.getRange("B" + Letzte_Zeile + ":C" + Letzte_Zeile).setValues([[Bewerber, "Evaluierung"]]);
        Sheet_Export.getRange("K" + Letzte_Zeile).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM") + " " + Bearbeiter);

        Sheet_Export_Bewerber.getRange("B" + Letzte_Zeile_Bewerber + ":C" + Letzte_Zeile_Bewerber).setValues([[Bewerber, true]]);
        Sheet_Export_Bewerber.getRange("BH" + Letzte_Zeile_Bewerber).setValue(true);

        Sheet_Export.getRange("B5:M35").sort(2);
      }
    }
    else if (SheetName == "Forenverwaltung NEU" && Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("E") && Value == "TRUE") {
      var Sheet_Forum = SpreadsheetApp.getActive().getSheetByName("Forenverwaltung NEU");
      var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Recruitment).getSheetByName("Bewerbungen Neu");

      var Bewerber = Sheet_Forum.getRange("C" + Zeile).getValue();
      var Bearbeiter = Sheet_Forum.getRange("I" + Zeile).getValue();

      var Letzte_Zeile = Sheet_Export.getRange("B3").getValue();

      var Array_Bewerber = Sheet_Export.getRange("B5:B" + Letzte_Zeile).getValues();

      Sheet_Forum.getRange("I" + Zeile).setValue("");

      for (var y = 0; y < Array_Bewerber.length; y++) {
        if (Array_Bewerber[y][0] == Bewerber) {
          Sheet_Export.getRange("C" + (y + 5)).setValue("Eingeladen");
          Sheet_Export.getRange("L" + (y + 5)).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM") + " " + Bearbeiter);
          break;
        }
      }
    }
    else if (SheetName == "Forenverwaltung NEU" && Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("F") && Value == "TRUE") {
      var Sheet_Bewerber = SpreadsheetApp.getActive().getSheetByName("Import Bewerber");
      var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Dienstblatt).getSheetByName("Startseite");
      var Sheet_Forum = SpreadsheetApp.getActive().getSheetByName("Forenverwaltung NEU");

      var Letzte_Zeile = Sheet_Export.getRange("X55").getValue();

      var Array_Bewerber = Sheet_Bewerber.getRange("D2:G" + Sheet_Bewerber.getRange("D1").getValue()).getValues();
      var Array_RideAlongs = Sheet_Export.getRange("S57:S" + Letzte_Zeile).getValues();

      for (var i = 0; i < Array_Bewerber.length; i++) {
        var Gefunden = false;

        for (var x = 0; x < Array_RideAlongs.length; x++) {
          if (Array_Bewerber[i][0] == Array_RideAlongs[x][0]) {
            Gefunden = true;
            break;
          }
        }
        Logger.log(Gefunden);
        if (Gefunden == false) {
          Logger.log("Ridealong add " + Array_Bewerber[i][0]);
          Letzte_Zeile++;
          // Sheet_Export.getRange("S" + Letzte_Zeile + ":W" + Letzte_Zeile).setValues([[Array_Bewerber[i][0],"","",Array_Bewerber[i][3],Array_Bewerber[i][1]]]);
        }
      }

      Sheet_Forum.getRange("C" + Zeile).setValue("");
      Sheet_Forum.getRange("F" + Zeile).setValue("");
    }

    else if (SheetName == "Export Academy / EST" && Zeile >= 5 && Zeile <= 11 && Spalte == Spalte_in_Index("C") && Value == "TRUE")    // Academy Hacken setzen bei Training
    {
      var Sheet_Training = SpreadsheetApp.openById(LSPD.ID_Training).getSheetByName("Ausbildungsblatt");
      var Sheet_Hacken = SpreadsheetApp.getActive().getSheetByName(SheetName);

      try {
        var Array_Name = Sheet_Training.getRange("B4:B").getValues();

        var RCT = Sheet_Hacken.getRange(Zeile, Spalte - 1).getValue();
        var Gefunden = false;

        for (var y = 0; y < Array_Name.length; y++) {
          if (Array_Name[y][0] == RCT) {
            var Formel = Sheet_Training.getRange("E" + (y + 4)).getFormula();

            var ID = Formel.substring(Formel.indexOf("\"") + 1, Formel.indexOf("\"", Formel.indexOf("\"") + 1));

            var Sheet_Ausbildungen = SpreadsheetApp.openById(ID).getSheetByName("Prüfungen & Praxis");

            Sheet_Ausbildungen.getRange("H9:L9").setValues([[new Date(), , "Recruitment Division", , true]]);

            Sheet_Hacken.getRange("B" + Zeile + ":C" + Zeile).setValue("");

            Gefunden = true;

            break;
          }
        }
        if (Gefunden == false) {
          Person_Nicht_Gefunden;
        }
      }
      catch (err) {
        Logger.log(err.stack);

        Sheet_Hacken.getRange(Zeile, Spalte - 1).setValue(err);

        Fehler;
      }
    }


    else if (SheetName == "Export Academy / EST" && Zeile >= 5 && Zeile <= 11 && Spalte == Spalte_in_Index("F") && Value == "TRUE")    // EST Hacken setzen bei Training
    {
      var Sheet_Training = SpreadsheetApp.openById(LSPD.ID_Training).getSheetByName("Ausbildungsblatt");
      var Sheet_Hacken = SpreadsheetApp.getActive().getSheetByName(SheetName);

      try {
        var Array_Name = Sheet_Training.getRange("B4:B").getValues();

        var RCT = Sheet_Hacken.getRange(Zeile, Spalte - 1).getValue();
        var Gefunden = false;

        for (var y = 0; y < Array_Name.length; y++) {
          if (Array_Name[y][0] == RCT) {
            var Formel = Sheet_Training.getRange("E" + (y + 4)).getFormula();

            var ID = Formel.substring(Formel.indexOf("\"") + 1, Formel.indexOf("\"", Formel.indexOf("\"") + 1));

            var Sheet_Ausbildungen = SpreadsheetApp.openById(ID).getSheetByName("Prüfungen & Praxis");

            Sheet_Ausbildungen.getRange("H11:L11").setValues([[new Date(), , "Recruitment Division", , true]]);

            Sheet_Hacken.getRange("E" + Zeile + ":F" + Zeile).setValue("");

            Gefunden = true;

            break;
          }
        }

        if (Gefunden == false) {
          Person_Nicht_Gefunden;
        }
      }
      catch (err) {
        Logger.log(err.stack);

        Sheet_Hacken.getRange(Zeile, Spalte - 1).setValue(err);

        Fehler;
      }
    }
  }
  catch(err)
  {
    throw Error("Fehler!\n" + err.stack);
  }
}