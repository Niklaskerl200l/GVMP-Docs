function ProjektStart()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var UI = SpreadsheetApp.getUi();

  var Sheet_Vorlage = SS.getSheetByName("[Vorlage] Projektplanung");
  
  var Abfrage1 = UI.prompt("Gib bitte den Namen des Projekt an.",UI.ButtonSet.OK_CANCEL)

      var Projekt = Abfrage1.getResponseText();

      var Sheet_Planung = Sheet_Vorlage.copyTo(SS).setName("Planung-" + Projekt).showSheet();
      
      Sheet_Planung.getRange("B2").setValue("" + Projekt);

      Sheet_Planung.setActiveSelection("A1");

} 

//--------------------------//

function ProjektArchivieren() {
  if (SpreadsheetApp.getActive().getSheetName() !== "[Vorlage] Projektplanung]") {
    var S = SpreadsheetApp.getActiveSpreadsheet();
    var Sheet_ProjektPlanung = SpreadsheetApp.getActiveSheet();
    var S_Export = SpreadsheetApp.openById("1fDUX7180AWcDqgA0kqfAzYEK15Yyxrw4tNoWCQg-TnM");

    var Datum = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy");

    var Sheet_Archiv = Sheet_ProjektPlanung.copyTo(S_Export);

    try {
      Sheet_Archiv.setName(Sheet_ProjektPlanung.getName());
    } catch (err) {
      try {
        Sheet_Archiv.setName(Sheet_ProjektPlanung.getName() + " 2");
      } catch (err) {
        SpreadsheetApp.getUi().alert("Trollst du?!");
        S_Export.deleteSheet(Sheet_Archiv);
        return;
      }
    }

    var Link = S_Export.getUrl() + "#gid=" + Sheet_Archiv.getSheetId();

    var rangeToClearValidations = Sheet_Archiv.getRange(1, 1, Sheet_Archiv.getMaxRows(), Sheet_Archiv.getMaxColumns());
    var dataValidations = rangeToClearValidations.getDataValidations();
    rangeToClearValidations.clearDataValidations();

    Sheet_Archiv.getRange(1, 1, Sheet_Archiv.getLastRow(), Sheet_Archiv.getLastColumn())
      .setValues(Sheet_ProjektPlanung.getRange(1, 1, Sheet_ProjektPlanung.getLastRow(), Sheet_ProjektPlanung.getLastColumn()).getValues());

    for (var i = 0; i < dataValidations.length; i++) {
      for (var j = 0; j < dataValidations[i].length; j++) {
        if (dataValidations[i][j] != null) {
          rangeToClearValidations.getCell(i + 1, j + 1).setDataValidation(dataValidations[i][j]);
        }
      }
    }

    var Sheet_Uebersicht = S_Export.getSheetByName("Übersicht");
    var sheetName = Sheet_ProjektPlanung.getName();
    var Link = S_Export.getUrl() + "#gid=" + Sheet_Archiv.getSheetId();

    Sheet_Uebersicht.getRange("B" + (Sheet_Uebersicht.getLastRow() + 1) + ":C" + (Sheet_Uebersicht.getLastRow() + 1)).setValues([[Datum,sheetName]])
    Sheet_Uebersicht.getRange("D" + Sheet_Uebersicht.getLastRow()).setFormula("=HYPERLINK(\""+Link+"\";\"Link\")");

    try
    {
    S.deleteSheet(Sheet_ProjektPlanung);
    } 
    catch (err) 
    {
    Logger.log('Fehler beim Löschen des Blatts: ' + err);
    }
  }
}

