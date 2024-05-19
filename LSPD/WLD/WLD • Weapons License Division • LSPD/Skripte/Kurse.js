function WLD(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("M") && Zeile >= 5 && Value == "Ja")
  {
    var Sheet_Kurs = SpreadsheetApp.getActive().getSheetByName(SheetName);
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Log");

    var Name = Sheet_Kurs.getRange("D" + Zeile).getValue();
    var Letzte_Zeile = Sheet_Log.getRange("B3").getValue() + 1;

    Sheet_Log.getRange("B" + Letzte_Zeile + ":D" + Letzte_Zeile).setValues([[Name,new Date(), LSPD.Umwandeln()]]);
  }
}


function Kurs_Erstellen()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var UI = SpreadsheetApp.getUi();

  var Sheet_Vorlage = SS.getSheetByName("WLD Vorlage");
  var Sheet_Startseite = SS.getSheetByName("Startseite");

  var Letzte_Zeile = Sheet_Startseite.getRange("O3").getValue() + 1;
  

  var Abfrage1 = UI.prompt("Eingabe Kurs","Gib das Datum für den neuen Kurs an im Format dd.MM.yyyy",UI.ButtonSet.OK_CANCEL)

  if(Abfrage1.getSelectedButton() == UI.Button.OK)
  {
    if(Abfrage1.getResponseText().match('(0[1-9]|[12][0-9]|3[01])[\.](0[1-9]|1[012])[\.](19|20)[0-9][0-9]') == null)
    {
      UI.alert(Abfrage1.getResponseText() + " ist nicht richtig");

      return 1;
    }

    var Abfrage2 = UI.prompt("Eingabe Kurs","Gib die Uhrzeit für den neuen Kurs an im Format HH:mm",UI.ButtonSet.OK_CANCEL)

    if(Abfrage2.getSelectedButton() == UI.Button.OK)
    {
      if(Abfrage2.getResponseText().match('([0-2][0-9])[\:]([0-5][0-9])') == null)
      {
        UI.alert(Abfrage2.getResponseText() + " ist nicht richtig");

        return 1;
      }

      var Datum = Abfrage1.getResponseText();
      var Uhrzeit = Abfrage2.getResponseText();

      var Sheet_Kurs = Sheet_Vorlage.copyTo(SS).setName("WLD " + Datum + " " + Uhrzeit).showSheet();

      Sheet_Kurs.getRange("B2").setValue("WLD " + Datum + " " + Uhrzeit + " Uhr");
      Sheet_Kurs.getRange("B1").setValue(Datum + " " + Uhrzeit);

      Sheet_Startseite.getRange("O" + Letzte_Zeile).setValue(Datum + " " + Uhrzeit);

      Sheet_Startseite.getRange("O6:O" + Letzte_Zeile).sort(Spalte_in_Index("O"));

      Sheet_Kurs.setActiveSelection("A1");
    }
  }
} 




function Kurse_Archivieren()
{
  if(SpreadsheetApp.getActive().getSheetName() != "WLD Vorlage")
  {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var SS_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_WLD_Kurse);

    var Sheet_Kurs = SpreadsheetApp.getActiveSheet();
    var Sheet_Uebersicht = SS_Export.getSheetByName("Übersicht");
    var Sheet_Startseite = SS.getSheetByName("Startseite");
    var Sheet_Log = SS.getSheetByName("Log");

    var Array_Kurse = Sheet_Startseite.getRange("O6:O20").getValues();
    var Array_Kurs_Personal = Sheet_Kurs.getRange("B5:B19").getValues();
    var Array_Teilnehmer = Sheet_Kurs.getRange("D5:M19").getValues();

    var Datum = Sheet_Kurs.getRange("B1").getValue();
    var Anzahl = Sheet_Kurs.getRange("D3").getValue();
    var Letzte_Zeile_Teilnehmer = Sheet_Log.getRange("F3").getValue() + 1
    var Letzte_Zeile_Personal = Sheet_Log.getRange("K3").getValue() + 1;


    var Sheet_Archiv = Sheet_Kurs.copyTo(SS_Export);

    Sheet_Archiv.setName(Sheet_Kurs.getName());    

    var Link = SS_Export.getUrl() + "#gid=" + Sheet_Archiv.getSheetId();

    Sheet_Archiv.getRange("A1:T30").setValues(Sheet_Kurs.getRange("A1:T30").getValues());

    Sheet_Archiv.getRange(1,1,Sheet_Archiv.getMaxRows(),Sheet_Archiv.getMaxColumns()).clearDataValidations();

    Sheet_Uebersicht.getRange("B" + (Sheet_Uebersicht.getLastRow() + 1) + ":C" + (Sheet_Uebersicht.getLastRow() + 1)).setValues([[Datum,Anzahl]])
    Sheet_Uebersicht.getRange("D" + Sheet_Uebersicht.getLastRow()).setFormula("=HYPERLINK(\""+Link+"\";\"Link\")");

    Sheet_Uebersicht.getRange("B5:D" + Sheet_Uebersicht.getLastRow()).sort({column: 2, ascending: false});

    for(var i = 0; i < Array_Kurse.length; i++)
    {
      if(Array_Kurse[i][0].toString() == Datum.toString())
      {
        Sheet_Startseite.getRange("O" + (i+6)).setValue("");
        Sheet_Startseite.getRange("O6:O20").sort(Spalte_in_Index("O"));
      }
    }


   
    for(var i = 0; i < Array_Kurs_Personal.length; i++)
    {
      if(Array_Kurs_Personal[i][0] != "")
      {
        Sheet_Log.getRange("K" + Letzte_Zeile_Personal + ":L" + Letzte_Zeile_Personal).setValues([[Array_Kurs_Personal[i][0], Datum]])
        Letzte_Zeile_Personal++;
      }
    }


    for(var i = 0; i < Array_Teilnehmer.length; i++)
    {
      if(Array_Teilnehmer[i][0] != "" && Array_Teilnehmer[i][2] == true)
      {
        Sheet_Log.getRange("F" + Letzte_Zeile_Teilnehmer + ":I" + Letzte_Zeile_Teilnehmer).setValues([[Array_Teilnehmer[i][0], Datum, Array_Teilnehmer[i][6], Array_Teilnehmer[i][9] ]]);
        Letzte_Zeile_Teilnehmer++;
      }
    }

    SS.deleteSheet(Sheet_Kurs);
  }
}

function Kurse_Loeschen()
{
  if(SpreadsheetApp.getActive().getSheetName() != "WLD Vorlage")
  {
    if(SpreadsheetApp.getUi().alert("Blatt Löschen?",SpreadsheetApp.getUi().ButtonSet.YES_NO) == SpreadsheetApp.getUi().Button.YES)
    {
      var SS = SpreadsheetApp.getActiveSpreadsheet();

      var Sheet_Startseite = SS.getSheetByName("Startseite");
      var Sheet_Kurs = SpreadsheetApp.getActiveSheet();

      var Array_Kurse = Sheet_Startseite.getRange("O6:O20").getValues();

      var Datum = Sheet_Kurs.getRange("B1").getValue();


      for(var i = 0; i < Array_Kurse.length; i++)
      {
        if(Array_Kurse[i][0].toString() == Datum.toString())
        {
          Sheet_Startseite.getRange("O" + (i+6)).setValue("");
          Sheet_Startseite.getRange("O6:O20").sort(Spalte_in_Index("O"));
        }
      }

      SS.deleteSheet(Sheet_Kurs);
    }
  }
}