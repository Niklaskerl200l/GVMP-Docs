function LT(Sergeant ="test",Pruefer1,Zeit,Versuche = 1,Pruefer2)   //LT Blatt erstellen
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_LT_Vorlage = SpreadsheetApp.getActive().getSheetByName("Prüfung Vorlage");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  Sheet_LT_Vorlage.copyTo(SS).setName("LT Prüfung " + Versuche + " " + Sergeant);

  var Sheet = SpreadsheetApp.getActive().getSheetByName("LT Prüfung " + Versuche + " " + Sergeant);

  Sheet.setActiveSelection("A1");

//-------------------------------Kopfdaten----------------------------------------//
  Sheet.getRange("B2").setValue("Lieutenant Prüfung "+ Versuche + " " + Sergeant)
  Sheet.getRange("F4").setValue(Sergeant);
  Sheet.getRange("F5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy"));
  Sheet.getRange("E5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"HH:mm"));
  Sheet.getRange("I5").setValue(Versuche);
  Sheet.getRange("P4").setValue(Pruefer1);
  Sheet.getRange("P5").setValue(Pruefer2);

//-------------------------------Theorie Fragen--------------------------------------//

  var Array_Fragen = Zufallsfragen();
  var Summe_Punkte = 0;

  var Anzahl_Fragen = Sheet_Auswertung.getRange("C6").getValue();

  for(var x = 1; x <= Math.round(Anzahl_Fragen / 2); x++)   //Spalte B Fragen
  {
    var Array_Punkte = [];

    for(var i = 0; i <= Array_Fragen[x - 1][2]; i++)
    {
      Array_Punkte.push(i);
    }

    Summe_Punkte += Array_Fragen[x - 1][2];

    var Datenvalidierung = SpreadsheetApp.newDataValidation().requireValueInList(Array_Punkte).build();

    Sheet.getRange("B" + (x * 3 + 5)).setValue(Array_Fragen[x - 1][0]);    
    Sheet.getRange("H" + (x * 3 + 5)).setDataValidation(Datenvalidierung);
    Sheet.getRange("B" + (x * 3 + 6)).setNote(Array_Fragen[x - 1][1]);
  }

  for(var x = Math.round(Anzahl_Fragen / 2) + 1; x <= Anzahl_Fragen; x++)   //Spalte H Fragen
  {
    var Array_Punkte = [];

    for(var i = 0; i <= Array_Fragen[x - 1][2]; i++)
    {
      Array_Punkte.push(i);
    }

    Summe_Punkte += Array_Fragen[x - 1][2];

    var Datenvalidierung = SpreadsheetApp.newDataValidation().requireValueInList(Array_Punkte).build();

    Sheet.getRange("L" + ((x - Math.round(Anzahl_Fragen / 2)) * 3 + 5)).setValue(Array_Fragen[x - 1][0]);
    Sheet.getRange("K" + ((x - Math.round(Anzahl_Fragen / 2)) * 3 + 5)).setDataValidation(Datenvalidierung);
    Sheet.getRange("L" + ((x - Math.round(Anzahl_Fragen / 2)) * 3 + 6)).setNote(Array_Fragen[x - 1][1]);
  }

  Sheet.getRange("P7").setValue(Summe_Punkte);
}

function LT_Archivieren()    //LT Archivieren
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_LT = SpreadsheetApp.getActiveSheet();
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Lieutenant_Prüfung);

  if(Sheet_LT.getName() == "Prüfung Vorlage") return;

  var Sergeant = Sheet_LT.getRange("F4").getValue();
  var Datum = Sheet_LT.getRange("F5").getValue();
  var Versuch = Sheet_LT.getRange("I5").getValue();
  var Pruefer1 = Sheet_LT.getRange("P4").getValue();
  var Pruefer2 = Sheet_LT.getRange("P5").getValue();
  var Status = Sheet_LT.getRange("Q34").getValue();

  var Zeile_Export = Sheet_Export.getSheetByName("Übersicht").getRange("B1").getValue();

  Sheet_LT.getRange("A1:L").setValues(Sheet_LT.getRange("A1:L").getValues());
  Sheet_LT.getRange("A1:R").clearDataValidations();

  Sheet_LT.copyTo(Sheet_Export).setName(Sheet_LT.getName());

  var URL = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_LT.getName()).getSheetId();

  Sheet_Export = Sheet_Export.getSheetByName("Übersicht");

  Sheet_Export.getRange("B" + Zeile_Export + ":G" + Zeile_Export).setValues([[Sergeant,Datum,Versuch,Pruefer1,Pruefer2,Status]]);

  Sheet_Export.getRange("H" + Zeile_Export).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  SS.deleteSheet(Sheet_LT);
}

