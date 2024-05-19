function EST(RCT,Pruefer1,Zeit,Versuche,Pruefer2)   //EST Blatt erstellen
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_EST_Vorlage = SpreadsheetApp.getActive().getSheetByName("EST Vorlage");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Sheet_Personaltabelle = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
  Sheet_EST_Vorlage.copyTo(SS).setName("EST " + Versuche + " " + RCT);

  var Sheet = SpreadsheetApp.getActive().getSheetByName("EST " + Versuche + " " + RCT);

  Sheet.setActiveSelection("A1");



//-------------------------------Kopfdaten----------------------------------------//
  Sheet.getRange("B2").setValue("Einstellungstest "+ Versuche + " " + RCT)
  Sheet.getRange("D4").setValue(RCT);
  Sheet.getRange("D5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy"));
  Sheet.getRange("C5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"HH:mm"));
  Sheet.getRange("F5").setValue(Versuche);
  Sheet.getRange("H4").setValue(Pruefer1);
  Sheet.getRange("H5").setValue(Pruefer2);

//-------------------------------Theorie Fragen--------------------------------------//

  var Array_Fragen = Zufallsfragen();

  var Anzahl_Fragen = Sheet_Auswertung.getRange("C6").getValue();

  var Array_Personal = Sheet_Personaltabelle.getRange("D4:E199").getValues();

  var DN = 0;

  for(var i = 0; i <= Array_Personal.length; i++)
  {
    if(Array_Personal[i][0] == RCT)
    {
      var DN = Array_Personal[i][1];
      break;
    }
  }


  for(var x = 1; x <= Math.round(Anzahl_Fragen / 2); x++)   //Spalte B Fragen
  {
    Sheet.getRange("B" + (x * 3 + 5)).setValue(Array_Fragen[x - 1][0]);
    Sheet.getRange("B" + (x * 3 + 6)).setNote(Array_Fragen[x - 1][1]);
  }

  Sheet.getRange("B" + ((Math.round(Anzahl_Fragen / 2) * 3 + 5) + 4)).setNote(DN)

  for(var x = Math.round(Anzahl_Fragen / 2) + 1; x <= Anzahl_Fragen; x++)   //Spalte H Fragen
  {
    Sheet.getRange("H" + ((x - Math.round(Anzahl_Fragen / 2)) * 3 + 5)).setValue(Array_Fragen[x - 1][0]);
    Sheet.getRange("H" + ((x - Math.round(Anzahl_Fragen / 2)) * 3 + 6)).setNote(Array_Fragen[x - 1][1]);
  }

  
//-----------------------------------Funk---------------------------------------------//

  var Array_Funk = Zufallsfunk();
  var Anzahl_Funk = Sheet_Auswertung.getRange("C11").getValue();
  var Zeile_Funk = Sheet_Auswertung.getRange("L11").getValue();

  for(var x = 0; x < Math.round(Anzahl_Funk / 2); x++)    //Spalte B Funk
  {
    Sheet.getRange("B" + (x * 5 + Zeile_Funk)).setValue(Array_Funk[x][0]);
    Sheet.getRange("B" + (x * 5 + Zeile_Funk + 1)).setNote(Array_Funk[x][1]);
  }

  for(var x = Math.round(Anzahl_Funk / 2); x < Anzahl_Funk; x++)    //Spalte H Funk
  {
    Sheet.getRange("H" + ((x - Math.round(Anzahl_Funk / 2)) * 5 + Zeile_Funk)).setValue(Array_Funk[x][0]);
    Sheet.getRange("H" + ((x - Math.round(Anzahl_Funk / 2)) * 5 + Zeile_Funk + 1)).setNote(Array_Funk[x][1]);
  }

//------------------------------------Orte--------------------------------------------//

  var Array_Orte = Zufallsort();
  var Anzahl_Orte = Sheet_Auswertung.getRange("C12").getValue();
  var Zeile_Orte = Sheet_Auswertung.getRange("L12").getValue();

  for(var x = 0; x < Math.round(Anzahl_Orte / 2); x++)    //Spalte B Orte
  {
    Sheet.getRange("B" + (x * 5 + Zeile_Orte)).setValue(Array_Orte[x]);
  }

  for(var x = Math.round(Anzahl_Orte / 2); x < Anzahl_Orte; x++)    //Spalte H Orte
  {
    Sheet.getRange("H" + ((x - Math.round(Anzahl_Orte / 2)) * 5 + Zeile_Orte)).setValue(Array_Orte[x]);
  }
//----------------------------------Alert To-Say--------------------------------------//

SpreadsheetApp.getUi().alert('Bitte den Prüfling darauf hinweisen :\n- Kein 10-31 (Drohne)\n- Keine Dashcam aufnahmen\nAbsolute geheimhaltung\nBei Betrugsverdacht kontrollen möglich\n\nmögliche folgen: Kündigung / Schweres Dienstvergehen etc.\n\nAblauf:\nERST THEORIE\nDANN ERST PRAXIS (RTS; ORTSKUNDE; FUNK..)\nDas Tempomat ist in der Praxis NICHT erlaubt!\n\nEs wird KEINER mitgenommen, der nicht in der Abteilung ist!\nEST werden nur von der Leitung Archiviert!\n\nEST die in der Zeit vin 00:00 - 16:00 Uhr Stattfinden, sollen bitte von der Leitung der Abteilung genehmigen lassen!');

//------------------------------------------------------------------------------------//
}


//--------------------------------EST Archivieren--------------------------------------//
function EST_Archivieren()    
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

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Sheet_EST = SpreadsheetApp.getActiveSheet();
  var Sheet_Export = SpreadsheetApp.openById("1kevqeCbYzZ-hXKOJSFMAjLUqa4IIF-UrFSIz_mdqLwU");

  if(Sheet_EST.getName() == "EST Vorlage") return;

  var RCT = Sheet_EST.getRange("D4").getValue();
  var Datum = Sheet_EST.getRange("D5").getValue();
  var Versuch = Sheet_EST.getRange("F5").getValue();
  var Pruefer1 = Sheet_EST.getRange("H4").getValue();
  var Pruefer2 = Sheet_EST.getRange("H5").getValue();
  var Punkte = Sheet_EST.getRange("I" + Sheet_Auswertung.getRange("L10").getValue()).getValue();
  var Status = Sheet_EST.getRange("K" + Sheet_Auswertung.getRange("L10").getValue()).getValue();

  var Zeile_Export = Sheet_Export.getLastRow() + 1;

  Sheet_EST.getRange("A1:L").setValues(Sheet_EST.getRange("A1:L").getValues());
  Sheet_EST.getRange("A1:L").clearDataValidations();

  Sheet_EST.copyTo(Sheet_Export).setName(Sheet_EST.getName());

  var URL = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_EST.getName()).getSheetId();

  Sheet_Export = Sheet_Export.getSheetByName("Übersicht");

  Sheet_Export.getRange("B" + Zeile_Export + ":H" + Zeile_Export).setValues([[RCT,Datum,Versuch,Pruefer1,Pruefer2,Punkte,Status]]);

  Sheet_Export.getRange("I" + Zeile_Export).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  SS.deleteSheet(Sheet_EST);

  SpreadsheetApp.flush();
  Lock.releaseLock();
}