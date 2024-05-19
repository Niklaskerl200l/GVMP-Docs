function Officer_Start()
{
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Prüfung UI').setWidth(1114).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME),"Prüfungs Daten eingeben");
}

function Officer(PBO,Pruefer1,Zeit,Versuche,Pruefer2)
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_EST_Vorlage = SpreadsheetApp.getActive().getSheetByName("Prüfung Vorlage");

  Sheet_EST_Vorlage.copyTo(SS).setName("Officer Prüfung " + Versuche + " " + PBO);

  var Sheet = SpreadsheetApp.getActive().getSheetByName("Officer Prüfung " + Versuche + " " + PBO);
  
  Sheet.setActiveSelection("A1");

  Sheet.getRange("F4").setValue(PBO);
  Sheet.getRange("F5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy"));
  Sheet.getRange("E5").setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"HH:mm"));
  Sheet.getRange("I5").setValue(Versuche);
  Sheet.getRange("P4").setValue(Pruefer1);
  Sheet.getRange("P5").setValue(Pruefer2);

  var Array_Fragen = Zufallsfragen(PBO);

  var Letze_Zeile_Links = "B39";
  var Letze_Zeile_Rechts = "L36";
  var Zeile = 9;
  var Spalte = ["B","L"];
  var Spalten_Counter = 0;
  var Frage = 0;

  while(true)
  {
    Sheet.getRange(Spalte[Spalten_Counter] + Zeile).setValue(Array_Fragen[Frage][0]);
    Sheet.getRange(Spalte[Spalten_Counter] + (Zeile + 1)).setNote(Array_Fragen[Frage][1]);

    if(Spalte[Spalten_Counter] + Zeile == Letze_Zeile_Links)
    {
      Zeile = 6;
      Spalten_Counter = 1;
    }
    else if (Spalte[Spalten_Counter] + Zeile == Letze_Zeile_Rechts)
    {
      break;
    }

    Zeile = Zeile + 3;
    Frage++;    
  }

  var Sheet_Import_Termine = SpreadsheetApp.getActive().getSheetByName("Import Termine");
  var Array_Termine = Sheet_Import_Termine.getRange("B2:C").getValues();

  try
  {
    for(var i = Array_Termine.length - 1; i => 0; i--)
    {
      if(Array_Termine[i][1] == PBO && Array_Termine[i][0] == "Theorie")
      {
        Sheet_Import_Termine.getRange("N" + (i + 2)).setValue(true);
      }
    }
  }
  catch(err)
  {
    Logger.log("Person hat keinen Termin gemacht");
    Logger.log(err.stack);
  }

  SpreadsheetApp.getUi().alert('Kein 10-31 sowie Dashcam\nEs wird nicht über die Fragen geredet.\nEs wird nur innerorts gefahren. Kein Highway oder Außerorts\nKein 11-90 oder 11-99 anfahren.\nDer Prüfling sowie Hauptprüfer sind nicht im Funk. Der 2 Prüfer ist im Funk um 11-90 oder 11-99 zu umfahren.\nDer Prüfling kann sein Handy auf anrufe Ablehnen oder stummschalten\nBei Kopfschmerzen schnellstmöglich Frage noch beantworten solange es geht wenn er einmal weg ist kann die frage nicht weiterbeantwortet werden\nWenn der Prüfling trinken oder toilette muss erst frage fertig beantworten dann rechts ranfahren');
}

function Officer_Praxis()
{
  var Sheet = SpreadsheetApp.getActiveSheet();

  if(Sheet.getSheetName() == "Prüfung Vorlage")
  {
    SpreadsheetApp.getUi().alert("HALT STOP!\nHier bleibt alles so wie's ist!");
    return 0;
  }

  Sheet.getRange("E45").setValue(new Date());
  Sheet.getRange("F45").setValue(new Date());
  Sheet.getRange("P44").setValue(LSPD.Umwandeln());


  var Orte = Zufallsort();
  Spalte = ["B","M"];
  Spalten_Counter = 0;
  var x = 0;
  var y = 0;

  while(true)
  {
    if(x * 2 == SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns").getRange("C14").getValue())
    {
      Spalten_Counter = 1;
      x = 0;
    }
    if (y == SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns").getRange("C14").getValue())
    {
      break;
    }

    Sheet.getRange(Spalte[Spalten_Counter] + (48 + x)).setValue(Orte[y])

    x++;
    y++;
  }

  Sheet.setActiveSelection("P45");

  var Sheet_Import_Termine = SpreadsheetApp.getActive().getSheetByName("Import Termine");
  var Array_Termine = Sheet_Import_Termine.getRange("B2:C").getValues();

  var Name = Sheet.getRange("F44").getValue();

  try
  {
    for(var i = Array_Termine.length - 1; i => 0; i--)
    {
      if(Array_Termine[i][1] == Name && Array_Termine[i][0] == "Praxis")
      {
        Sheet_Import_Termine.getRange("N" + (i + 2)).setValue(true);
      }
    }
  }
  catch(err)
  {
    Logger.log("Person hat keinen Termin gemacht");
    Logger.log(err.stack);
  }

  SpreadsheetApp.getUi().alert('Es finden mehrere Praxis Module statt.\nEs wird kein 11-90 oder 11-99 angefahren.')
}

function PBO_Alle()
{
  try
  {
    var Sheet = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
    var Array = Sheet.getRange("G3:G" + Sheet.getRange("G1").getValue()).getValues();
    
    Logger.log(Array);
    
    return Array;
  }
  catch(err)
  {
    Logger.log(err.stack);
  }
}

function Pruefer_Alle()
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Array = Sheet.getRange("I3:I" + Sheet.getRange("I1").getValue()).getValues();
  return Array;
}