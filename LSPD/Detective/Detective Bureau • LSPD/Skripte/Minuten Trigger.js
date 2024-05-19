function Minuten_Trigger() 
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);

  if(Stunde == 0 && Minute == 0)
  {
    var Sheet_Dienstblatt = SpreadsheetApp.openById(LSPD.ID_Dienstblatt).getSheetByName("Startseite");
    var Sheet_Statistiken = SpreadsheetApp.getActive().getSheetByName("Export Statistiken");

    var Array_Statistiken = Sheet_Statistiken.getRange("C3:C12").getValues();

    Logger.log("Update Statistiken\n" + Array_Statistiken);

    Sheet_Dienstblatt.getRange("I38:I47").setValues(Array_Statistiken);

    if(Datum.getDay() == 0)
    {
      var SS_Geschwindigkeitstickets = SpreadsheetApp.openById(LSPD.ID_Geschwindigkeits_Tickets);
      var Sheet_Geschwindigkeitstickets = SS_Geschwindigkeitstickets.getSheetByName("Export Tickets");
      var Array_Geschwindigkeitstickets = Sheet_Geschwindigkeitstickets.getRange("B3:J1000").getValues();

      var Sheet_Personal = SS_Geschwindigkeitstickets.getSheetByName("Import Personaltabelle");
      var Array_Personal = Sheet_Personal.getRange("D4:D199").getValues();

      var Array_Meldung = [];
    }
  }

  if(Datum.getMinutes() % 10 == 0)
  {
    Personalvermerke_Aktualisieren();

    var Sheet_Anwesenheitskontrolle = SpreadsheetApp.getActive().getSheetByName("Anwesenheitskontrolle");

    Sheet_Anwesenheitskontrolle.getRange("E5").clearContent();
    SpreadsheetApp.flush();

    Sheet_Anwesenheitskontrolle.getRange("E5").setFormula(`=${Sheet_Anwesenheitskontrolle.getRange("E1").getValue()}`);
    SpreadsheetApp.flush();
  }

  LST_Plus_Zeitsystem();
  LST_Minus_Zeitsystem();
}

function test4()
{
  var SS_Geschwindigkeitstickets = SpreadsheetApp.openById(LSPD.ID_Geschwindigkeits_Tickets);
  var Sheet_Geschwindigkeitstickets = SS_Geschwindigkeitstickets.getSheetByName("Export Tickets");
  var Array_Geschwindigkeitstickets = Sheet_Geschwindigkeitstickets.getRange("B3:J1000").getValues();

  var Sheet_Personal = SS_Geschwindigkeitstickets.getSheetByName("Import Personaltabelle");
  var Array_Personal = Sheet_Personal.getRange("D4:D199").getValues();

  var Array_Meldung = [];

  var Grenze = 5;
  var Grenzdatum = new Date();
  Grenzdatum.setDate(Grenzdatum.getDate() - 7);

  for(var i = 0; i < Array_Personal.length; i++)
  {
    if(Array_Personal[i][0] != "")
    {
      var Insgesamt = 0;

      for(var o = 0; o < Array_Geschwindigkeitstickets.length; o++)
      {
        if(Array_Geschwindigkeitstickets[o][0] == Array_Personal[i][0] && Array_Geschwindigkeitstickets[o][5] != "Einsatz" && Array_Geschwindigkeitstickets[o][8] >= Grenzdatum)
        {
        }
      }
    }
  }
}