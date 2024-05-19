function Minuten_Trigger()
{
  var Zeitstempel = new Date();
  var Zeit = Utilities.formatDate(Zeitstempel, "CET", "HH:mm");

  Logger.log(Zeit);

  Beschlagnahme_Zugehoerigkeit();

  if(Zeit == "07:57")
  {
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

    Sheet_Startseite.getRange("I4").setValue("3");
    Sheet_Startseite.getRange("E5").setValue("1000");
    Sheet_Startseite.getRange("D4:D5").clearContent();
    Sheet_Startseite.getRange("D25:J25").setValue("0");

    Logout_Alle();
    Einsatz_Clear();
  }
  else if(Zeit == "15:57")
  {
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

    Sheet_Startseite.getRange("I4").setValue("4");
    Sheet_Startseite.getRange("E5").setValue("1000");
    Sheet_Startseite.getRange("D4:D5").clearContent();
    Sheet_Startseite.getRange("D25:J25").setValue("0");

    Logout_Alle();
    Einsatz_Clear();
  }
  else if(Zeit == "18:00")
  {
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
    var Array_Ridealong = Sheet_Startseite.getRange("S61:W105").getValues();

    var Array_Ausgabe = [];

    for(i = 0; i < Array_Ridealong.length; i++)
    {
      if(Array_Ridealong[i][0] != "" && new Date(Array_Ridealong[i][4]) > Zeitstempel)
      {
        Array_Ausgabe.push(Array_Ridealong[i]);
      }
    }

    try
    {
      Sheet_Startseite.getRange("S61:W105").clearContent();
      Sheet_Startseite.getRange(61, Spalte_in_Index("S"), Array_Ausgabe.length, Array_Ausgabe[0].length);
    }
    catch(err)
    {
      Logger.log(err.stack);
    }
  }
  else if(Zeit == "18:30")
  {
    Leitstelle_Mobil_Freigabe(false);
  }
  else if(Zeit == "23:57")
  {
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

    Sheet_Startseite.getRange("I4").setValue("3");
    Sheet_Startseite.getRange("E5").setValue("1000");
    Sheet_Startseite.getRange("D4:D5").clearContent();
    Sheet_Startseite.getRange("D25:J25").setValue("0");

    Leitstelle_Mobil_Freigabe(true);

    Logout_Alle();
    Einsatz_Clear();
    

    Log_Archivieren();
    Einsatz_Archivieren();
    Auto_Entfernen_Minderung();
    Einsatz_Auto_Archivieren();

    Aktenklaerung_Verjaehrung();

    Fahndungen_Verjaehrung();
    Fahndungen_Archivierung();

    Mobile_Leitstelle_Tagestrigger();
  }
  else if(Zeit == "00:05")
  {
    if(Zeitstempel.getDate() == 1)
    {
      Mobile_Leitstelle_Monatstrigger();
    }
  }

  if(Zeitstempel.getMinutes() % 15 == 0)
  {
    Anwesenheit();
  }
  else if(Zeitstempel.getMinutes() % 5 == 0)
  {
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
    
    var Status = Sheet_Startseite.getRange("C8").getValue();
    var Defcon = Sheet_Startseite.getRange("I4").getValue();

    if(Status == "Dienstpflicht" || Defcon <= 2)
    {
      var Datum = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns").getRange("E8").getValue();
      Datum = Datum.setMinutes(Datum.getMinutes() + 10);
    }
  }

  if(Zeit >= "12:01" && Zeit <= "15:59")
  {
    Leitstelle_Check_State(20);
  }

  if(Zeit >= "16:01" && Zeit <= "18:29")
  {
    Leitstelle_Check_State();
  }

  Mobile_Leitstelle_Minutentrigger();

  Leitstelle_Stempeluhr_Check();
  Leitstelle_Suche_Auto_Archivieren();
}