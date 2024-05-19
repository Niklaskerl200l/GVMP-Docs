function Leitstelle(e)
{
  var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("D") && Zeile == 4 && Value != undefined && Value != "Suche")
  {
    if(Sheet_Startseite.getRange("C" + Zeile).getValue() != "📢 Mobil")
    {
      Stempeluhr(0, Value);

      SpreadsheetApp.getActive().getSheetByName("Stempeluhr").getRange("O3").setValue(new Date());
    }
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 5 && Value != undefined && Value != "Suche")
  {
    if(Sheet_Startseite.getRange("C" + Zeile).getValue() != "- 🔒 -")
    {
      Stempeluhr(0, Value);

      SpreadsheetApp.getActive().getSheetByName("Stempeluhr").getRange("O4").setValue(new Date());
    }
    else
    {
      
      Sheet_Startseite.getRange(Zeile, Spalte).clearContent();
      SpreadsheetApp.flush();

      SpreadsheetApp.getUi().alert("Fehler!\n\nDie 2. Leitstelle ist nicht freigegeben...");
    }
  }
}

function LST_Suche_Neu(Ausloeser = LSPD.Umwandeln())
{
  try
  {
    var SS_LSPD = SpreadsheetApp.getActive();

    var Sheet_Stempeluhr = SS_LSPD.getSheetByName("Stempeluhr");
    var Sheet_Personal = SS_LSPD.getSheetByName("Personaltabelle");

    var Array_Stempeluhr = Sheet_Stempeluhr.getRange("B3:F202").getDisplayValues();
    var Array_Personal = Sheet_Personal.getRange("D4:J199").getValues();
    var Array_Anwesenheit = [];

    for(var i = 0; i < Array_Stempeluhr.length; i++)
    {
      if(Array_Stempeluhr[i][0] < 8 && Array_Stempeluhr[i][0] >= 2 && Array_Stempeluhr[i][2] != "" && Array_Stempeluhr[i][3] == "" && LSPD.Z.includes(Array_Stempeluhr[i][2]) == false)
      {
        Array_Anwesenheit.push([Array_Stempeluhr[i][2], Array_Stempeluhr[i][0]]);
      }
    }

    for(var i = 0; i < Array_Personal.length; i++)
    {
      for(var o = 0; o < Array_Anwesenheit.length; o++)
      {
        if(Array_Personal[i][0].toString().toUpperCase() == Array_Anwesenheit[o][0].toString().toUpperCase())
        {
          Array_Anwesenheit[o].push(Array_Personal[i][6]);
          break;
        }
      }
    }

    var Array_Anwesenheit_Temp = Array_Anwesenheit;

    Array_Anwesenheit = Array_Anwesenheit.filter(function(e){return e[2] < new Date(1899,11,30,4,0)});
    Logger.log(Array_Anwesenheit);

    if(Array_Anwesenheit.length > 0)
    {
      var Zufallszahl = Math.floor(Math.random() * (Array_Anwesenheit.length)) + 0;
      var Zeit = Array_Anwesenheit[Zufallszahl][2].getHours() + ":" + Array_Anwesenheit[Zufallszahl][2].getMinutes();

      SS_LSPD.getSheetByName("Log Leitstelle").appendRow(["", new Date(), Array_Anwesenheit[Zufallszahl][0], Ausloeser, ""]);
      SpreadsheetApp.flush();

      return SpreadsheetApp.getUi().alert(`Neue Leitstelle:\nName: ${Array_Anwesenheit[Zufallszahl][0]}\nZeit: ${Zeit}`);
    }
    else
    {
      Array_Anwesenheit = Array_Anwesenheit_Temp;

      var Array_Minimum = Array_Anwesenheit[0];
      var Array_Minimum_Alle = [];

      Logger.log(Array_Anwesenheit);

      for(var i = 0; i < Array_Anwesenheit.length; i++)
      {
        if(Array_Anwesenheit[i][2] < Array_Minimum[2])
        {
          Array_Minimum = Array_Anwesenheit[i];
        }
      }

      for(var i = 0; i < Array_Anwesenheit.length; i++)
      {
        if(Array_Anwesenheit[i][2].toString() == Array_Minimum[2].toString())
        {
          Array_Minimum_Alle.push(Array_Anwesenheit[i]);
        }
      }

      Logger.log(Array_Minimum);

      if(Array_Minimum_Alle.length > 1)
      {
        var Zufallszahl = Math.floor(Math.random() * (Array_Minimum_Alle.length)) + 0;
        var Zeit = Array_Minimum_Alle[Zufallszahl][2].getHours() + ":" + Array_Minimum_Alle[Zufallszahl][2].getMinutes();

        SS_LSPD.getSheetByName("Log Leitstelle").appendRow(["", new Date(), Array_Minimum_Alle[Zufallszahl][0], Ausloeser, ""]);
        SpreadsheetApp.flush();

        return SpreadsheetApp.getUi().alert(`Neue Leitstelle (alle haben ihre 4 Stunden voll):\nName: ${Array_Minimum_Alle[Zufallszahl][0]}\nZeit: ${Zeit}`);
      }
      else
      {
        var Zeit = Array_Minimum[2].getHours() + ":" + Array_Minimum[2].getMinutes();

        SS_LSPD.getSheetByName("Log Leitstelle").appendRow(["", new Date(), Array_Minimum[0], Ausloeser, ""]);
        SpreadsheetApp.flush();

        return SpreadsheetApp.getUi().alert(`Neue Leitstelle (alle haben ihre 4 Stunden voll):\nName: ${Array_Minimum[0]}\nZeit: ${Zeit}`);
      }
    }
  }
  catch(err)
  {
    Logger.log(err.stack);
    Fehler++;
  }
}

function Leitstelle_Suche_Auto_Archivieren()
{
  var Dienstblatt = SpreadsheetApp.getActive().getSheetByName("Log Leitstelle");
  var Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_LSTSuche).getSheetByName("Archiv");

  var Daten = Dienstblatt.getRange("B4:D").getValues();

  var Archiv_Array = [];

  if(Dienstblatt.getRange("B1").getValue() > 3)
  {
    for(var i = Daten.length - 1; i >= 0; i--)
    {
      if(Daten[i][0] != "")
      {
        Archiv_Array.push(Daten[i]);
      }
    }

    if(Archiv_Array.length > 0)
    {
      Archiv.insertRowsBefore(3, Archiv_Array.length);
      Archiv.getRange("B3:D" + (Archiv_Array.length + 2)).setValues(Archiv_Array);
      Dienstblatt.getRange("B3:D" + Dienstblatt.getLastRow()).setValue("");
    }
  }
}

function Leitstelle_Check_State(Grenzebeamte = 7)
{
  var Sheet_Stempeluhr = SpreadsheetApp.getActive().getSheetByName("Stempeluhr");
  var Array_Stempeluhr = Sheet_Stempeluhr.getRange("D3:E99").getValues().filter(function(e){return e[0] != "" && e[1] == ""});

  if(Array_Stempeluhr.length >= Grenzebeamte)
  {
    return Leitstelle_Mobil_Freigabe(false);
  }
  else
  {
    return Leitstelle_Mobil_Freigabe(true);
  }
}

function Leitstelle_Mobil_Freigabe(Status = false)
{
  SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns").getRange("E20").setValue(Status).insertCheckboxes();
}

function Leitstelle_Stempeluhr_Check()
{
  var SS_Dienstblatt = SpreadsheetApp.getActive();
  var Sheet_Auswertung = SS_Dienstblatt.getSheetByName("Auswertungsgedöns");
  var State_Leitstelle = Sheet_Auswertung.getRange("E20").getValue();

  var Sheet_Stempeluhr = SS_Dienstblatt.getSheetByName("Stempeluhr");
  var Array_Stempeluhr = Sheet_Stempeluhr.getRange("D3:D99").getDisplayValues();

  var Startseite = SS_Dienstblatt.getSheetByName("Startseite");
  var Leitstelle = Startseite.getRange("D4").getValue();

  var Gefunden = false;
  if(Leitstelle != "")
  {
    if(State_Leitstelle == false)
    {
      for(var i = 0; i < Array_Stempeluhr.length; i++)
      {
        if(Array_Stempeluhr[i][0] != "" && Array_Stempeluhr[i][0] == Leitstelle)
        {
          Gefunden = true;
          break;
        }
      }

      Array_Stempeluhr = Sheet_Stempeluhr.getRange("J3:J99").getDisplayValues();

      for(var i = 0; i < Array_Stempeluhr.length; i++)
      {
        if(Array_Stempeluhr[i][0] != "" && Array_Stempeluhr[i][0] == Leitstelle)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == true)
      {
        Logger.log("Leitstelle " + Leitstelle + " austragen...");

        Stempeluhr(0, Leitstelle);
      }
    }
    else if(State_Leitstelle == true)
    {
      for(var i = 0; i < Array_Stempeluhr.length; i++)
      {
        if(Array_Stempeluhr[i][0] != "" && Array_Stempeluhr[i][0] == Leitstelle)
        {
          Gefunden = true;
          break;
        }
      }

      Array_Stempeluhr = Sheet_Stempeluhr.getRange("J3:J99").getDisplayValues();

      for(var i = 0; i < Array_Stempeluhr.length; i++)
      {
        if(Array_Stempeluhr[i][0] != "" && Array_Stempeluhr[i][0] == Leitstelle)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == false)
      {
        Logger.log("Leitstelle " + Leitstelle + " eintragen...");
        
        Stempeluhr(1, Leitstelle);
      }
    }
  }
}