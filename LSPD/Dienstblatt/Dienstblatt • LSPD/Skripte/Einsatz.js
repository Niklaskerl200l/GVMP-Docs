function Szenario(e)
{
  var Sheet = e.source.getActiveSheet();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("C") && Zeile == 11 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","EL",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 12 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","VHL",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 13 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","Taktik",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 14 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","Backstab",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 15 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","Air Support",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 16 && Value == "TRUE")
  {
    Aufgabenverteilung("Login","Abtransport",e);
    Sheet.setActiveSelection("A1");
  }


  else if(Spalte == Spalte_in_Index("D") && Zeile == 11 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","EL",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 12 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","VHL",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 13 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","Taktik",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 14 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","Backstab",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 15 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","Air Support",e);
    Sheet.setActiveSelection("A1");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 16 && Value == "TRUE")
  {
    Aufgabenverteilung("Logout","Abtransport",e);
    Sheet.setActiveSelection("A1");
  }


  else if(Spalte == Spalte_in_Index("I") && Zeile == 8 && Value == "TRUE")
  {
    var Sheet_Szenario = SpreadsheetApp.getActive().getSheetByName("Einsatz");

    Einsatz_Clear();

    Sheet_Szenario.setActiveSelection("A1");
  }

  else
  {
    Sheet.getRange("K4").setValue("Sammelakte - " + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"HH:mm") + " Uhr");
  }
}

function Einsatz_Clear()
{
  var Sheet_Szenario = SpreadsheetApp.getActive().getSheetByName("Einsatz");

  Sheet_Szenario.getRange("B5:C5").setValue("");
  Sheet_Szenario.getRange("G5:I5").setValue("");
  Sheet_Szenario.getRange("B8:G8").setValue("");
  Sheet_Szenario.getRange("I8").setValue("");
  Sheet_Szenario.getRange("K5:K9").setValue("");
  Sheet_Szenario.getRange("C11:D16").setValue("");
  Sheet_Szenario.getRange("F12:I29").setValue("");
  Sheet_Szenario.getRange("B19:D23").setValue("");
  Sheet_Szenario.getRange("B26:D29").setValue("");
  Sheet_Szenario.getRange("K4").setValue("Sammelakte");
}

function Aufgabenverteilung(Aktion = "Logout",Aufgabe = "EL",e)
{
  var Lock = LockService.getDocumentLock();
  
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

  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();

  var Sheet_Szenario = SpreadsheetApp.getActive().getSheetByName("Einsatz");

  var Array_Aufgaben = Sheet_Szenario.getRange("F12:I29").getValues();
  

  var Array_Beamter = Umwandeln();
  var Gefunden = false;
  var Letzte_Zeile = 0;

  for(var i = 0; i < Array_Aufgaben.length; i++)
  {
    if(Array_Aufgaben[i][1] == Array_Beamter[2])
    {
      if(Aktion == "Login" && Aufgabe == Array_Aufgaben[i][3])
      {
        Logger.log("Ist bereits auf dem Einsatz");

        Gefunden = true;
        break;
      }
      else if(Aktion == "Logout" && Aufgabe == Array_Aufgaben[i][3])
      {
        Logger.log("Logout " + Array_Beamter[2] + " von " + Aufgabe);

        Array_Aufgaben[i] = ["","","",""];

        Array_Aufgaben.sort(Sort_Aufgabe);

        Sheet_Szenario.getRange("F12:I29").setValues(Array_Aufgaben);

        break;
      }
    }

    if(Array_Aufgaben[i][1] != "")
    {
      Letzte_Zeile = i;
    }
  }

  if(Aktion == "Login" && Gefunden == false)
  {
    if(Letzte_Zeile+1 > Array_Aufgaben.length)
    {
      Logger.log("Voll")

      SpreadsheetApp.getActiveSpreadsheet().toast("Voll");
    }
    else
    {
      Logger.log("Login " + Array_Beamter[2] + " auf " + Aufgabe);

      Array_Aufgaben[Letzte_Zeile+1] = [Array_Beamter[0],Array_Beamter[2],Array_Beamter[4],Aufgabe];
      Array_Aufgaben.sort(Sort_Aufgabe);

      Sheet_Szenario.getRange("F12:I29").setValues(Array_Aufgaben);

      var Sheet_Log_Szenario = SpreadsheetApp.getActive().getSheetByName("Log Szenario");
      var Einsatz = Sheet_Szenario.getRange("B5").getValue();
      Sheet_Log_Szenario.appendRow(["",Array_Beamter[2],Einsatz,Aufgabe,new Date()]);
    }
  }
  
  Sheet_Szenario.getRange(Zeile,Spalte).setValue(false);

  SpreadsheetApp.flush();
  Lock.releaseLock();
}

function Sort_Aufgabe(a,b)
{
  var Rang = 0;
  var Name = 1;
  var Aufgabe = 3;

  Logger.log(a[Aufgabe] == "Abtransport")

  if(a[Aufgabe] == "EL")
  {
    if(b[Aufgabe] != "EL")      return -1;
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "VHL")
  {
    if(b[Aufgabe] != "VHL")
    {
      if(b[Aufgabe] == "EL")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "Taktik")
  {
    if(b[Aufgabe] != "Taktik")
    {
      if(b[Aufgabe] == "EL" || b[Aufgabe] == "VHL")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "Backstab")
  {
    if(b[Aufgabe] != "Backstab")
    {
      if(b[Aufgabe] == "EL" || b[Aufgabe] == "VHL" || b[Aufgabe] == "Taktik")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "Caller")
  {
    if(b[Aufgabe] != "Caller")
    {
      if(b[Aufgabe] == "EL" || b[Aufgabe] == "VHL" || b[Aufgabe] == "Taktik" || b[Aufgabe] == "Backstab")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "Air Support")
  {
    if(b[Aufgabe] != "Air Support")
    {
      if(b[Aufgabe] == "EL" || b[Aufgabe] == "VHL" || b[Aufgabe] == "Taktik" || b[Aufgabe] == "Backstab" || b[Aufgabe] == "Caller")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  else if(a[Aufgabe] == "Abtransport")
  {
    if(b[Aufgabe] != "Abtransport")
    {
      if(b[Aufgabe] == "EL" || b[Aufgabe] == "VHL" || b[Aufgabe] == "Taktik" || b[Aufgabe] == "Backstab" || b[Aufgabe] == "Caller" || b[Aufgabe] == "Air Support")    return 1;
      return -1;
    }
    else if(a[Rang] > b[Rang])  return -1;
    else if(a[Rang] < b[Rang])  return 1;
    else if(a[Rang] == b[Rang])
    {
      if(a[Name].toLowerCase() < b[Name].toLowerCase())       return -1;
      else if(a[Name].toLowerCase() > b[Name].toLowerCase())  return 1;
    }
  }

  return 0;
}