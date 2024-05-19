var Sheet_Bewerbungen = SpreadsheetApp.getActive().getSheetByName("Bewerbungen");

function Start_Ergebniss()
{
  Logger.log("Benutzer: " + LSPD.Umwandeln());
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Ergebnis').setHeight(550).setWidth(800).setSandboxMode(HtmlService.SandboxMode.IFRAME)," ");
}

function Start_Abstimmung()
{
  Logger.log("Benutzer: " + LSPD.Umwandeln());
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Index').setHeight(550).setWidth(800).setSandboxMode(HtmlService.SandboxMode.IFRAME)," ");
}

function Start_Voting()
{
  Logger.log("Benutzer: " + LSPD.Umwandeln());
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Voting').setHeight(550).setWidth(800).setSandboxMode(HtmlService.SandboxMode.IFRAME)," ");
}

function Start_Settings()
{
  Logger.log("Benutzer: " + LSPD.Umwandeln());
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Settings').setHeight(550).setWidth(1000).setSandboxMode(HtmlService.SandboxMode.IFRAME)," ");
}

function Get_User()
{
  return LSPD.Umwandeln();
}

function Get_Bewerber_Array()
{
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("B1").getValue();

  if(Letzte_Zeile < 3)
  {
    return 0;
  }
  else 
  {
    if(Letzte_Zeile >= 14)
    {
      Letzte_Zeile = 14;
    }

    var Array_Werte = Sheet_Bewerbungen.getRange("BD3:BG" + Letzte_Zeile).getValues();
    var Array_Aushabe = new Array();

    if(Array_Werte[0][0] == "")
    {
      return 0;
    }

    for(var y = 0; y < Array_Werte.length; y++)
    {
      if(Get_Bewerber_Status(Array_Werte[y][0],"Abstimmung") == true || Get_Bewerber_Status(Array_Werte[y][0],"Bemerkung") == true)
      {
        Array_Aushabe.push(Array_Werte[y]);

        var Voting = Get_Voting(Array_Werte[y][0]);

        switch(Voting)
        {
          case "Nein": Array_Aushabe[Array_Aushabe.length-1].push("‼️"); break;
          case "Dafür": Array_Aushabe[Array_Aushabe.length-1].push("✅"); break;
          case "Dagegen": Array_Aushabe[Array_Aushabe.length-1].push("❌"); break;
          case "Enthalten": Array_Aushabe[Array_Aushabe.length-1].push("➖"); break;
        }
      }
    }

    return Array_Aushabe;
  }

}

function Get_Bewerber_Bewertung(Name)
{
  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues()

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    if(Array_Bewerber[0][x] == Name)
    {
      var Letzte_Zeile = Sheet_Bewerbungen.getRange(1, (x + 5)).getValue();

      var Array_Bewertungen = Sheet_Bewerbungen.getRange(3,(x + 5),(Letzte_Zeile - 2),3).getValues();
      
      Logger.log(Array_Bewertungen);

      for(var i = 0; i < Array_Bewertungen.length; i++)
      {
        if(Array_Bewertungen[i][0] != "")
        {
          Array_Bewertungen[i][0] = Utilities.formatDate(new Date(Array_Bewertungen[i][0]),SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM HH:mm");
        }
      }
      return Array_Bewertungen;
    }
  }
  return 0;
}

function Voting_Setzen(Bewerber, Wahl)
{
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("BD1").getValue();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("BD3:BD" + Letzte_Zeile).getValues();

  var Beamter = LSPD.Umwandeln();

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0] == Bewerber)
    {
      var Zeile = y + 3;

      break;
    }
  }

  if(Zeile == undefined)
  {
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  var Anzahl_Spalte = Sheet_Bewerbungen.getRange("BC" + Zeile).getValue();

  if(Anzahl_Spalte != 0)
  {
    var Array_Bewertung = Sheet_Bewerbungen.getRange(Zeile,Spalte_in_Index("BI"),1,Anzahl_Spalte).getValues();

    var Array_Status = new Array();

    for(var x = 0; x < Array_Bewertung[0].length; x++)    // Abstimmungen aufsplitten
    {
      if(Array_Bewertung[0][x] != "")
      {
        Array_Status.push(Array_Bewertung[0][x].toString().split("#_#",2));
      }
    }

    for(var y = 0; y < Array_Status.length; y++)    // Prüfen ob Beamter schon Abgestimmt hat
    {
      if(Array_Status[y][0] == Beamter)
      {
        SpreadsheetApp.getUi().alert("Du hast bereits für " + Bewerber + " Abgestimmt.\nWahl: " + Array_Status[y][1]);
        return 0;
      }
    }
  }

  Sheet_Bewerbungen.getRange(Zeile,Spalte_in_Index("BI") + Anzahl_Spalte).setValue(Beamter + "#_#" + Wahl);   // Wahl setzen

  Start_Voting();
}

// ------------------------------------------------------------------------- //

function Get_Voting(Bewerber)
{
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("BD1").getValue();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("BD3:BD" + Letzte_Zeile).getValues();

  var Beamter = LSPD.Umwandeln();

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0] == Bewerber)
    {
      var Zeile = y + 3;

      break;
    }
  }

  if(Zeile == undefined)
  {
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  var Anzahl_Spalte = Sheet_Bewerbungen.getRange("BC" + Zeile).getValue();

  if(Anzahl_Spalte == 0)
  {
    return "Nein";
  }

  var Array_Bewertung = Sheet_Bewerbungen.getRange(Zeile,Spalte_in_Index("BI"),1,Anzahl_Spalte).getValues();

  var Array_Status = new Array();

  for(var x = 0; x < Array_Bewertung[0].length; x++)    // Abstimmungen aufsplitten
  {
    if(Array_Bewertung[0][x] != "")
    {
      Array_Status.push(Array_Bewertung[0][x].toString().split("#_#",2));
    }
  }

  for(var y = 0; y < Array_Status.length; y++)    // Prüfen ob Beamter schon Abgestimmt hat
  {
    if(Array_Status[y][0] == Beamter)
    {
      Logger.log(Array_Status[y][1]);
      return Array_Status[y][1];
    }
  }

  return "Nein";
}



function Settings_Laden(Ansicht = true)
{
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("BD1").getValue();

  var Array_Bewerber_Alle = Sheet_Bewerbungen.getRange("B3:B" + Letzte_Zeile).getValues();
  var Array_Bewerber_Abstimmung = Sheet_Bewerbungen.getRange("BD3:BG" + Letzte_Zeile).getValues();
  var Array_Freigabe_Bemerkung = Sheet_Bewerbungen.getRange("C3:C" + Letzte_Zeile).getValues();
  var Array_Freigabe_Abstimmung = Sheet_Bewerbungen.getRange("BH3:BH" + Letzte_Zeile).getValues();

  try
  {
    if(Sheet_Bewerbungen.getLastColumn() - Spalte_in_Index("BH") == 0)
    {
      var Array_Abstimmung = Sheet_Bewerbungen.getRange(3,Spalte_in_Index("BI"),(Letzte_Zeile - 2),Sheet_Bewerbungen.getLastColumn() - Spalte_in_Index("BH") + 1).getValues();
    }
    else
    {
      var Array_Abstimmung = Sheet_Bewerbungen.getRange(3,Spalte_in_Index("BI"),(Letzte_Zeile - 2),Sheet_Bewerbungen.getLastColumn() - Spalte_in_Index("BH")).getValues();
    }
    
  }
  catch(err)
  {
    var Array_Abstimmung = [[]];
  }

  var Gefunden = false;

  var Array_Ausgabe = new Array();
  var Array_Abstimmung_Namen = new Array();

  for(var i = 0; i < Array_Bewerber_Alle.length; i++)
  {
    Array_Ausgabe[i] = new Array();
    Gefunden = false;

    for(var y = 0; y < Array_Bewerber_Abstimmung.length; y++)
    {
      if(Array_Bewerber_Alle[i][0] == Array_Bewerber_Abstimmung[y][0] && Array_Bewerber_Abstimmung[y][0] != "")
      {
        Gefunden = true;

        Array_Ausgabe[i][0] = Array_Bewerber_Abstimmung[y][0];

        if(Ansicht == true)
        {
          var Dafuer = "";
          var Dagegen = "";
          var Enthalten = "";

          for(var x = 0; x < Array_Abstimmung[y].length; x++)
          {
            var Dummy = Array_Abstimmung[y][x].toString().split("#_#",2);

            switch(Dummy[1])
            {
              case "Dafür" : Dafuer = Dafuer + Dummy[0] + "#_#"; break;
              case "Dagegen" : Dagegen = Dagegen + Dummy[0] + "#_#"; break;
              case "Enthalten" : Enthalten = Enthalten + Dummy[0] + "#_#"; break;
            }
          }

          Array_Ausgabe[i][1] = Dafuer;
          Array_Ausgabe[i][2] = Dagegen;
          Array_Ausgabe[i][3] = Enthalten;
          Array_Ausgabe[i][4] = Array_Freigabe_Bemerkung[i][0];
          Array_Ausgabe[i][5] = Array_Freigabe_Abstimmung[i][0];
        }
        else if(Ansicht == false)
        {
          Array_Ausgabe[i][1] = Array_Bewerber_Abstimmung[y][1];
          Array_Ausgabe[i][2] = Array_Bewerber_Abstimmung[y][2];
          Array_Ausgabe[i][3] = Array_Bewerber_Abstimmung[y][3];
          Array_Ausgabe[i][4] = Array_Freigabe_Bemerkung[i][0];
          Array_Ausgabe[i][5] = Array_Freigabe_Abstimmung[i][0];
        }

        break;
      }
    }

    if(Gefunden == false)
    {
      Array_Ausgabe[i][0] = Array_Bewerber_Alle[i][0];
      Array_Ausgabe[i][1] = "";
      Array_Ausgabe[i][2] = "";
      Array_Ausgabe[i][3] = "";
      Array_Ausgabe[i][4] = false;
      Array_Ausgabe[i][5] = false;
    }
  }

  Logger.log(Array_Ausgabe.length);

  return Array_Ausgabe;
}

function Set_Bewerber_Status(Bewerber,Status,Art)
{  

  var Letzte_Zeile = Sheet_Bewerbungen.getRange("B1").getValue();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("B3:B" + Letzte_Zeile).getValues();

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0] == Bewerber)
    {
      var Zeile = y + 3;

      break;
    }
  }

  if(Zeile == undefined)
  {
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  if(Art == "Bemerkung")
  {
    Sheet_Bewerbungen.getRange("C" + Zeile).setValue(Status);
  }
  else if(Art == "Abstimmung")
  {
    Sheet_Bewerbungen.getRange("BH" + Zeile).setValue(Status);
  }
}

function Get_Bewerber_Status(Bewerber,Art)
{  
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("B1").getValue();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("B3:B" + Letzte_Zeile).getValues();

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0] == Bewerber)
    {
      var Zeile = y + 3;

      break;
    }
  }

  if(Zeile == undefined)
  {
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  if(Art == "Bemerkung")
  {
    return Sheet_Bewerbungen.getRange("C" + Zeile).getValue();
  }
  else if(Art == "Abstimmung")
  {
    return Sheet_Bewerbungen.getRange("BH" + Zeile).getValue();
  }

  return false;
}

function Set_Neu_Bewerber(Name)
{
  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues();
  var Array_Bewertungen = [];
  var Array_Name = [];

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    if(Array_Bewerber[0][x] != "")
    {
      Array_Bewertungen.push(Sheet_Bewerbungen.getRange(3,(x + 5),Sheet_Bewerbungen.getRange(1,x+5).getValue() - 2,3).getValues());
      Array_Name.push(Sheet_Bewerbungen.getRange(2,x + 5).getValue());
    }
  }

  Sheet_Bewerbungen.getRange("E3:BA" + Sheet_Bewerbungen.getLastRow()).setValue("");

  Sheet_Bewerbungen.getRange(3,2,Sheet_Bewerbungen.getLastRow() - 2, Sheet_Bewerbungen.getLastColumn() - 1).sort(2);

  var Letzte_Zeile = Sheet_Bewerbungen.getRange("B1").getValue() + 1;

  try
  {
    Sheet_Bewerbungen.getRange("B" + Letzte_Zeile + ":C" + Letzte_Zeile).setValues([[Name,false]]);

    Sheet_Bewerbungen.getRange("BH" + Letzte_Zeile).setValue(false);

    Sheet_Bewerbungen.getRange(Letzte_Zeile,Spalte_in_Index("BI"),1, Sheet_Bewerbungen.getLastColumn() + 1).setValue("");
  }
  catch(err)
  {
    Logger.log(err.stack);
  }

  SpreadsheetApp.flush();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues();

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    for(var i = 0; i < Array_Name.length; i++)
    {
      if(Array_Bewerber[0][x] == Array_Name[i])
      {
        Sheet_Bewerbungen.getRange(3,x + 5,Array_Bewertungen[i].length,3).setValues(Array_Bewertungen[i]);
        break;
      }
    }
  }
}

function Entferne_Bewerber(Name)
{
  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues();
  var Array_Bewerber2 = Sheet_Bewerbungen.getRange("B3:B" + Sheet_Bewerbungen.getRange("B1").getValue()).getValues();
  var Array_Bewertungen = [];
  var Array_Name = [];
  var Spalte;

  Logger.log("Name: " + Name + "\tBeamter: " + LSPD.Umwandeln());

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    if(Array_Bewerber[0][x] != "")
    {
      if(Array_Bewerber[0][x] == Name)
      {
        Spalte = x + 5;
      }

      Array_Bewertungen.push(Sheet_Bewerbungen.getRange(3,(x + 5),Sheet_Bewerbungen.getRange(1,x+5).getValue() - 2,3).getValues());
      Array_Name.push(Sheet_Bewerbungen.getRange(2,x + 5).getValue());
    }
  }

  for(var y = 0; y < Array_Bewerber2.length; y++)
  {
    if(Array_Bewerber2[y][0] == Name)
    {
      var Zeile = y + 3;
      break;
    }
  }

  if(Zeile == undefined || Spalte == undefined)
  {
    Logger.log(Name);
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  var Voting_Spalten = Sheet_Bewerbungen.getRange("BC" + Zeile).getValue();

  if(Voting_Spalten == 0)
  {
    Voting_Spalten = 1;
  }

  var Array_Abstimmung = Sheet_Bewerbungen.getRange("BE" + Zeile + ":BG" + Zeile).getValues();
  var Array_Bemerkungen = Sheet_Bewerbungen.getRange(3,Spalte,Sheet_Bewerbungen.getRange(1,Spalte).getValue() - 2,3).getValues();
  var Array_Voting = Sheet_Bewerbungen.getRange(Zeile,61,1,Voting_Spalten).getValues();

  Bewerber_Archivieren(Name,Array_Abstimmung,Array_Bemerkungen,Array_Voting);

  Sheet_Bewerbungen.getRange("E3:BA" + Sheet_Bewerbungen.getLastRow()).setValue("");

  Sheet_Bewerbungen.getRange("B" + Zeile + ":C" + Zeile).setValue("");

  Sheet_Bewerbungen.getRange("BH" + Zeile).setValue("");

  Sheet_Bewerbungen.getRange(Zeile,Spalte_in_Index("BI"),1, Sheet_Bewerbungen.getLastColumn() + 1).setValue("");

  Sheet_Bewerbungen.getRange(3,2,Sheet_Bewerbungen.getLastRow() - 2, Sheet_Bewerbungen.getLastColumn() - 1).sort(2);

  SpreadsheetApp.flush();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues();

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    for(var i = 0; i < Array_Name.length; i++)
    {
      if(Array_Bewerber[0][x] == Array_Name[i])
      {
        Sheet_Bewerbungen.getRange(3,x + 5,Array_Bewertungen[i].length,3).setValues(Array_Bewertungen[i]);
        break;
      }
    }
  }
}

function Set_Bewertung(Bewerber,Bemerkung)
{
  if(Bemerkung == "")
  {
    return 0;
  }

  var Array_Bewerber = Sheet_Bewerbungen.getRange("E2:BA2").getValues()

  for(var x = 0; x < Array_Bewerber[0].length; x++)
  {
    if(Array_Bewerber[0][x] == Bewerber)
    {
      var Letzte_Zeile = Sheet_Bewerbungen.getRange(1, (x + 5)).getValue();

      Logger.log(Sheet_Bewerbungen.getRange(Letzte_Zeile, (x + 5)).getValue());

      if(Sheet_Bewerbungen.getRange(Letzte_Zeile, (x + 5)).getValue() != "")
      {
        Letzte_Zeile++;
      }

      var Array_Bemerkung_Set = [[new Date(), Bemerkung, LSPD.Umwandeln()]];
      
      Sheet_Bewerbungen.getRange(Letzte_Zeile,x + 5,1,3).setValues(Array_Bemerkung_Set);

      Logger.log(Array_Bemerkung_Set);

      Sheet_Bewerbungen.getRange(3,x + 5, Sheet_Bewerbungen.getRange(1,x+5).getValue() - 2,3).sort(x+5);

      return 1;
    }
  }
  return 0;
}

function Get_Abstimmungen(Bewerber = "Gary Koenig")
{
  var Letzte_Zeile = Sheet_Bewerbungen.getRange("BD1").getValue();

  var Array_Bewerber = Sheet_Bewerbungen.getRange("BD3:BD" + Letzte_Zeile).getValues();

  var Beamter = LSPD.Umwandeln();

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0] == Bewerber)
    {
      var Zeile = y + 3;

      break;
    }
  }

  if(Zeile == undefined)
  {
    SpreadsheetApp.getUi().alert("Bewerber nicht Gefunden");
    return 0;
  }

  var Anzahl_Spalte = Sheet_Bewerbungen.getRange("BC" + Zeile).getValue();

  if(Anzahl_Spalte != 0)
  {
    var Array_Bewertung = Sheet_Bewerbungen.getRange(Zeile,Spalte_in_Index("BI"),1,Anzahl_Spalte).getValues();

    var Array_Status = new Array();
    var Array_Abstimmungen = new Array();

    for(var x = 0; x < Array_Bewertung[0].length; x++)    // Abstimmungen aufsplitten
    {
      if(Array_Bewertung[0][x] != "")
      {
        Array_Status.push(Array_Bewertung[0][x].toString().split("#_#",2));
      }
    }

    return Array_Status;
  }

  return 0;
}

function Bewerber_Archivieren(Bewerber,Array_Abstimmung,Array_Bemerkungen,Array_Voting)
{
  Logger.log(Array_Bemerkungen);
  var Sheet_Archiv = SpreadsheetApp.openById(LSPD.ID_Leitung_Detective).getSheetByName("Bewerber Archiv");

  var Abstimungen = "";

  for(var y = 0; y < Array_Bemerkungen.length; y++)
  {
    if(y == Array_Bemerkungen.length - 1)
    {
      Abstimungen = Abstimungen + Array_Bemerkungen[y][1] + " // " + Array_Bemerkungen[y][2];
    }
    else
    {
      Abstimungen = Abstimungen + Array_Bemerkungen[y][1] + " // " + Array_Bemerkungen[y][2] + "\n";
    }
  }

  var Dafuer = "";
  var Dagegen = "";
  var Enthalten = "";

  for(var x = 0; x < Array_Voting[0].length; x++)
  {
    var Dummy = Array_Voting[0][x].toString().split("#_#",2);

    if(x == Array_Voting[0].length - 1)
    {
      switch(Dummy[1])
      {
        case "Dafür" : Dafuer = Dafuer + Dummy[0]; break;
        case "Dagegen" : Dagegen = Dagegen + Dummy[0]; break;
        case "Enthalten" : Enthalten = Enthalten + Dummy[0]; break;
      }
    }
    else
    {
      switch(Dummy[1])
      {
        case "Dafür" : Dafuer = Dafuer + Dummy[0] + "\n"; break;
        case "Dagegen" : Dagegen = Dagegen + Dummy[0] + "\n"; break;
        case "Enthalten" : Enthalten = Enthalten + Dummy[0] + "\n"; break;
      }
    }
  }

  var Array_Archiv = [[
    Bewerber
    ,Array_Abstimmung[0][0]
    ,Array_Abstimmung[0][1]
    ,Array_Abstimmung[0][2]
    ,Abstimungen
    ,new Date()
    ,""
    ,Dafuer
    ,Dagegen
    ,Enthalten
  ]];

  if(Array_Archiv[0][1] != 0 || Array_Archiv[0][2] != 0 || Array_Archiv[0][3] != 0 || Array_Archiv[0][4] != " // ")
  {
    Sheet_Archiv.insertRowAfter(5);

    Sheet_Archiv.getRange("B6:K6").setValues(Array_Archiv);
  }
}

function Get_Bewerber_Link(Name)
{
  var Array_Bewerber_Neu = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Neu").getRange("B5:I35").getValues();

  for(var y = 0; y < Array_Bewerber_Neu.length; y++)
  {
    if(Array_Bewerber_Neu[y][0] == Name)
    {
      return Array_Bewerber_Neu[y][7];
    }
  }

  return ""
}

function Get_Settings_Zugriff()
{
  return Sheet_Bewerbungen.getRange("A3:A" + Sheet_Bewerbungen.getRange("A1").getValue()).getValues();
}

function Bewerber_onOpen()
{
  var Beamter = LSPD.Umwandeln();

  var Sheet_Bewerbung = SpreadsheetApp.getActive().getSheetByName("Bewerbungen");
  var Letzte_Zeile = Sheet_Bewerbung.getRange("A1").getValue();

  var Array_Ausblenden = Sheet_Bewerbung.getRange("A3:A" + Letzte_Zeile).getValues();
  var Array_Bewerber = Get_Bewerber_Array();

  for(var y = 0; y < Array_Ausblenden.length; y++)
  {
    if(Array_Ausblenden[y][0] == Beamter)
    {
      return 0;
    }
  }

  var Nachricht = "";

  for(var i = 0; i < Array_Bewerber.length; i++)
  {
    var Freigabe_Bemerkung = Get_Bewerber_Status(Array_Bewerber[i][0],"Bemerkung");
    var Freigabe_Abstimmung = Get_Bewerber_Status(Array_Bewerber[i][0],"Abstimmung");
    var Array_Bewertungen = Get_Bewerber_Bewertung(Array_Bewerber[i][0]);
    var Array_Abstimmungen = Get_Abstimmungen(Array_Bewerber[i][0]);

    if(Freigabe_Bemerkung == true && Freigabe_Abstimmung == false)
    {
      var Gefunden_Bemerkung = false;

      for(var i2 = 0; i2 < Array_Bewertungen.length; i2++)
      {
        if(Array_Bewertungen[i2][2] == Beamter)
        {
          Gefunden_Bemerkung = true;
          break;
        }
      }

      if(Gefunden_Bemerkung == false)
      {
        Nachricht += "Du hast noch nicht deine Meinung zu " + Array_Bewerber[i][0] + " abgegeben!\n"
      }
    }

    if(Freigabe_Abstimmung == true)
    {
      var Gefunden_Abstimmungen = false;

      for(var i2 = 0; i2 < Array_Abstimmungen.length; i2++)
      {
        if(Array_Abstimmungen[i2][0] == Beamter)
        {
          Gefunden_Abstimmungen = true;
          break;
        }
      }

      if(Gefunden_Abstimmungen == false)
      {
        Nachricht += "Du hast noch nicht für " + Array_Bewerber[i][0] + " Abgestimmt!\n"
      }
    }
  }

  if(Nachricht != "")
  {
    Logger.log(Nachricht);
    SpreadsheetApp.getUi().alert(Nachricht);
  }
}